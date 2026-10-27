use serde::Serialize;
use std::collections::HashSet;

#[repr(C)]
#[derive(Default)]
struct Bsd {
    pid: u32,
    ppid: u32,
    pgid: u32,
    status: u32,
    name: [u8; 16],
    flags: u32,
    uid: u32,
    gid: u32,
    ruid: u32,
    rgid: u32,
    svuid: u32,
    svgid: u32,
    reserved: u32,
}
#[repr(C)]
#[derive(Default)]
struct Task {
    virtual_size: u64,
    resident_size: u64,
    times: [u64; 4],
    counters: [i32; 12],
}
unsafe extern "C" {
    fn proc_listpids(kind: u32, info: u32, buffer: *mut libc::c_void, size: i32) -> i32;
    fn proc_pidinfo(pid: i32, flavor: i32, arg: u64, buffer: *mut libc::c_void, size: i32) -> i32;
    fn proc_pidpath(pid: i32, buffer: *mut libc::c_void, size: u32) -> i32;
}
fn info<T: Default>(pid: i32, flavor: i32) -> Option<T> {
    let mut value = T::default();
    let size = std::mem::size_of::<T>();
    (unsafe { proc_pidinfo(pid, flavor, 0, (&mut value as *mut T).cast(), size as i32) }
        == size as i32)
        .then_some(value)
}
// Optional XNU coalition query. It is absent from public SDK headers; require
// the documented OSS structure size and treat rejection as missing coverage.
// Never query WKWebView private selectors or count every WebKit process.
fn coalition(pid: i32) -> Option<u64> {
    info::<[u64; 5]>(pid, 20)
        .map(|ids| ids[0])
        .filter(|id| *id > 1)
}
fn path(pid: i32) -> Option<String> {
    let mut bytes = [0u8; 4096];
    if unsafe { proc_pidpath(pid, bytes.as_mut_ptr().cast(), bytes.len() as u32) } <= 0 {
        return None;
    }
    let end = bytes.iter().position(|byte| *byte == 0)?;
    Some(String::from_utf8_lossy(&bytes[..end]).into_owned())
}
#[derive(Clone)]
struct Candidate {
    pid: u32,
    parent: u32,
    coalition: Option<u64>,
    path: String,
}
fn owned(
    candidates: &[Candidate],
    shell: u32,
    shell_coalition: Option<u64>,
) -> Vec<(Candidate, &'static str)> {
    let mut tree = HashSet::from([shell]);
    loop {
        let before = tree.len();
        for candidate in candidates {
            if tree.contains(&candidate.parent) {
                tree.insert(candidate.pid);
            }
        }
        if before == tree.len() {
            break;
        }
    }
    candidates
        .iter()
        .filter_map(|candidate| {
            if tree.contains(&candidate.pid) {
                Some((candidate.clone(), "descendant"))
            } else if shell_coalition.is_some()
                && candidate.coalition == shell_coalition
                && candidate.path.contains("/WebKit.framework/")
            {
                Some((candidate.clone(), "coalition"))
            } else {
                None
            }
        })
        .collect()
}
#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Process {
    pid: u32,
    parent_pid: u32,
    role: &'static str,
    rss_bytes: u64,
    virtual_bytes: u64,
    ownership: &'static str,
}
#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Snapshot {
    sampled_at_ms: u64,
    shell_pid: u32,
    processes: Vec<Process>,
    rss_total_bytes: u64,
    renderer_bytes: Option<u64>,
    web_kit_attribution: &'static str,
    skipped_processes: usize,
}
fn role(pid: u32, shell: u32, path: &str) -> &'static str {
    if pid == shell {
        "shell"
    } else if path.contains("WebContent") {
        "web-content"
    } else if path.contains("WebKit.Networking") {
        "web-network"
    } else if path.contains("WebKit.GPU") {
        "web-gpu"
    } else if path.ends_with("/bin/superdoc") || path.ends_with("/bun") {
        // The contained JavaScript runtime: the coordinator, its tool servers
        // and the document engine all run from this executable.
        "bun"
    } else if path.ends_with("/node") {
        // Only a comparison build or a user's own tool server.
        "node"
    } else if path.ends_with("/platform-service") {
        "rust"
    } else if path.ends_with("/opencode") {
        "opencode"
    } else {
        "other"
    }
}
pub fn sample() -> Result<Snapshot, String> {
    sample_for(std::process::id())
}
/// Another running app's numbers (the main app's `--memory-pid`).
#[allow(dead_code)]
pub fn sample_app(pid: u32) -> Result<Snapshot, String> {
    if !path(pid as i32).is_some_and(|path| path.ends_with("/platform-desktop")) {
        return Err("validation: Memory target is not a Tauri studio process".into());
    }
    sample_for(pid)
}
fn sample_for(shell: u32) -> Result<Snapshot, String> {
    let group = coalition(shell as i32);
    let size = unsafe { proc_listpids(4, libc::getuid(), std::ptr::null_mut(), 0) };
    if size <= 0 {
        return Err("unavailable: Cannot enumerate application processes".into());
    }
    let mut pids = vec![0i32; (size as usize / 4 + 128).min(32768)];
    let count = unsafe {
        proc_listpids(
            4,
            libc::getuid(),
            pids.as_mut_ptr().cast(),
            (pids.len() * 4) as i32,
        )
    };
    if count <= 0 {
        return Err("unavailable: Cannot read application process list".into());
    }
    pids.truncate((count as usize / 4).min(pids.len()));
    let mut candidates = Vec::new();
    for pid in pids.into_iter().filter(|pid| *pid > 0) {
        if let Some(bsd) = info::<Bsd>(pid, 13) {
            if let Some(path) = path(pid) {
                let webkit = path.contains("/WebKit.framework/");
                candidates.push(Candidate {
                    pid: pid as u32,
                    parent: bsd.ppid,
                    coalition: if webkit { coalition(pid) } else { None },
                    path,
                });
            }
        }
    }
    let mut processes = Vec::new();
    let mut skipped = 0;
    for (candidate, ownership) in owned(&candidates, shell, group) {
        if let Some(task) = info::<Task>(candidate.pid as i32, 4) {
            processes.push(Process {
                pid: candidate.pid,
                parent_pid: candidate.parent,
                role: role(candidate.pid, shell, &candidate.path),
                rss_bytes: task.resident_size,
                virtual_bytes: task.virtual_size,
                ownership,
            });
        } else {
            skipped += 1
        }
    }
    if !processes.iter().any(|process| process.pid == shell) {
        return Err("unavailable: Shell memory could not be measured".into());
    }
    let content = processes
        .iter()
        .filter(|process| process.role == "web-content")
        .map(|process| process.rss_bytes)
        .sum::<u64>();
    let total = processes.iter().map(|process| process.rss_bytes).sum();
    processes.sort_by_key(|process| process.pid);
    Ok(Snapshot {
        sampled_at_ms: std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .map_err(|e| e.to_string())?
            .as_millis() as u64,
        shell_pid: shell,
        processes,
        rss_total_bytes: total,
        renderer_bytes: (content > 0).then_some(content),
        web_kit_attribution: if group.is_some() {
            "coalition"
        } else {
            "unavailable"
        },
        skipped_processes: skipped,
    })
}
#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn membership_excludes_other_apps_and_global_webkit_when_attribution_is_missing() {
        let candidate = |pid, parent, coalition, path: &str| Candidate {
            pid,
            parent,
            coalition,
            path: path.into(),
        };
        let rows = vec![
            candidate(10, 1, None, "/app/shell"),
            candidate(20, 10, None, "/app/node"),
            candidate(30, 20, None, "/app/platform-service"),
            candidate(40, 1, Some(42), "/System/WebKit.framework/WebContent"),
            candidate(50, 1, Some(99), "/System/WebKit.framework/WebContent"),
            candidate(60, 1, Some(42), "/other/app"),
        ];
        assert_eq!(
            owned(&rows, 10, Some(42))
                .iter()
                .map(|(row, _)| row.pid)
                .collect::<Vec<_>>(),
            vec![10, 20, 30, 40]
        );
        assert_eq!(owned(&rows, 10, None).len(), 3);
    }
    #[test]
    fn public_task_info_returns_nonzero_memory_for_this_process() {
        let task = info::<Task>(std::process::id() as i32, 4).unwrap();
        assert!(task.resident_size > 0);
        assert!(task.virtual_size >= task.resident_size);
        assert_eq!(std::mem::size_of::<Task>(), 96);
        assert_eq!(std::mem::size_of::<Bsd>(), 64);
    }
}
