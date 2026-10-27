use super::locked::Locked;
use std::{
    fs::{File, OpenOptions},
    os::{fd::AsRawFd, unix::fs::OpenOptionsExt},
    path::{Path, PathBuf},
    sync::Mutex,
};

/// The shell owns this lock for its whole lifetime, including a Node crash.
/// A lock file may remain on disk; the kernel lock, not that file's existence,
/// determines whether another process is writing. Never unlink a live lock.
#[derive(Debug)]
pub struct Profile {
    pub root: PathBuf,
    pub data_store_id: [u8; 16],
    lock: Mutex<Option<File>>,
}
impl Profile {
    pub fn acquire(root: PathBuf) -> Result<Self, String> {
        std::fs::create_dir_all(&root)
            .map_err(|e| format!("unavailable: Cannot create profile: {e}"))?;
        let root = root.canonicalize().map_err(|e| e.to_string())?;
        if root.join(".tauri-importing.json").exists() {
            return Err("conflict: This profile import is incomplete; use a completed import or recover it before opening".into());
        }
        // The frozen Electron build cannot participate in this new lock. Never
        // silently reuse its production profile or add files to it.
        if !root.join(".tauri-profile.json").exists()
            && ["Preferences", "Local State", "GPUCache"]
                .iter()
                .any(|name| root.join(name).exists())
        {
            return Err("conflict: Import the Electron profile into a separate Tauri profile before opening it".into());
        }
        let lock = OpenOptions::new()
            .read(true)
            .write(true)
            .create(true)
            .mode(0o600)
            .custom_flags(libc::O_NOFOLLOW)
            .open(root.join(".tauri-profile.lock"))
            .map_err(|e| format!("permission_denied: Cannot open profile lock: {e}"))?;
        if unsafe { libc::flock(lock.as_raw_fd(), libc::LOCK_EX | libc::LOCK_NB) } != 0 {
            return Err(format!(
                "conflict: This profile is already open: {}",
                root.display()
            ));
        }
        let data_store_id = store_identity(&root)?;
        Ok(Self {
            root,
            data_store_id,
            lock: Mutex::new(Some(lock)),
        })
    }
    pub fn release(&self) {
        self.lock.locked().take();
    }
    pub fn duplicate_lock(&self) -> Result<File, String> {
        self.lock
            .locked()
            .as_ref()
            .ok_or("unavailable: Profile was released")?
            .try_clone()
            .map_err(|e| e.to_string())
    }
}
fn store_identity(root: &Path) -> Result<[u8; 16], String> {
    use std::io::{Read, Write};
    let marker = root.join(".tauri-profile.json");
    let mut value = match OpenOptions::new()
        .read(true)
        .custom_flags(libc::O_NOFOLLOW)
        .open(&marker)
    {
        Ok(file) => {
            if !file.metadata().map_err(|e| e.to_string())?.is_file() {
                return Err("validation: Invalid profile metadata".into());
            }
            let mut bytes = Vec::new();
            file.take(8193)
                .read_to_end(&mut bytes)
                .map_err(|e| e.to_string())?;
            if bytes.len() > 8192 {
                return Err("validation: Profile metadata exceeds bound".into());
            }
            let value: serde_json::Value = serde_json::from_slice(&bytes)
                .map_err(|_| "validation: Invalid profile metadata")?;
            if value["version"] != 1 || value["runtime"] != "tauri" {
                return Err("validation: Unsupported profile metadata".into());
            }
            value
        }
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => {
            serde_json::json!({"version":1,"runtime":"tauri"})
        }
        Err(error) => {
            return Err(format!(
                "permission_denied: Cannot read profile metadata: {error}"
            ));
        }
    };
    let stored = value["dataStoreId"]
        .as_str()
        .map(uuid::Uuid::parse_str)
        .transpose()
        .map_err(|_| "validation: Invalid WebKit store identifier")?;
    // A directory copy must not share browser storage with its original.
    let id = if value["dataStoreRoot"] == root.to_string_lossy().as_ref() {
        stored.unwrap_or_else(uuid::Uuid::new_v4)
    } else {
        uuid::Uuid::new_v4()
    };
    if id.is_nil() {
        return Err("validation: Invalid WebKit store identifier".into());
    }
    value["dataStoreId"] = id.to_string().into();
    value["dataStoreRoot"] = root.to_string_lossy().as_ref().into();
    let temporary = root.join(format!(".tauri-profile-{}.tmp", uuid::Uuid::new_v4()));
    let result = (|| {
        let mut file = OpenOptions::new()
            .write(true)
            .create_new(true)
            .mode(0o600)
            .open(&temporary)
            .map_err(|e| e.to_string())?;
        file.write_all(serde_json::to_string(&value).unwrap().as_bytes())
            .and_then(|_| file.sync_all())
            .map_err(|e| e.to_string())?;
        std::fs::rename(&temporary, &marker).map_err(|e| e.to_string())?;
        File::open(root)
            .and_then(|dir| dir.sync_all())
            .map_err(|e| e.to_string())
    })();
    if result.is_err() {
        let _ = std::fs::remove_file(temporary);
    }
    result?;
    Ok(*id.as_bytes())
}
/// The profile to open: `--profile DIR`, else `{prefix}_USER_DATA`, else `default`.
pub fn selected(prefix: &str, default: PathBuf) -> Result<PathBuf, String> {
    let mut args = std::env::args_os().skip(1);
    while let Some(arg) = args.next() {
        if arg == "--profile" {
            let path = args
                .next()
                .ok_or("validation: --profile requires a directory")?;
            if path.is_empty() {
                return Err("validation: Profile directory is empty".into());
            }
            return Ok(PathBuf::from(path));
        }
    }
    Ok(std::env::var_os(format!("{prefix}_USER_DATA"))
        .map(PathBuf::from)
        .unwrap_or(default))
}

/// Restart is launched while the old process is leaving its event loop. Wait
/// before Tauri's single-instance plugin runs, so it cannot hand the new launch
/// back to the exiting instance. The profile lock remains the writer boundary.
pub fn wait_for_restart_parent() -> Result<(), String> {
    let args: Vec<_> = std::env::args_os().collect();
    if let Some(index) = args.iter().position(|arg| arg == "--restart-after") {
        let pid = args
            .get(index + 1)
            .and_then(|arg| arg.to_str())
            .and_then(|arg| arg.parse::<i32>().ok())
            .filter(|pid| *pid > 1)
            .ok_or("validation: Invalid restart parent")?;
        let start = std::time::Instant::now();
        while unsafe { libc::kill(pid, 0) } == 0 {
            if start.elapsed() > std::time::Duration::from_secs(10) {
                return Err("timeout: Previous application has not exited".into());
            }
            std::thread::sleep(std::time::Duration::from_millis(25));
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{BufRead, BufReader};
    /// Where the child process of the lock test locks its profile.
    const LOCK_CHILD: &str = "STANDALONE_HOST_LOCK_CHILD_PATH";
    // fork temporarily inherits every descriptor until exec closes CLOEXEC
    // handles. Serialize these process-spawning fixtures so an unrelated child
    // cannot briefly extend another test's lock lifetime.
    static FORK_FIXTURES: Mutex<()> = Mutex::new(());
    fn temporary() -> PathBuf {
        std::env::temp_dir().join(format!(
            "tauri-profile-{}-{}",
            std::process::id(),
            uuid::Uuid::new_v4()
        ))
    }
    #[test]
    fn lock_rejects_second_holder_and_releases_without_deleting_inode() {
        let _serial = FORK_FIXTURES.locked();
        let root = temporary();
        let first = Profile::acquire(root.clone()).unwrap();
        assert!(Profile::acquire(root.clone()).is_err());
        first.release();
        assert!(Profile::acquire(root.clone()).is_ok());
        assert!(root.join(".tauri-profile.lock").exists());
        std::fs::remove_dir_all(root).unwrap();
    }
    #[test]
    fn electron_profile_and_symlink_lock_are_rejected() {
        let _serial = FORK_FIXTURES.locked();
        let root = temporary();
        std::fs::create_dir(&root).unwrap();
        std::fs::write(root.join("Preferences"), b"{}").unwrap();
        assert!(
            Profile::acquire(root.clone())
                .unwrap_err()
                .contains("Electron")
        );
        std::fs::remove_file(root.join("Preferences")).unwrap();
        std::os::unix::fs::symlink("Preferences", root.join(".tauri-profile.lock")).unwrap();
        assert!(Profile::acquire(root.clone()).is_err());
        std::fs::remove_dir_all(root).unwrap();
    }
    #[test]
    fn incomplete_import_is_never_opened_or_marked_ready() {
        let root = temporary();
        std::fs::create_dir(&root).unwrap();
        std::fs::write(root.join(".tauri-importing.json"), b"{}").unwrap();
        assert!(
            Profile::acquire(root.clone())
                .unwrap_err()
                .contains("incomplete")
        );
        assert!(!root.join(".tauri-profile.json").exists());
        assert!(!root.join(".tauri-profile.lock").exists());
        std::fs::remove_dir_all(root).unwrap();
    }
    #[test]
    fn browser_store_survives_reopen_and_is_distinct_for_other_or_copied_profiles() {
        let root = temporary();
        let first = Profile::acquire(root.clone()).unwrap();
        let id = first.data_store_id;
        first.release();
        let again = Profile::acquire(root.clone()).unwrap();
        assert_eq!(again.data_store_id, id);
        again.release();
        let other = temporary();
        std::fs::create_dir(&other).unwrap();
        std::fs::copy(
            root.join(".tauri-profile.json"),
            other.join(".tauri-profile.json"),
        )
        .unwrap();
        let copied = Profile::acquire(other.clone()).unwrap();
        assert_ne!(copied.data_store_id, id);
        copied.release();
        std::fs::remove_dir_all(root).unwrap();
        std::fs::remove_dir_all(other).unwrap();
    }
    #[test]
    fn child_profile_owner() {
        if let Some(root) = std::env::var_os(LOCK_CHILD) {
            let _owner = Profile::acquire(root.into()).unwrap();
            println!("LOCKED");
            loop {
                std::thread::sleep(std::time::Duration::from_millis(100));
            }
        }
    }
    #[test]
    fn a_crashed_process_releases_the_kernel_lock() {
        let _serial = FORK_FIXTURES.locked();
        let root = temporary();
        let mut child = std::process::Command::new(std::env::current_exe().unwrap())
            .args([
                "--exact",
                "host::profile::tests::child_profile_owner",
                "--nocapture",
            ])
            .env(LOCK_CHILD, &root)
            .stdout(std::process::Stdio::piped())
            .spawn()
            .unwrap();
        let mut reader = BufReader::new(child.stdout.take().unwrap());
        let mut line = String::new();
        while !line.contains("LOCKED") {
            line.clear();
            assert!(reader.read_line(&mut line).unwrap() > 0);
        }
        assert!(Profile::acquire(root.clone()).is_err());
        child.kill().unwrap();
        child.wait().unwrap();
        assert!(Profile::acquire(root.clone()).is_ok());
        std::fs::remove_dir_all(root).unwrap();
    }
    #[test]
    fn inherited_writer_keeps_lock_after_shell_releases_its_handle() {
        let _serial = FORK_FIXTURES.locked();
        use std::os::unix::process::CommandExt;
        let root = temporary();
        let parent = Profile::acquire(root.clone()).unwrap();
        let inherited = parent.duplicate_lock().unwrap();
        let fd = inherited.as_raw_fd();
        let mut command = std::process::Command::new("/bin/sleep");
        command.arg("60");
        unsafe {
            command.pre_exec(move || {
                let flags = libc::fcntl(fd, libc::F_GETFD);
                if flags < 0 || libc::fcntl(fd, libc::F_SETFD, flags & !libc::FD_CLOEXEC) < 0 {
                    return Err(std::io::Error::last_os_error());
                }
                Ok(())
            });
        }
        let mut child = command.spawn().unwrap();
        drop(inherited);
        parent.release();
        assert!(Profile::acquire(root.clone()).is_err());
        child.kill().unwrap();
        child.wait().unwrap();
        assert!(Profile::acquire(root.clone()).is_ok());
        std::fs::remove_dir_all(root).unwrap();
    }
}
