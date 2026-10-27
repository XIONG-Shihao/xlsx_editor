//! Conservative per-session disk reservations, including SQLite spill/journal
//! headroom. Budgeted writers refuse bytes before writing them. Physical filesystem
//! metadata/allocation blocks are outside the logical-byte accounting contract.
use crate::{Result, err};
use std::{
    cell::RefCell,
    fs::{self, File, OpenOptions, TryLockError},
    io::{self, Seek, Write},
    path::{Path, PathBuf},
    time::Duration,
};
thread_local! {static BUDGET:RefCell<Option<(u64,u64,u64)>>=const{RefCell::new(None)};}
const DEFAULT_LIMIT: u64 = 4 * 1024 * 1024 * 1024;
pub fn used(path: &Path) -> Result<u64> {
    let mut total = 0;
    for entry in fs::read_dir(path).map_err(err)? {
        let entry = entry.map_err(err)?;
        let stat = entry.metadata().map_err(err)?;
        if entry.file_type().map_err(err)?.is_symlink() {
            return Err(crate::coded(
                "session-refused",
                "Symlink in workbook session",
            ));
        }
        if stat.is_dir() {
            total += used(&entry.path())?
        } else {
            total += stat.len();
            if entry.file_name() == "candidate-reservation.json" {
                let data: serde_json::Value =
                    serde_json::from_slice(&fs::read(entry.path()).map_err(err)?).map_err(err)?;
                total += data["bytes"].as_u64().ok_or_else(|| {
                    crate::coded("internal-invalid-field", "Invalid disk reservation")
                })?;
            }
        }
    }
    Ok(total)
}
pub struct Guard(File, PathBuf);
impl Drop for Guard {
    fn drop(&mut self) {
        BUDGET.with(|b| *b.borrow_mut() = None);
        let _ = fs::remove_file(&self.1);
        let _ = self.0.unlock();
    }
}
pub fn begin(root: &Path, operation: &str) -> Result<Guard> {
    let parent = root.parent().unwrap_or(root);
    let session = if parent.join("owner.json").is_file() {
        parent
    } else {
        root
    };
    let metadata = fs::read(session.join("owner.json"))
        .ok()
        .and_then(|b| serde_json::from_slice::<serde_json::Value>(&b).ok());
    let limit = metadata
        .as_ref()
        .and_then(|m| m["maxSessionBytes"].as_u64())
        .unwrap_or(DEFAULT_LIMIT)
        .min(DEFAULT_LIMIT);
    let lock = OpenOptions::new()
        .create(true)
        .truncate(false)
        .read(true)
        .write(true)
        .open(session.join(".quota.lock"))
        .map_err(err)?;
    loop {
        crate::cancel::check()?;
        match lock.try_lock() {
            Ok(()) => break,
            Err(TryLockError::WouldBlock) => std::thread::sleep(Duration::from_millis(10)),
            Err(TryLockError::Error(e)) => return Err(err(e)),
        }
    }
    let active = session.join("native-active.json");
    fs::write(
        &active,
        serde_json::to_vec(&serde_json::json!({"pid":std::process::id()})).map_err(err)?,
    )
    .map_err(err)?;
    let guard = Guard(lock, active);
    if operation == "release-candidate" {
        return Ok(guard);
    }
    let occupied = used(session)?;
    let free = limit.checked_sub(occupied).ok_or_else(|| {
        crate::coded(
            "session-disk-budget",
            "Workbook session disk budget exhausted; save and reopen",
        )
    })?;
    let db = fs::metadata(root.join("cells.sqlite"))
        .map(|m| m.len())
        .unwrap_or(0);
    let reading = read_only(operation);
    let available = if reading {
        free
    } else {
        free.checked_sub(db.saturating_mul(2)).ok_or_else(|| {
            crate::coded(
                "session-disk-budget",
                "Workbook session disk budget cannot reserve journal/compaction space; save and reopen",
            )
        })?
    };
    let slice = available / 6;
    if slice < 16384 {
        return Err(crate::coded(
            "session-disk-budget",
            "Workbook session disk budget exhausted; save and reopen",
        ));
    }
    BUDGET.with(|b| {
        *b.borrow_mut() = Some((
            slice,
            if reading { db.max(4096) } else { db + slice },
            slice,
        ))
    });
    Ok(guard)
}
pub fn database(db: &rusqlite::Connection) -> Result<()> {
    if let Some((_, main, temp)) = BUDGET.with(|b| *b.borrow()) {
        let size: u64 = db
            .query_row("PRAGMA page_size", [], |r| r.get(0))
            .map_err(err)?;
        db.execute_batch(&format!("PRAGMA max_page_count={}; PRAGMA temp.max_page_count={}; PRAGMA temp.journal_mode=OFF;",(main/size).max(1),(temp/4096).max(1))).map_err(err)?;
    }
    Ok(())
}
fn reserve(bytes: u64) -> io::Result<()> {
    crate::cancel::check().map_err(io::Error::other)?;
    BUDGET.with(|b| {
        let mut b = b.borrow_mut();
        if let Some((remaining, _, _)) = b.as_mut() {
            *remaining = remaining.checked_sub(bytes).ok_or_else(|| {
                io::Error::other(crate::coded(
                    "session-disk-budget",
                    "Workbook session disk budget exhausted; save and reopen",
                ))
            })?;
        }
        Ok(())
    })
}
/// Reserve an artifact written through SQLite rather than BudgetFile. Writers
/// include page/journal headroom and reserve before each bounded insert.
pub(crate) fn reserve_artifact(bytes: u64) -> Result<()> {
    reserve(bytes).map_err(err)
}
pub struct BudgetFile(File);
impl Write for BudgetFile {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        reserve(buffer.len() as u64)?;
        self.0.write(buffer)
    }
    fn flush(&mut self) -> io::Result<()> {
        self.0.flush()
    }
}
impl Seek for BudgetFile {
    fn seek(&mut self, p: io::SeekFrom) -> io::Result<u64> {
        self.0.seek(p)
    }
}
impl BudgetFile {
    pub fn sync_all(&self) -> io::Result<()> {
        self.0.sync_all()
    }
}
pub fn create(path: impl AsRef<Path>) -> io::Result<BudgetFile> {
    Ok(BudgetFile(File::create(path)?))
}
/// Native private reservations must never replace an existing artifact.
pub(crate) fn create_new(path: impl AsRef<Path>) -> io::Result<BudgetFile> {
    let mut options = OpenOptions::new();
    options.write(true).create_new(true);
    #[cfg(unix)]
    {
        use std::os::unix::fs::OpenOptionsExt;
        options.mode(0o600);
    }
    Ok(BudgetFile(options.open(path)?))
}
pub fn write(path: impl AsRef<Path>, contents: impl AsRef<[u8]>) -> io::Result<()> {
    let mut file = create(path)?;
    file.write_all(contents.as_ref())
}
pub fn artifact(root: &Path, input: &serde_json::Value) -> Result<serde_json::Value> {
    let name = crate::field(input, "name")?;
    if ![
        "export-metadata.json",
        "display-formats.json",
        "layout-fragment.xml",
        "package-metadata.json",
        "sort-order.jsonl",
        "dependency-plan.json",
        "insertion-metadata.json",
        "insertion-formulas.jsonl",
    ]
    .contains(&name)
    {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown workbook artifact",
        ));
    }
    let path = crate::local(root, name)?;
    let data = crate::field(input, "data")?;
    if input["append"].as_bool().unwrap_or(false) {
        let mut file = BudgetFile(
            OpenOptions::new()
                .create(true)
                .append(true)
                .open(path)
                .map_err(err)?,
        );
        file.write_all(data.as_bytes()).map_err(err)?;
    } else {
        write(path, data).map_err(err)?;
    }
    Ok(serde_json::json!({"written":true}))
}
pub fn copy(root: &Path, input: &serde_json::Value) -> Result<serde_json::Value> {
    let source = File::open(crate::field(input, "source")?).map_err(err)?;
    let mut destination = create(crate::local(root, "source.xlsx")?).map_err(err)?;
    io::copy(&mut crate::cancel::Reader(source), &mut destination).map_err(err)?;
    Ok(serde_json::json!({"copied":true}))
}

pub fn candidate(root: &Path, input: &serde_json::Value) -> Result<serde_json::Value> {
    let path = crate::local(root, "candidate-reservation.json")?;
    if crate::field(input, "operation")? == "release-candidate" {
        if path.exists() {
            fs::remove_file(path).map_err(err)?;
        }
    } else {
        let bytes = input["bytes"].as_u64().ok_or_else(|| {
            crate::coded("internal-missing-field", "Missing candidate byte count")
        })?;
        reserve(bytes).map_err(err)?;
        write(
            path,
            serde_json::to_vec(&serde_json::json!({"bytes":bytes})).map_err(err)?,
        )
        .map_err(err)?;
    }
    Ok(serde_json::json!({"reserved":true}))
}

/// Operations that only read a session (`read-only-operations.json`, shared
/// with the tests).
pub(crate) fn read_only(operation: &str) -> bool {
    static LIST: std::sync::OnceLock<Vec<String>> = std::sync::OnceLock::new();
    LIST.get_or_init(|| {
        let file: serde_json::Value = serde_json::from_str(include_str!("../read-only-operations.json"))
            .expect("read-only-operations.json is valid JSON");
        file["operations"]
            .as_array()
            .expect("read-only-operations.json lists operations")
            .iter()
            .filter_map(|value| value.as_str().map(str::to_owned))
            .collect()
    })
    .iter()
    .any(|name| name == operation)
}
