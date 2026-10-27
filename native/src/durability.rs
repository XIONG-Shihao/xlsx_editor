//! Durability for one quiescent native XLSX index, separate from its owning
//! application's history manifest and publication ledger.
use crate::{Result, cancel, err, session_root};
use rusqlite::{Connection, OpenFlags, OptionalExtension};
use serde::Deserialize;
use std::{
    collections::{HashMap, HashSet},
    fs::{self, File, Metadata},
    io::BufReader,
    path::{Path, PathBuf},
};

const CORE_ARTIFACTS: [&str; 4] = ["source.xlsx", "metadata.xlsx", "index.json", "cells.sqlite"];
const INDEX_LIMIT: u64 = 64 * 1024 * 1024;

/// Validate a newly indexed or recovered session before accepting its native
/// versions. This performs SQLite quick_check, checks retained ancestry and
/// index-declared worksheet artifacts, but does not hash/re-copy the database
/// or re-certify the application's immutable-base hash. It reads the native
/// index metadata and database pages; call it at admission/recovery, not after
/// each edit. Raw/passive indexes are outside this XLSX session boundary.
pub fn validate_session(root: &Path, versions: &[u64]) -> Result<()> {
    let session = Session::open(root)?;
    let db = session.database()?;
    retained_versions(&db, versions)?;
    retained_ancestry(&db, versions)?;
    let mut integrity = db.prepare_cached("PRAGMA quick_check(1)").map_err(err)?;
    let mut rows = integrity.query([]).map_err(err)?;
    if rows
        .next()
        .map_err(err)?
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing SQLite integrity result"))?
        .get::<_, String>(0)
        .map_err(err)?
        != "ok"
        || rows.next().map_err(err)?.is_some()
    {
        return Err(crate::coded(
            "session-refused",
            "Workbook session SQLite integrity check failed",
        ));
    }
    for sql in [
        "SELECT id,parent FROM versions LIMIT 0",
        "SELECT version,sheet,r,c,data,edit FROM changes LIMIT 0",
        "SELECT sheet,r,c,kind,raw,inline,formula,formula_attrs,style,start,open,close,end,spans FROM cells LIMIT 0",
        "SELECT id,value FROM strings LIMIT 0",
        "SELECT sheet,r,start,open,close,end FROM rows LIMIT 0",
    ] {
        db.prepare_cached(sql).map_err(err)?;
    }
    if db
        .query_row(
            "SELECT EXISTS(SELECT 1 FROM sqlite_master WHERE type IN('view','trigger'))",
            [],
            |row| row.get::<_, bool>(0),
        )
        .map_err(err)?
    {
        return Err(crate::coded(
            "session-refused",
            "Workbook session executable schema is refused",
        ));
    }
    let index = session.artifact("index.json")?;
    if index.before.len() > INDEX_LIMIT {
        return Err(crate::coded(
            "session-limit",
            "Workbook session index metadata exceeds its byte bound",
        ));
    }
    let index_file = index.open()?;
    let index_data: Index = serde_json::from_reader(BufReader::new(&index_file)).map_err(err)?;
    index.check_file(&index_file)?;
    let index = index_data;
    if index.storage != "rust-sqlite-v1"
        || index.revision.len() != 64
        || !index
            .revision
            .bytes()
            .all(|byte| byte.is_ascii_digit() || (b'a'..=b'f').contains(&byte))
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Workbook session index identity is invalid",
        ));
    }
    let mut paths = HashSet::new();
    for (ordinal, sheet) in index.sheets.iter().enumerate() {
        cancel::check()?;
        if sheet.path.is_empty()
            || !paths.insert(&sheet.path)
            || sheet.artifact != format!("sheet-{ordinal}.xml")
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Workbook session worksheet artifact identity is invalid",
            ));
        }
        session.artifact(&sheet.artifact)?;
    }
    drop(rows);
    drop(integrity);
    drop(db);
    session.check()
}

/// Persist already committed native versions before the application advances
/// its accepted manifest. The caller holds the single document-writer lease and
/// permits no native transaction or session-file mutation until this returns.
/// Call validate_session after creation/recovery; immutable base artifacts must
/// remain unchanged between admissions. Native writers use SQLite's default
/// FULL synchronous DELETE journal commits; other journal profiles and leftover
/// journals are refused here. This read-only connection neither commits nor
/// changes the database schema.
///
/// Per call: retained-version indexed lookups, artifact metadata checks, fsync of
/// the database, source/metadata/index, worksheet XML and optional display-format
/// palette, then session and parent directories. Unchanged base-file syncs add
/// syscalls but no hash/data-copy/database-integrity scan. Derived exports and
/// unrelated temporary payloads are not read or fsynced. The application owns
/// durability of its manifest and any newly created directories above parent.
pub fn flush_session(root: &Path, versions: &[u64]) -> Result<()> {
    let session = Session::open(root)?;
    let db = session.database()?;
    retained_versions(&db, versions)?;
    drop(db);
    session.check()?;
    for artifact in session
        .artifacts
        .iter()
        .filter(|artifact| required_to_sync(&artifact.name))
    {
        cancel::check()?;
        let file = artifact.open()?;
        file.sync_all().map_err(err)?;
        artifact.check_file(&file)?;
    }
    cancel::check()?;
    session.directory.sync_all().map_err(err)?;
    session.parent.sync_all().map_err(err)?;
    session.check()
}

#[derive(Deserialize)]
struct Index {
    storage: String,
    revision: String,
    sheets: Vec<Sheet>,
}
#[derive(Deserialize)]
struct Sheet {
    path: String,
    artifact: String,
}

struct Artifact {
    name: String,
    path: PathBuf,
    before: Metadata,
}
impl Artifact {
    fn open(&self) -> Result<File> {
        self.check()?;
        let file = File::open(&self.path).map_err(err)?;
        self.check_file(&file)?;
        Ok(file)
    }
    fn check_file(&self, file: &File) -> Result<()> {
        self.check()?;
        if !same_file(&self.before, &file.metadata().map_err(err)?) {
            return Err(crate::coded(
                "session-stale",
                "Workbook session artifact changed during admission",
            ));
        }
        Ok(())
    }
    fn check(&self) -> Result<()> {
        let current = fs::symlink_metadata(&self.path).map_err(err)?;
        if !current.is_file()
            || current.file_type().is_symlink()
            || !same_file(&self.before, &current)
        {
            return Err(crate::coded(
                "session-stale",
                "Workbook session artifact changed during durability work",
            ));
        }
        Ok(())
    }
}
struct Session {
    root: PathBuf,
    directory: File,
    before: Metadata,
    parent: File,
    artifacts: Vec<Artifact>,
}
impl Session {
    fn open(supplied: &Path) -> Result<Self> {
        let root = session_root(supplied)?;
        cancel::begin(&root);
        cancel::check()?;
        let before = fs::symlink_metadata(&root).map_err(err)?;
        if !before.is_dir() || before.file_type().is_symlink() || !safe_directory(&before) {
            return Err(crate::coded(
                "session-unsupported",
                "Unsafe workbook durability directory",
            ));
        }
        let directory = File::open(&root).map_err(err)?;
        if !same_file(&before, &directory.metadata().map_err(err)?) {
            return Err(crate::coded(
                "session-stale",
                "Workbook durability directory changed during admission",
            ));
        }
        let parent = File::open(root.parent().ok_or_else(|| {
            crate::coded(
                "session-unsupported",
                "Workbook durability parent is unavailable",
            )
        })?)
        .map_err(err)?;
        let mut artifacts = Vec::new();
        for entry in fs::read_dir(&root).map_err(err)? {
            cancel::check()?;
            let entry = entry.map_err(err)?;
            let name = entry.file_name().into_string().map_err(|_| {
                crate::coded(
                    "session-bad-input",
                    "Workbook session artifact name is not UTF-8",
                )
            })?;
            let path = entry.path();
            let info = fs::symlink_metadata(&path).map_err(err)?;
            if !info.is_file() || info.file_type().is_symlink() || !safe_artifact(&before, &info) {
                return Err(crate::coded(
                    "session-unsupported",
                    "Unsafe workbook session artifact; regular private files are required",
                ));
            }
            if [
                "cells.sqlite-journal",
                "cells.sqlite-wal",
                "cells.sqlite-shm",
            ]
            .contains(&name.as_str())
            {
                return Err(crate::coded(
                    "session-bad-input",
                    "Workbook durability requires a committed DELETE-journal session",
                ));
            }
            // Keep metadata, not one descriptor per sheet. Each artifact is
            // opened and checked only when needed; large books cannot exhaust
            // the descriptor limit, and temporary payloads are never read.
            artifacts.push(Artifact {
                name,
                path,
                before: info,
            });
        }
        artifacts.sort_by(|a, b| a.name.cmp(&b.name));
        let session = Self {
            root,
            directory,
            before,
            parent,
            artifacts,
        };
        for name in CORE_ARTIFACTS {
            session.artifact(name)?;
        }
        session.check()?;
        Ok(session)
    }
    fn artifact(&self, name: &str) -> Result<&Artifact> {
        let artifact = self
            .artifacts
            .binary_search_by(|artifact| artifact.name.as_str().cmp(name))
            .ok()
            .map(|index| &self.artifacts[index])
            .ok_or_else(|| {
                crate::coded(
                    "internal-missing-field",
                    "Required workbook session artifact is missing",
                )
            })?;
        if artifact.before.len() == 0 {
            return Err(crate::coded(
                "session-bad-input",
                "Required workbook session artifact is empty",
            ));
        }
        Ok(artifact)
    }
    fn database(&self) -> Result<Connection> {
        let artifact = self.artifact("cells.sqlite")?;
        let db = Connection::open_with_flags(
            &artifact.path,
            OpenFlags::SQLITE_OPEN_READ_ONLY
                | OpenFlags::SQLITE_OPEN_NO_MUTEX
                | OpenFlags::SQLITE_OPEN_NOFOLLOW,
        )
        .map_err(err)?;
        db.progress_handler(10_000, Some(cancel::interrupted));
        db.execute_batch("PRAGMA query_only=ON; PRAGMA cache_size=-4096; PRAGMA temp_store=FILE; PRAGMA mmap_size=0;").map_err(err)?;
        let mode: String = db
            .query_row("PRAGMA journal_mode", [], |row| row.get(0))
            .map_err(err)?;
        let synchronous: u32 = db
            .query_row("PRAGMA synchronous", [], |row| row.get(0))
            .map_err(err)?;
        if mode != "delete" || synchronous < 2 {
            return Err(crate::coded(
                "session-bad-input",
                "Workbook durability requires SQLite FULL synchronous DELETE journaling",
            ));
        }
        artifact.check()?;
        Ok(db)
    }
    fn check(&self) -> Result<()> {
        cancel::check()?;
        let current = fs::symlink_metadata(&self.root).map_err(err)?;
        if !current.is_dir()
            || current.file_type().is_symlink()
            || !same_file(&self.before, &current)
            || !same_file(&self.before, &self.directory.metadata().map_err(err)?)
        {
            return Err(crate::coded(
                "session-stale",
                "Workbook session directory changed during durability work",
            ));
        }
        for artifact in &self.artifacts {
            artifact.check()?;
        }
        Ok(())
    }
}

fn required_to_sync(name: &str) -> bool {
    CORE_ARTIFACTS.contains(&name)
        || name == "display-formats.json"
        || name
            .strip_prefix("sheet-")
            .and_then(|name| name.strip_suffix(".xml"))
            .is_some_and(|index| {
                !index.is_empty() && index.bytes().all(|byte| byte.is_ascii_digit())
            })
}
fn retained_versions(db: &Connection, versions: &[u64]) -> Result<()> {
    if versions.is_empty() {
        return Err(crate::coded(
            "session-refused",
            "At least one retained workbook version is required",
        ));
    }
    let mut query = db
        .prepare_cached("SELECT EXISTS(SELECT 1 FROM versions WHERE id=?1)")
        .map_err(err)?;
    for version in versions {
        cancel::check()?;
        let version = i64::try_from(*version).map_err(|_| {
            crate::coded(
                "internal-invalid-field",
                "Invalid retained workbook version",
            )
        })?;
        if !query
            .query_row([version], |row| row.get::<_, bool>(0))
            .map_err(err)?
        {
            return Err(crate::coded(
                "internal-unknown-reference",
                "Unknown retained workbook version",
            ));
        }
    }
    Ok(())
}
fn retained_ancestry(db: &Connection, versions: &[u64]) -> Result<()> {
    let mut complete = HashSet::new();
    let mut parents = HashMap::new();
    let mut query = db
        .prepare_cached("SELECT parent FROM versions WHERE id=?1")
        .map_err(err)?;
    for version in versions {
        let mut path = HashSet::new();
        let mut at = Some(*version);
        while let Some(version) = at {
            cancel::check()?;
            if complete.contains(&version) {
                break;
            }
            if !path.insert(version) {
                return Err(crate::coded(
                    "session-malformed",
                    "Cyclic retained workbook ancestry",
                ));
            }
            let parent = if let Some(parent) = parents.get(&version) {
                *parent
            } else {
                let parent: Option<u64> = query
                    .query_row(
                        [i64::try_from(version).map_err(|_| {
                            crate::coded("internal-invalid-field", "Invalid workbook ancestor")
                        })?],
                        |row| row.get(0),
                    )
                    .optional()
                    .map_err(err)?
                    .ok_or_else(|| {
                        crate::coded(
                            "internal-missing-field",
                            "Missing retained workbook ancestor",
                        )
                    })?;
                parents.insert(version, parent);
                parent
            };
            if parent.is_none() && version != 0 {
                return Err(crate::coded(
                    "session-unsupported",
                    "Retained workbook ancestry does not reach its imported base",
                ));
            }
            at = parent;
        }
        complete.extend(path);
    }
    Ok(())
}

#[cfg(unix)]
fn safe_directory(info: &Metadata) -> bool {
    use std::os::unix::fs::MetadataExt;
    info.mode() & 0o022 == 0
}
#[cfg(not(unix))]
fn safe_directory(_: &Metadata) -> bool {
    true
}
#[cfg(unix)]
fn safe_artifact(directory: &Metadata, file: &Metadata) -> bool {
    use std::os::unix::fs::MetadataExt;
    file.uid() == directory.uid() && file.nlink() == 1 && file.mode() & 0o022 == 0
}
#[cfg(not(unix))]
fn safe_artifact(_: &Metadata, _: &Metadata) -> bool {
    true
}
fn same_file(before: &Metadata, after: &Metadata) -> bool {
    #[cfg(unix)]
    {
        use std::os::unix::fs::MetadataExt;
        before.dev() == after.dev()
            && before.ino() == after.ino()
            && before.uid() == after.uid()
            && before.nlink() == after.nlink()
            && before.len() == after.len()
            && before.mtime() == after.mtime()
            && before.mtime_nsec() == after.mtime_nsec()
            && before.ctime() == after.ctime()
            && before.ctime_nsec() == after.ctime_nsec()
    }
    #[cfg(not(unix))]
    {
        before.len() == after.len()
            && before.modified().ok() == after.modified().ok()
            && before.created().ok() == after.created().ok()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::{Value, json};
    use std::io::Write;
    const SHEET: &str = "xl/worksheets/main.xml";

    fn fixture() -> tempfile::TempDir {
        let root = tempfile::tempdir().unwrap();
        let mut zip = zip::ZipWriter::new(File::create(root.path().join("source.xlsx")).unwrap());
        for (name, text) in [
            (
                "[Content_Types].xml",
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Override PartName=\"/xl/worksheets/main.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>",
            ),
            (
                "xl/workbook.xml",
                "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"Main\" sheetId=\"1\" r:id=\"main\"/></sheets></workbook>",
            ),
            (
                "xl/_rels/workbook.xml.rels",
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"main\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/main.xml\"/></Relationships>",
            ),
            (
                SHEET,
                "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><dimension ref=\"A1:B1\"/><sheetData><row r=\"1\"><c r=\"A1\"><v>2</v></c><c r=\"B1\"><f>A1+1</f><v>91</v></c></row></sheetData></worksheet>",
            ),
            (
                "xl/styles.xml",
                "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><fonts count=\"1\"><font/></fonts><fills count=\"1\"><fill/></fills><borders count=\"1\"><border/></borders><cellXfs count=\"1\"><xf numFmtId=\"0\"/></cellXfs></styleSheet>",
            ),
            ("custom/preserved.bin", "Preserve these source bytes"),
        ] {
            zip.start_file(name, zip::write::SimpleFileOptions::default())
                .unwrap();
            zip.write_all(text.as_bytes()).unwrap();
        }
        zip.finish().unwrap();
        crate::command(json!({"operation":"index","root":root.path()})).unwrap();
        root
    }
    fn calculate(root: &Path, version: u64) -> u64 {
        let view = json!({"readOnly":false,"revision":crate::hash(&root.join("source.xlsx")).unwrap(),"sheets":[{"name":"Main","path":SHEET,"rows":1,"columns":2,"date1904":false,"merges":[],"storedFormulaCount":1}],"namedRanges":[],"tables":[]});
        let reply=crate::command(json!({"operation":"calculate","root":root,"version":version,"revision":view["revision"],"view":view,"metadata":"durability-calculation"})).unwrap();
        assert_eq!(reply["supported"], true);
        reply["version"].as_u64().unwrap()
    }
    fn read(root: &Path, version: u64) -> Value {
        crate::command(json!({"operation":"range","root":root,"version":version,"sheet":SHEET,"range":"B1","limit":1})).unwrap()["cells"][0][1].clone()
    }

    #[test]
    fn retained_calculation_versions_survive_flushed_session_reopen() {
        let root = fixture();
        validate_session(root.path(), &[0]).unwrap();
        let first = calculate(root.path(), 0);
        let literal =
            json!({"value":4.0,"dataType":"number","rawValue":"4","cachedValue":null,"styleId":0});
        let edited=crate::command(json!({"operation":"patch","root":root.path(),"version":first,"changes":[{"sheet":SHEET,"cell":"A1","value":literal,"edits":[{"sheet":"Main","cell":"A1","value":4.0}]}]})).unwrap()["version"].as_u64().unwrap();
        let second = calculate(root.path(), edited);
        let before = fs::read(root.path().join("cells.sqlite")).unwrap();
        flush_session(root.path(), &[0, first, second]).unwrap();
        assert_eq!(
            before,
            fs::read(root.path().join("cells.sqlite")).unwrap(),
            "Flush must not rewrite native history/schema"
        );
        let cold = tempfile::tempdir().unwrap();
        for entry in fs::read_dir(root.path()).unwrap() {
            let entry = entry.unwrap();
            if required_to_sync(entry.file_name().to_str().unwrap()) {
                fs::copy(entry.path(), cold.path().join(entry.file_name())).unwrap();
            }
        }
        validate_session(cold.path(), &[0, first, second]).unwrap();
        assert_eq!(read(cold.path(), first)["calculation"]["value"], 3.0);
        assert_eq!(read(cold.path(), second)["calculation"]["value"], 5.0);
        let db = crate::database(cold.path()).unwrap();
        assert_eq!(
            db.prepare_cached("PRAGMA synchronous")
                .unwrap()
                .query_row([], |row| row.get::<_, u32>(0))
                .unwrap(),
            2
        );
    }

    #[test]
    fn missing_versions_and_invalid_ancestry_cannot_be_admitted() {
        let root = fixture();
        assert!(
            flush_session(root.path(), &[99])
                .unwrap_err()
                .contains("Unknown retained")
        );
        assert!(
            validate_session(root.path(), &[99])
                .unwrap_err()
                .contains("Unknown retained")
        );
        assert!(flush_session(root.path(), &[]).is_err());
        let db = crate::database(root.path()).unwrap();
        db.execute("INSERT INTO versions(id,parent)VALUES(1,99)", [])
            .unwrap();
        assert!(
            validate_session(root.path(), &[1])
                .unwrap_err()
                .contains("ancestor")
        );
        db.execute("UPDATE versions SET parent=1 WHERE id=1", [])
            .unwrap();
        assert!(
            validate_session(root.path(), &[1])
                .unwrap_err()
                .contains("Cyclic")
        );
    }

    #[test]
    fn corruption_and_missing_required_import_artifacts_fail_validation() {
        let root = fixture();
        fs::write(root.path().join("cells.sqlite"), b"This is not SQLite").unwrap();
        assert!(validate_session(root.path(), &[0]).is_err());
        assert!(flush_session(root.path(), &[0]).is_err());
        let root = fixture();
        fs::remove_file(root.path().join("sheet-0.xml")).unwrap();
        assert!(
            validate_session(root.path(), &[0])
                .unwrap_err()
                .contains("missing")
        );
        fs::remove_file(root.path().join("metadata.xlsx")).unwrap();
        assert!(
            flush_session(root.path(), &[0])
                .unwrap_err()
                .contains("missing")
        );
    }

    #[test]
    fn cancellation_refuses_flush_and_validation_before_acceptance() {
        let root = fixture();
        fs::write(root.path().join("cancelled"), b"1").unwrap();
        for check in [flush_session, validate_session] {
            assert!(check(root.path(), &[0]).unwrap_err().contains("cancelled"));
        }
        fs::remove_file(root.path().join("cancelled")).unwrap();
        validate_session(root.path(), &[0]).unwrap();
        flush_session(root.path(), &[0]).unwrap();
    }

    #[test]
    fn unrelated_candidate_payloads_are_not_parsed_or_used_for_recovery() {
        let root = fixture();
        for name in [
            "export.xlsx",
            "export.partial",
            "export-metadata.json",
            "layout-fragment.xml",
        ] {
            fs::write(root.path().join(name), b"Unusable derived candidate bytes").unwrap();
        }
        validate_session(root.path(), &[0]).unwrap();
        flush_session(root.path(), &[0]).unwrap();
        for name in [
            "export.xlsx",
            "export.partial",
            "export-metadata.json",
            "layout-fragment.xml",
        ] {
            assert_eq!(
                fs::read(root.path().join(name)).unwrap(),
                b"Unusable derived candidate bytes"
            );
        }
        assert_eq!(read(root.path(), 0)["cachedValue"], 91.0);
    }

    #[test]
    fn index_artifact_paths_cannot_escape_the_session() {
        let root = fixture();
        let index = root.path().join("index.json");
        let mut value: Value = serde_json::from_slice(&fs::read(&index).unwrap()).unwrap();
        value["sheets"][0]["artifact"] = json!("../foreign.xml");
        fs::write(&index, serde_json::to_vec(&value).unwrap()).unwrap();
        assert!(
            validate_session(root.path(), &[0])
                .unwrap_err()
                .contains("artifact identity")
        );
    }

    #[cfg(unix)]
    #[test]
    fn symlinks_hardlinks_and_nonregular_artifacts_are_refused() {
        use std::os::unix::fs::symlink;
        let root = fixture();
        let foreign = tempfile::tempdir().unwrap();
        let bytes = b"Foreign source must remain unchanged";
        fs::write(foreign.path().join("source.xlsx"), bytes).unwrap();
        fs::remove_file(root.path().join("source.xlsx")).unwrap();
        symlink(
            foreign.path().join("source.xlsx"),
            root.path().join("source.xlsx"),
        )
        .unwrap();
        assert!(
            validate_session(root.path(), &[0])
                .unwrap_err()
                .contains("Unsafe")
        );
        assert!(
            flush_session(root.path(), &[0])
                .unwrap_err()
                .contains("Unsafe")
        );
        assert_eq!(fs::read(foreign.path().join("source.xlsx")).unwrap(), bytes);
        let root = fixture();
        symlink(
            root.path().join("source.xlsx"),
            root.path().join("export.xlsx"),
        )
        .unwrap();
        assert!(
            flush_session(root.path(), &[0]).is_err(),
            "An unrelated artifact still cannot be a symlink"
        );
        let root = fixture();
        fs::hard_link(
            root.path().join("source.xlsx"),
            foreign.path().join("alias.xlsx"),
        )
        .unwrap();
        assert!(flush_session(root.path(), &[0]).is_err());
        let root = fixture();
        fs::create_dir(root.path().join("unexpected-directory")).unwrap();
        assert!(flush_session(root.path(), &[0]).is_err());
        let alias = foreign.path().join("session-alias");
        symlink(root.path(), &alias).unwrap();
        assert!(flush_session(&alias, &[0]).is_err());
    }

    #[test]
    fn wal_or_unfinished_journal_profiles_do_not_get_false_durable_receipts() {
        let root = fixture();
        let db = crate::database(root.path()).unwrap();
        db.execute_batch("PRAGMA journal_mode=WAL;").unwrap();
        drop(db);
        assert!(flush_session(root.path(), &[0]).is_err());
        let root = fixture();
        fs::write(
            root.path().join("cells.sqlite-journal"),
            b"Unfinished journal",
        )
        .unwrap();
        assert!(
            flush_session(root.path(), &[0])
                .unwrap_err()
                .contains("committed DELETE")
        );
    }
}
