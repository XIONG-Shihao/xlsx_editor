//! What a host can check about the editor's durable journal files without
//! opening a workbook: staged paste attachments and calculation witnesses.
//! The host stores the files and checks their identity; their format is here.
use rusqlite::Connection;
use serde_json::Value;
use std::io::{BufRead, Read};
use std::path::Path;

const MIB: u64 = 1024 * 1024;

fn keys(value: &Value, expected: &[&str]) -> Result<(), String> {
    let object = value
        .as_object()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid workbook journal object"))?;
    if object.len() != expected.len() || object.keys().any(|key| !expected.contains(&key.as_str()))
    {
        return Err(crate::coded(
            "workbook-malformed",
            "Unexpected workbook journal fields",
        ));
    }
    Ok(())
}

fn typed_paste_conversion(value: &Value, rows: u64, columns: u64) -> Result<(), String> {
    keys(
        value,
        &[
            "schema",
            "policies",
            "date1904",
            "keepHeader",
            "outputRows",
            "outputColumns",
        ],
    )?;
    let policies = value["policies"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid typed query policies"))?;
    let output_rows = value["outputRows"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid typed query output rows"))?;
    let output_columns = value["outputColumns"].as_u64().ok_or_else(|| {
        crate::coded(
            "internal-invalid-field",
            "Invalid typed query output columns",
        )
    })?;
    if value["schema"] != 1
        || value["date1904"].as_bool().is_none()
        || value["keepHeader"].as_bool().is_none()
        || policies.is_empty()
        || policies.len() > 256
        || policies
            .iter()
            .any(|p| !["text", "number", "boolean", "iso-date"].contains(&p.as_str().unwrap_or("")))
        || output_rows > rows
        || output_columns == 0
        || output_columns > columns
        || output_columns != policies.len() as u64
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid typed query conversion metadata",
        ));
    }
    Ok(())
}

/// The rows, columns and cell count a staged paste descriptor declares.
pub fn staged_paste_shape(command: &Value) -> Result<(u64, u64, u64), String> {
    let object = command
        .as_object()
        .ok_or_else(|| crate::coded("workbook-malformed", "Invalid staged import descriptor"))?;
    let typed = command["policy"] == "typed-text-v1";
    let fields: &[&str] = if typed {
        &[
            "type",
            "id",
            "sha256",
            "count",
            "bytes",
            "shape",
            "policy",
            "schema",
            "conversion",
        ]
    } else {
        &["type", "id", "sha256", "count", "bytes", "shape", "policy"]
    };
    if object.len() != fields.len()
        || object.keys().any(|key| !fields.contains(&key.as_str()))
        || command["type"] != "importCells"
        || !["grid", "text", "typed-text-v1"].contains(&command["policy"].as_str().unwrap_or(""))
        || typed && !format_matches!(command["schema"], SCHEMA_STAGED_PASTE_V2)
    {
        return Err(crate::coded(
            "workbook-malformed",
            "Invalid staged import descriptor",
        ));
    }
    let shape = command["shape"]
        .as_object()
        .ok_or_else(|| crate::coded("object-malformed", "Invalid staged import shape"))?;
    if shape.len() != 4
        || shape
            .keys()
            .any(|key| !["sheet", "start", "rows", "columns"].contains(&key.as_str()))
    {
        return Err(crate::coded(
            "object-malformed",
            "Invalid staged import shape",
        ));
    }
    let sheet = command["shape"]["sheet"]
        .as_str()
        .ok_or_else(|| crate::coded("workbook-malformed", "Invalid staged import sheet"))?;
    if sheet.is_empty()
        || sheet.encode_utf16().count() > 31
        || sheet
            .chars()
            .any(|c| c.is_control() || "\\/:?*[]".contains(c))
    {
        return Err(crate::coded(
            "workbook-malformed",
            "Invalid staged import sheet",
        ));
    }
    let start = command["shape"]["start"]
        .as_str()
        .ok_or_else(|| crate::coded("workbook-malformed", "Invalid staged import start"))?;
    let split = start
        .find(|c: char| !c.is_ascii_uppercase())
        .ok_or_else(|| crate::coded("workbook-malformed", "Invalid staged import start"))?;
    if split == 0 || split > 3 {
        return Err(crate::coded(
            "workbook-malformed",
            "Invalid staged import start",
        ));
    }
    let mut column = 0u64;
    for c in start[..split].bytes() {
        column = column * 26 + (c - b'A' + 1) as u64;
    }
    let raw_row = &start[split..];
    if raw_row.is_empty()
        || raw_row.starts_with('0')
        || !raw_row.bytes().all(|c| c.is_ascii_digit())
    {
        return Err(crate::coded(
            "workbook-malformed",
            "Invalid staged import start",
        ));
    }
    let row: u64 = raw_row
        .parse()
        .map_err(|_| crate::coded("workbook-malformed", "Invalid staged import start"))?;
    let rows = command["shape"]["rows"]
        .as_u64()
        .ok_or_else(|| crate::coded("workbook-malformed", "Invalid staged import rows"))?;
    let columns = command["shape"]["columns"]
        .as_u64()
        .ok_or_else(|| crate::coded("workbook-malformed", "Invalid staged import columns"))?;
    let positions = rows
        .checked_mul(columns)
        .ok_or_else(|| crate::coded("object-malformed", "Invalid staged import shape"))?;
    let count = command["count"]
        .as_u64()
        .ok_or_else(|| crate::coded("workbook-malformed", "Invalid staged import count"))?;
    if rows == 0
        || columns == 0
        || positions > 1_000_000
        || count == 0
        || count > positions
        || row == 0
        || column == 0
        || row.checked_add(rows - 1).is_none_or(|n| n > 1_048_576)
        || column.checked_add(columns - 1).is_none_or(|n| n > 16384)
    {
        return Err(crate::coded(
            "workbook-limit",
            "Staged import exceeds its worksheet bounds",
        ));
    }
    if typed {
        if count != positions {
            return Err(crate::coded(
                "object-bad-input",
                "Typed query attachment must cover its entire shape",
            ));
        }
        typed_paste_conversion(&command["conversion"], rows, columns)?;
    }
    Ok((rows, columns, count))
}

/// Checks a staged paste attachment against its descriptor, line by line: the
/// header binds it to the document and parent state, and every cell must lie
/// inside the declared shape. The caller owns the file and its identity checks.
pub fn verify_staged_paste<R: BufRead>(
    reader: &mut R,
    command: &Value,
    document: &str,
    parent: Option<&str>,
) -> Result<(), String> {
    let (rows, columns, count) = staged_paste_shape(command)?;
    let mut raw = Vec::new();
    reader
        .by_ref()
        .take(4097)
        .read_until(b'\n', &mut raw)
        .map_err(|e| e.to_string())?;
    if raw.len() > 4096 || raw.last() != Some(&b'\n') {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid bounded paste attachment header",
        ));
    }
    let header: Value = serde_json::from_slice(&raw).map_err(|e| e.to_string())?;
    let typed = command["policy"] == "typed-text-v1";
    let fields: &[&str] = if typed {
        &[
            "schema",
            "documentId",
            "state",
            "shape",
            "count",
            "policy",
            "conversion",
        ]
    } else {
        &["schema", "documentId", "state", "shape", "count", "policy"]
    };
    if keys(&header, fields).is_err()
        || !(if typed {
            format_matches!(header["schema"], SCHEMA_STAGED_PASTE_V2)
        } else {
            format_matches!(header["schema"], SCHEMA_STAGED_PASTE_V1)
        })
        || header["documentId"] != document
        || header["state"] != parent.unwrap_or("base")
        || header["shape"] != command["shape"]
        || header["count"] != command["count"]
        || header["policy"] != command["policy"]
        || typed && header["conversion"] != command["conversion"]
    {
        return Err(crate::coded(
            "session-stale",
            "Workbook attachment belongs to another durable parent",
        ));
    }
    let mut total = 0u64;
    let mut previous = None;
    let mut max_row = 0u64;
    let mut max_column = 0u64;
    loop {
        raw.clear();
        reader
            .by_ref()
            .take(MIB + 1)
            .read_until(b'\n', &mut raw)
            .map_err(|e| e.to_string())?;
        if raw.is_empty() {
            break;
        }
        if raw.len() as u64 > MIB || raw.last() != Some(&b'\n') {
            return Err(crate::coded(
                "workbook-limit",
                "Staged attachment line exceeds its bound",
            ));
        }
        let item: Value = serde_json::from_slice(&raw).map_err(|e| e.to_string())?;
        let item = item
            .as_array()
            .ok_or_else(|| crate::coded("workbook-malformed", "Invalid staged attachment field"))?;
        if item.len() != 2 {
            return Err(crate::coded(
                "workbook-malformed",
                "Invalid staged attachment field",
            ));
        }
        let offset = item[0].as_u64().ok_or_else(|| {
            crate::coded("workbook-malformed", "Invalid staged attachment offset")
        })?;
        let text = item[1]
            .as_str()
            .ok_or_else(|| crate::coded("workbook-malformed", "Invalid staged attachment text"))?;
        if offset >= rows * columns
            || previous.is_some_and(|n| offset <= n)
            || text.encode_utf16().count() > 32767
            || text.chars().any(|c| {
                !matches!(c, '\t' | '\n' | '\r')
                    && (c < '\u{20}' || c == '\u{fffe}' || c == '\u{ffff}')
            })
        {
            return Err(crate::coded(
                "workbook-malformed",
                "Invalid staged attachment ordering or text",
            ));
        }
        previous = Some(offset);
        max_row = max_row.max(offset / columns);
        max_column = max_column.max(offset % columns);
        total += 1;
        if total > count {
            return Err(crate::coded(
                "workbook-limit",
                "Staged attachment exceeds its declared count",
            ));
        }
    }
    if total != count || max_row + 1 != rows || max_column + 1 != columns {
        return Err(crate::coded(
            "object-stale",
            "Staged attachment count or shape changed",
        ));
    }
    Ok(())
}

/// Checks a calculation witness the engine wrote for a saved workbook: its
/// identity, schema and object inventory, bounded metadata, and that it belongs
/// to the workbook output with hash `source_sha256`.
pub fn verify_calculation_witness(path: &Path, source_sha256: &str) -> Result<(), String> {
    let witness = Connection::open_with_flags(
        &path,
        rusqlite::OpenFlags::SQLITE_OPEN_READ_ONLY | rusqlite::OpenFlags::SQLITE_OPEN_NO_MUTEX,
    )
    .map_err(|e| e.to_string())?;
    witness.execute_batch("PRAGMA query_only=ON;PRAGMA trusted_schema=OFF;PRAGMA cache_size=-4096;PRAGMA mmap_size=0;").map_err(|e|e.to_string())?;
    if witness
        .query_row("PRAGMA application_id", [], |row| row.get::<_, u32>(0))
        .map_err(|e| e.to_string())?
        != 0x52504658
        || witness
            .query_row("PRAGMA user_version", [], |row| row.get::<_, u32>(0))
            .map_err(|e| e.to_string())?
            != 1
        || witness
            .query_row("PRAGMA quick_check", [], |row| row.get::<_, String>(0))
            .map_err(|e| e.to_string())?
            != "ok"
    {
        return Err(crate::coded(
            "formula-malformed",
            "Unsupported or damaged native calculation witness",
        ));
    }
    let mut statement = witness
        .prepare("SELECT type,name FROM sqlite_master WHERE name NOT LIKE 'sqlite_%' ORDER BY name")
        .map_err(|e| e.to_string())?;
    let tables = statement
        .query_map([], |row| {
            Ok((row.get::<_, String>(0)?, row.get::<_, String>(1)?))
        })
        .map_err(|e| e.to_string())?
        .collect::<Result<Vec<_>, _>>()
        .map_err(|e| e.to_string())?;
    if tables
        != ["cells", "derived", "groups", "metadata"]
            .map(|name| ("table".to_string(), name.to_string()))
    {
        return Err(crate::coded(
            "object-stale",
            "Native calculation witness object inventory changed",
        ));
    }
    if witness
        .query_row("SELECT count(*) FROM metadata", [], |row| {
            row.get::<_, u64>(0)
        })
        .map_err(|e| e.to_string())?
        != 1
    {
        return Err(crate::coded(
            "formula-refused",
            "Native calculation witness metadata is ambiguous",
        ));
    }
    let (bindings_bytes, policy_bytes): (u64, u64) = witness
        .query_row(
            "SELECT length(bindings),length(policy) FROM metadata WHERE schema=1",
            [],
            |row| Ok((row.get(0)?, row.get(1)?)),
        )
        .map_err(|e| e.to_string())?;
    if bindings_bytes > 64 * MIB || policy_bytes > 4096 {
        return Err(crate::coded(
            "formula-limit",
            "Native calculation witness metadata exceeds its bound",
        ));
    }
    let source: String = witness
        .query_row(
            "SELECT source_revision FROM metadata WHERE schema=1",
            [],
            |row| row.get(0),
        )
        .map_err(|e| e.to_string())?;
    if source != source_sha256 {
        return Err(crate::coded(
            "formula-stale",
            "Native calculation witness belongs to another workbook output",
        ));
    }
    drop(statement);
    drop(witness);
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    fn grid(rows: u64, columns: u64, count: u64) -> Value {
        json!({"type":"importCells","id":"p","sha256":"h","count":count,"bytes":1,
            "shape":{"sheet":"Data","start":"B2","rows":rows,"columns":columns},"policy":"grid"})
    }
    fn attachment(document: &str, state: &str, command: &Value, cells: &[(u64, &str)]) -> Vec<u8> {
        let header = json!({"schema":workbook_format!(SCHEMA_STAGED_PASTE_V1),"documentId":document,"state":state,
            "shape":command["shape"],"count":command["count"],"policy":command["policy"]});
        let mut bytes = serde_json::to_vec(&header).unwrap();
        bytes.push(b'\n');
        for (offset, text) in cells {
            bytes.extend(serde_json::to_vec(&json!([offset, text])).unwrap());
            bytes.push(b'\n');
        }
        bytes
    }

    #[test]
    fn paste_shape_is_bounded_by_the_worksheet() {
        assert_eq!(staged_paste_shape(&grid(2, 3, 4)).unwrap(), (2, 3, 4));
        let mut command = grid(2, 3, 4);
        command["shape"]["sheet"] = json!("a/b");
        assert!(staged_paste_shape(&command).is_err());
        let mut command = grid(2, 3, 4);
        command["shape"]["start"] = json!("A0");
        assert!(staged_paste_shape(&command).is_err());
        assert!(staged_paste_shape(&grid(1_048_576, 1, 1)).is_err());
        assert!(staged_paste_shape(&grid(2, 3, 7)).is_err());
    }

    #[test]
    fn paste_attachment_belongs_to_its_document_and_parent() {
        let command = grid(2, 2, 2);
        let good = attachment("doc", "base", &command, &[(0, "a"), (3, "b")]);
        verify_staged_paste(&mut good.as_slice(), &command, "doc", None).unwrap();
        assert!(verify_staged_paste(&mut good.as_slice(), &command, "other", None).is_err());
        assert!(verify_staged_paste(&mut good.as_slice(), &command, "doc", Some("head")).is_err());
        let unordered = attachment("doc", "base", &command, &[(3, "b"), (0, "a")]);
        assert!(verify_staged_paste(&mut unordered.as_slice(), &command, "doc", None).is_err());
        let short = attachment("doc", "base", &command, &[(0, "a")]);
        assert!(verify_staged_paste(&mut short.as_slice(), &command, "doc", None).is_err());
    }

    #[test]
    fn the_helper_protocol_reports_the_hash_of_exactly_the_bytes_it_checked() {
        use sha2::Digest;
        let command = grid(2, 2, 2);
        let bytes = attachment("doc", "head", &command, &[(0, "a"), (3, "d")]);
        let directory =
            std::env::temp_dir().join(format!("workbook-journal-{}", std::process::id()));
        std::fs::create_dir_all(&directory).unwrap();
        let path = directory.join("paste.jsonl");
        std::fs::write(&path, &bytes).unwrap();
        let input = |document: &str| {
            json!({"action":"journal","operation":"verify-staged-paste",
            "path":path.to_str().unwrap(),"command":command,"document":document,"parent":"head"})
        };
        let answer = crate::command(input("doc")).unwrap();
        assert_eq!(
            answer["sha256"],
            format!("{:x}", sha2::Sha256::digest(&bytes))
        );
        assert_eq!(answer["bytes"], bytes.len());
        assert!(
            crate::command(input("other")).is_err(),
            "A paste for another document is refused"
        );
        let missing = json!({"action":"journal","operation":"verify-calculation-witness","path":directory.join("absent").to_str().unwrap(),"sourceSha256":"x"});
        assert!(crate::command(missing).is_err());
        std::fs::remove_dir_all(directory).unwrap();
    }
    #[test]
    fn calculation_witness_is_identified_and_bound_to_its_output() {
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("witness.sqlite");
        let db = Connection::open(&path).unwrap();
        db.execute_batch(
            "PRAGMA application_id=0x52504658;PRAGMA user_version=1;
             CREATE TABLE cells(x);CREATE TABLE derived(x);CREATE TABLE groups(x);
             CREATE TABLE metadata(schema,bindings,policy,source_revision);
             INSERT INTO metadata VALUES(1,'[]','{}','abc');",
        )
        .unwrap();
        drop(db);
        verify_calculation_witness(&path, "abc").unwrap();
        assert!(verify_calculation_witness(&path, "other").is_err());
        let db = Connection::open(&path).unwrap();
        db.execute_batch("PRAGMA application_id=1;").unwrap();
        drop(db);
        assert!(verify_calculation_witness(&path, "abc").is_err());
    }
}

/// Reads `path` through a SHA-256 counter, so the bytes checked are the bytes hashed.
struct Counted<R> {
    inner: R,
    digest: sha2::Sha256,
    bytes: u64,
}
impl<R: Read> Read for Counted<R> {
    fn read(&mut self, buffer: &mut [u8]) -> std::io::Result<usize> {
        use sha2::Digest;
        let read = self.inner.read(buffer)?;
        self.digest.update(&buffer[..read]);
        self.bytes += read as u64;
        Ok(read)
    }
}

/// The journal checks over the helper protocol, for a host that runs the
/// engine program instead of linking this crate (`{"action":"journal",...}`).
///
/// - `verify-staged-paste` `{path, command, document, parent}`: checks the file
///   as `verify_staged_paste` does, reading it to its end, and answers with the
///   SHA-256 and size of exactly the bytes it read. The host compares them with
///   the identity it stored and checks the file did not change meanwhile.
/// - `verify-calculation-witness` `{path, sourceSha256}`: as
///   `verify_calculation_witness`.
pub fn command(input: &Value) -> Result<Value, String> {
    let path = Path::new(crate::field(input, "path")?);
    match crate::field(input, "operation")? {
        "verify-staged-paste" => {
            use sha2::Digest;
            let file = std::fs::File::open(path).map_err(|e| e.to_string())?;
            let mut reader = std::io::BufReader::new(Counted {
                inner: file,
                digest: sha2::Sha256::new(),
                bytes: 0,
            });
            let parent = input["parent"].as_str();
            verify_staged_paste(
                &mut reader,
                &input["command"],
                crate::field(input, "document")?,
                parent,
            )?;
            std::io::copy(&mut reader, &mut std::io::sink()).map_err(|e| e.to_string())?;
            let counted = reader.into_inner();
            Ok(
                serde_json::json!({"sha256":format!("{:x}", counted.digest.finalize()),"bytes":counted.bytes}),
            )
        }
        "verify-calculation-witness" => {
            verify_calculation_witness(path, crate::field(input, "sourceSha256")?)?;
            Ok(serde_json::json!({"verified":true}))
        }
        _ => Err(crate::coded(
            "internal-invalid-field",
            "Unknown journal operation",
        )),
    }
}
