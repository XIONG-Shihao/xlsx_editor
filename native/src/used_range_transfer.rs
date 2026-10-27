//! Native extent authority across a single private package rewrite. An export
//! input is not a rewrite receipt: only a qualified writer can finish one, and
//! it is bound to one sealed candidate directory rather than merely its bytes.
use crate::used_range::Bounds;
use crate::{Result, err, field, hash, local};
use quick_xml::{
    NsReader,
    events::Event,
    name::{Namespace, ResolveResult},
};
use rusqlite::{Connection, OpenFlags, OptionalExtension, params};
use serde::{Deserialize, Serialize};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};
use std::{
    collections::{BTreeMap, BTreeSet},
    fs::{self, File},
    io::{Read, Write},
    path::{Path, PathBuf},
};

const SEAL_FILE: &str = "used-range-rewrite-seal.json";
const MAX_RECORD: usize = 64 * 1024 * 1024;
const MAIN_NS: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_NS: &[u8] = b"http://schemas.openxmlformats.org/package/2006/relationships";

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub(crate) struct SheetBinding {
    pub name: String,
    pub path: String,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub(crate) enum WriterKind {
    Package,
    Insertion,
    Transform,
    Layout,
    Dedup,
}
impl WriterKind {
    fn file(self) -> &'static str {
        match self {
            Self::Package => "package.xlsx",
            Self::Insertion => "insertion.xlsx",
            Self::Transform => "transform.xlsx",
            Self::Layout => "layout.xlsx",
            Self::Dedup => "dedup.xlsx",
        }
    }
}
#[derive(Clone, Copy, Debug, Serialize)]
#[serde(rename_all = "lowercase")]
pub(crate) enum Axis {
    Row,
    Column,
}
/// Constructed by a qualified Rust writer, never decoded from a transfer RPC.
#[derive(Debug, Serialize)]
pub(crate) enum NativeMapping {
    Identity,
    Axis {
        sheet: String,
        axis: Axis,
        index: u32,
        count: u32,
        remove: bool,
    },
    Rows {
        sheet: String,
        first: u32,
        last: u32,
        rows: Vec<(u32, u32)>,
    },
    Compact {
        sheet: String,
        first: (u32, u32),
        last: (u32, u32),
        body: u32,
        rows: Vec<(u32, Option<u32>)>,
    },
    Sheets {
        predecessors: BTreeMap<String, Option<String>>,
    },
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct Extent {
    sheet: String,
    floor: Option<Bounds>,
    reasons: Vec<Value>,
}
#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct Seal {
    schema: String,
    input_id: String,
    nonce: String,
    source_root: String,
    target_root: String,
    owner_revision: String,
}
#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct InputRecord {
    id: String,
    source_root: String,
    target_root: String,
    owner_revision: String,
    target_identity: Option<(u64, u64)>,
    seal_revision: String,
    version: u64,
    source_revision: String,
    index_revision: String,
    metadata_revision: String,
    input_file: String,
    input_revision: String,
    sheets: Vec<SheetBinding>,
    extents: Vec<Extent>,
}
#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct OutputRecord {
    id: String,
    input: InputRecord,
    writer: WriterKind,
    output_file: String,
    output_revision: String,
    mapping_revision: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    clear_hyperlinks_command: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    hyperlink_authoring_command: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    page_theme_command: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    background_command: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    note_authoring_command: Option<Value>,
    sheets: Vec<SheetBinding>,
    predecessors: BTreeMap<String, Option<String>>,
    extents: Vec<Extent>,
}
pub(crate) struct ExportCandidate {
    record: InputRecord,
    seal_active: bool,
}
pub(crate) struct WriterInput {
    record: InputRecord,
    kind: WriterKind,
}
pub(crate) struct PreparedOutput {
    record: OutputRecord,
}

fn sha(bytes: &[u8]) -> String {
    format!("{:x}", Sha256::digest(bytes))
}
fn encoded<T: Serialize>(value: &T) -> Result<String> {
    let text = serde_json::to_string(value).map_err(err)?;
    if text.len() > MAX_RECORD {
        return Err(crate::coded(
            "layout-limit",
            "Private extent receipt exceeds its bound",
        ));
    }
    Ok(text)
}
fn uuid(value: &str) -> bool {
    value.len() == 36
        && value.as_bytes()[14] == b'4'
        && matches!(value.as_bytes()[19], b'8' | b'9' | b'a' | b'b')
        && value.bytes().enumerate().all(|(i, b)| {
            if [8, 13, 18, 23].contains(&i) {
                b == b'-'
            } else {
                b.is_ascii_digit() || (b'a'..=b'f').contains(&b)
            }
        })
}
fn new_id(db: &Connection) -> Result<String> {
    let mut bytes: Vec<u8> = db
        .query_row("SELECT randomblob(16)", [], |r| r.get(0))
        .map_err(err)?;
    bytes[6] = (bytes[6] & 0x0f) | 0x40;
    bytes[8] = (bytes[8] & 0x3f) | 0x80;
    let text = bytes.iter().map(|b| format!("{b:02x}")).collect::<String>();
    Ok(format!(
        "{}-{}-{}-{}-{}",
        &text[..8],
        &text[8..12],
        &text[12..16],
        &text[16..20],
        &text[20..]
    ))
}
fn revision(value: &str) -> bool {
    value.len() == 64
        && value
            .bytes()
            .all(|b| b.is_ascii_digit() || (b'a'..=b'f').contains(&b))
}
fn version(input: &Value, key: &str) -> Result<u64> {
    input[key]
        .as_u64()
        .filter(|v| *v <= 9_007_199_254_740_991)
        .ok_or_else(|| {
            crate::coded(
                "internal-invalid-field",
                format!("Invalid private extent {key}"),
            )
        })
}
fn known_version(db: &Connection, value: u64) -> Result<()> {
    if !db
        .query_row(
            "SELECT EXISTS(SELECT 1 FROM versions WHERE id=?1)",
            [value],
            |r| r.get::<_, bool>(0),
        )
        .map_err(err)?
    {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown private extent version",
        ));
    }
    Ok(())
}
fn closed(input: &Value, required: &[&str]) -> Result<()> {
    let object = input
        .as_object()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid private extent request"))?;
    if object.get("diagnostics").is_some_and(|v| !v.is_boolean()) {
        return Err(crate::coded(
            "layout-unsupported",
            "Private extent diagnostics must be a boolean",
        ));
    }
    let actual = object
        .keys()
        .filter(|k| k.as_str() != "diagnostics")
        .map(String::as_str)
        .collect::<BTreeSet<_>>();
    if actual != required.iter().copied().collect() {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid private extent request fields",
        ));
    }
    Ok(())
}
fn regular(path: &Path, max: Option<u64>) -> Result<()> {
    let stat = fs::symlink_metadata(path).map_err(err)?;
    if !stat.is_file()
        || stat.file_type().is_symlink()
        || max.is_some_and(|limit| stat.len() > limit)
    {
        return Err(crate::coded(
            "layout-unsupported",
            "Unsafe private extent artifact",
        ));
    }
    #[cfg(unix)]
    {
        use std::os::unix::fs::MetadataExt;
        if stat.nlink() != 1 {
            return Err(crate::coded(
                "layout-unsupported",
                "Private extent artifact has multiple links",
            ));
        }
    }
    Ok(())
}
fn canonical_directory(path: &Path) -> Result<PathBuf> {
    if !path.is_absolute() {
        return Err(crate::coded(
            "layout-unsupported",
            "Private extent root must be absolute",
        ));
    }
    let stat = fs::symlink_metadata(path).map_err(err)?;
    if !stat.is_dir() || stat.file_type().is_symlink() {
        return Err(crate::coded(
            "layout-unsupported",
            "Unsafe private extent root",
        ));
    }
    fs::canonicalize(path).map_err(err)
}
fn directory_identity(path: &Path) -> Result<Option<(u64, u64)>> {
    #[cfg(unix)]
    {
        use std::os::unix::fs::MetadataExt;
        let stat = fs::symlink_metadata(path).map_err(err)?;
        Ok(Some((stat.dev(), stat.ino())))
    }
    #[cfg(not(unix))]
    {
        let _ = path;
        Ok(None)
    }
}
fn owner_revision(root: &Path, required: bool) -> Result<String> {
    let parent = root
        .parent()
        .ok_or_else(|| crate::coded("layout-unsupported", "Private extent root lacks an owner"))?;
    let marker = parent.join("owner.json");
    match fs::symlink_metadata(&marker) {
        Ok(_) => {}
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => {
            if required {
                return Err(crate::coded(
                    "layout-unsupported",
                    "Private extent rewrite requires a native owner marker",
                ));
            }
            return Ok(String::new());
        }
        Err(error) => return Err(err(error)),
    }
    regular(&marker, Some(4096))?;
    let bytes = fs::read(&marker).map_err(err)?;
    let value: Value = serde_json::from_slice(&bytes).map_err(err)?;
    let keys = value
        .as_object()
        .ok_or_else(|| {
            crate::coded(
                "internal-invalid-field",
                "Invalid private extent owner marker",
            )
        })?
        .keys()
        .map(String::as_str)
        .collect::<BTreeSet<_>>();
    if keys != ["schema", "pid", "sessionId", "maxSessionBytes"].into()
        || !format_matches!(value["schema"], WORKBOOK_INDEX_OWNER_V1)
        || value["pid"]
            .as_u64()
            .is_none_or(|v| v == 0 || v > 9_007_199_254_740_991)
        || value["sessionId"]
            .as_str()
            .is_none_or(|v| v.is_empty() || v.len() > 128 || v.contains('\0'))
        || value["maxSessionBytes"]
            .as_u64()
            .is_none_or(|v| !(128 * 1024..=4 * 1024 * 1024 * 1024).contains(&v))
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid private extent owner marker",
        ));
    }
    Ok(sha(&bytes))
}
fn owner_identity(root: &Path, required: bool) -> Result<String> {
    let marker_revision = owner_revision(root, required)?;
    if marker_revision.is_empty() {
        return Ok(String::new());
    }
    let parent = root
        .parent()
        .ok_or_else(|| crate::coded("layout-unsupported", "Private extent root lacks an owner"))?;
    let value: Value =
        serde_json::from_slice(&fs::read(parent.join("owner.json")).map_err(err)?).map_err(err)?;
    Ok(format!(
        "{}:{}",
        parent.to_string_lossy(),
        field(&value, "sessionId")?
    ))
}
fn read_db(root: &Path) -> Result<Connection> {
    regular(&local(root, "cells.sqlite")?, None)?;
    let db = Connection::open_with_flags(
        local(root, "cells.sqlite")?,
        OpenFlags::SQLITE_OPEN_READ_ONLY | OpenFlags::SQLITE_OPEN_NO_MUTEX,
    )
    .map_err(err)?;
    db.execute_batch("PRAGMA cache_size=-4096;PRAGMA temp_store=FILE;PRAGMA mmap_size=0;")
        .map_err(err)?;
    db.progress_handler(10000, Some(crate::cancel::interrupted));
    Ok(db)
}
fn index(root: &Path) -> Result<Value> {
    let path = local(root, "index.json")?;
    regular(&path, Some(16 * 1024 * 1024))?;
    serde_json::from_slice(&fs::read(path).map_err(err)?).map_err(err)
}
fn table(db: &Connection, name: &str) -> Result<bool> {
    db.prepare_cached("SELECT EXISTS(SELECT 1 FROM sqlite_master WHERE type='table' AND name=?1)")
        .map_err(crate::err)?
        .query_row([name], |r| r.get(0))
        .map_err(err)
}
fn binding_schema(db: &Connection) -> Result<()> {
    db.execute_batch("CREATE TABLE IF NOT EXISTS used_range_index_binding(singleton INTEGER PRIMARY KEY CHECK(singleton=1),root TEXT NOT NULL,owner TEXT NOT NULL,source_revision TEXT NOT NULL,state TEXT NOT NULL CHECK(state IN('live','source-snapshot')));").map_err(err)
}
pub(crate) fn bind_import(root: &Path, db: &Connection, source_revision: &str) -> Result<()> {
    if !revision(source_revision) {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid imported extent source revision",
        ));
    }
    binding_schema(db)?;
    let root = canonical_directory(root)?;
    let owner = owner_identity(&root, false)?;
    db.execute(
        "INSERT INTO used_range_index_binding VALUES(1,?1,?2,?3,'live')",
        params![root.to_string_lossy(), owner, source_revision],
    )
    .map_err(err)?;
    Ok(())
}
fn binding(db: &Connection) -> Result<(String, String, String, String)> {
    if !table(db, "used_range_index_binding")? {
        return Err(crate::coded(
            "layout-unsupported",
            "Private extent index lacks its immutable owner binding; reimport the workbook",
        ));
    }
    db.prepare_cached(
        "SELECT root,owner,source_revision,state FROM used_range_index_binding WHERE singleton=1",
    )
    .map_err(crate::err)?
    .query_row([], |r| Ok((r.get(0)?, r.get(1)?, r.get(2)?, r.get(3)?)))
    .map_err(err)
}
fn live_binding(root: &Path, db: &Connection, source_revision: &str) -> Result<()> {
    let value = binding(db)?;
    let root = canonical_directory(root)?;
    if value.0 != root.to_string_lossy()
        || value.1 != owner_identity(&root, true)?
        || value.2 != source_revision
        || value.3 != "live"
    {
        return Err(crate::coded(
            "layout-stale",
            "Private extent index owner binding is stale",
        ));
    }
    Ok(())
}
fn source_check(
    root: &Path,
    db: &Connection,
    selected: u64,
    source_revision: &str,
) -> Result<Value> {
    known_version(db, selected)?;
    let index = index(root)?;
    regular(&local(root, "source.xlsx")?, None)?;
    if !revision(source_revision)
        || index["revision"] != source_revision
        || hash(&local(root, "source.xlsx")?)? != source_revision
    {
        return Err(crate::coded(
            "layout-stale",
            "Private extent source revision is stale",
        ));
    }
    Ok(index)
}

fn part(zip: &mut zip::ZipArchive<File>, name: &str) -> Result<Vec<u8>> {
    let mut member = zip.by_name(name).map_err(err)?;
    if member.size() > 4 * 1024 * 1024 {
        return Err(crate::coded(
            "layout-limit",
            "Private extent worksheet inventory exceeds its bound",
        ));
    }
    let mut bytes = Vec::new();
    Read::by_ref(&mut member)
        .take(4 * 1024 * 1024 + 1)
        .read_to_end(&mut bytes)
        .map_err(err)?;
    if bytes.len() > 4 * 1024 * 1024 {
        return Err(crate::coded(
            "layout-limit",
            "Private extent worksheet inventory exceeds its bound",
        ));
    }
    Ok(bytes)
}
fn resolve_sheet(target: &str) -> Result<String> {
    if target.contains(['\\', ':', '?', '#', '\0']) {
        return Err(crate::coded(
            "layout-unsupported",
            "Opaque private extent worksheet target",
        ));
    }
    let mut parts = if target.starts_with('/') {
        Vec::new()
    } else {
        vec!["xl"]
    };
    for p in target.trim_start_matches('/').split('/') {
        match p {
            "." => {}
            ".." => {
                parts.pop().ok_or_else(|| {
                    crate::coded(
                        "layout-unsupported",
                        "Private extent worksheet escapes package",
                    )
                })?;
            }
            "" => {
                return Err(crate::coded(
                    "layout-malformed",
                    "Invalid private extent worksheet path",
                ));
            }
            _ => parts.push(p),
        }
    }
    Ok(parts.join("/"))
}
/// Resolve the actual canonical workbook/relationship graph, independently of
/// the writer's plan and the target's caller-supplied metadata projection.
pub(crate) fn worksheet_inventory(file: &Path) -> Result<Vec<SheetBinding>> {
    regular(file, None)?;
    let mut zip = zip::ZipArchive::new(File::open(file).map_err(err)?).map_err(err)?;
    let names = zip.file_names().map(str::to_owned).collect::<BTreeSet<_>>();
    if names.len() != zip.len()
        || names.iter().any(|name| {
            name.starts_with('/')
                || name.contains('\\')
                || name
                    .trim_end_matches('/')
                    .split('/')
                    .any(|p| p.is_empty() || p == "." || p == "..")
        })
    {
        return Err(crate::coded(
            "layout-ambiguous",
            "Ambiguous private extent package inventory",
        ));
    }
    let bytes = part(&mut zip, "xl/_rels/workbook.xml.rels")?;
    let mut reader = NsReader::from_reader(bytes.as_slice());
    let mut buffer = Vec::new();
    let mut relationships = BTreeMap::new();
    let mut depth = 0usize;
    let mut root_seen = false;
    loop {
        crate::cancel::check()?;
        let (ns, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
        let main = ns == ResolveResult::Bound(Namespace(PACKAGE_NS));
        match event {
            Event::Start(ref e) | Event::Empty(ref e) => {
                let empty = matches!(event, Event::Empty(_));
                if depth == 0 {
                    if !main || e.local_name().as_ref() != b"Relationships" || root_seen {
                        return Err(crate::coded(
                            "layout-malformed",
                            "Invalid private extent relationship root",
                        ));
                    }
                    root_seen = true;
                } else if depth == 1 {
                    if !main || e.local_name().as_ref() != b"Relationship" {
                        return Err(crate::coded(
                            "layout-unsupported",
                            "Opaque private extent worksheet relationship",
                        ));
                    }
                    let a = crate::xml::attrs(e)?;
                    let id = a.get("Id").filter(|v| !v.is_empty()).ok_or_else(|| {
                        crate::coded(
                            "internal-missing-field",
                            "Missing private extent relationship ID",
                        )
                    })?;
                    let kind = a.get("Type").ok_or_else(|| {
                        crate::coded(
                            "internal-missing-field",
                            "Missing private extent relationship kind",
                        )
                    })?;
                    let path = if kind
                        == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
                    {
                        if a.get("TargetMode").is_some_and(|v| v != "Internal") {
                            return Err(crate::coded(
                                "layout-unsupported",
                                "External private extent worksheet",
                            ));
                        }
                        Some(resolve_sheet(a.get("Target").ok_or_else(|| {
                            crate::coded(
                                "internal-missing-field",
                                "Missing private extent worksheet target",
                            )
                        })?)?)
                    } else {
                        None
                    };
                    if relationships.insert(id.clone(), path).is_some() {
                        return Err(crate::coded(
                            "layout-duplicate",
                            "Duplicate private extent relationship ID",
                        ));
                    }
                } else {
                    return Err(crate::coded(
                        "layout-unsupported",
                        "Nested private extent relationship",
                    ));
                }
                if !empty {
                    depth += 1;
                }
            }
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    crate::coded(
                        "layout-malformed",
                        "Invalid private extent relationship nesting",
                    )
                })?;
            }
            Event::Text(ref t) if !t.unescape().map_err(err)?.trim().is_empty() => {
                return Err(crate::coded(
                    "layout-unsupported",
                    "Opaque private extent relationship text",
                ));
            }
            Event::DocType(_) => {
                return Err(crate::coded(
                    "layout-malformed",
                    "Private extent XML entities are unsupported",
                ));
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if !root_seen || depth != 0 {
        return Err(crate::coded(
            "layout-unsupported",
            "Incomplete private extent relationships",
        ));
    }
    let bytes = part(&mut zip, "xl/workbook.xml")?;
    let mut reader = NsReader::from_reader(bytes.as_slice());
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut sheets_depth = None;
    let mut seen = false;
    let mut container = false;
    let (mut result, mut paths, mut labels) = (Vec::new(), BTreeSet::new(), BTreeSet::new());
    loop {
        crate::cancel::check()?;
        let (ns, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
        let main = ns == ResolveResult::Bound(Namespace(MAIN_NS));
        match event {
            Event::Start(ref e) | Event::Empty(ref e) => {
                let empty = matches!(event, Event::Empty(_));
                let name = e.local_name();
                if depth == 0 {
                    if !main || name.as_ref() != b"workbook" || seen {
                        return Err(crate::coded(
                            "layout-malformed",
                            "Invalid private extent workbook root",
                        ));
                    }
                    seen = true;
                }
                if main && depth == 1 && name.as_ref() == b"sheets" {
                    if container {
                        return Err(crate::coded(
                            "layout-duplicate",
                            "Duplicate private extent sheet inventory",
                        ));
                    }
                    container = true;
                    sheets_depth = Some(depth);
                } else if sheets_depth.is_some_and(|d| depth == d + 1) {
                    if !main || name.as_ref() != b"sheet" {
                        return Err(crate::coded(
                            "layout-unsupported",
                            "Opaque private extent sheet inventory",
                        ));
                    }
                    let a = crate::xml::attrs(e)?;
                    let label = a
                        .get("name")
                        .filter(|v| !v.is_empty() && v.len() <= 124)
                        .ok_or_else(|| {
                            crate::coded(
                                "internal-invalid-field",
                                "Invalid private extent sheet name",
                            )
                        })?;
                    let mut relation = None;
                    for attr in e.attributes().with_checks(true) {
                        let attr = attr.map_err(err)?;
                        let (binding, name) = reader.resolve_attribute(attr.key);
                        if binding == ResolveResult::Bound(Namespace(REL_NS))
                            && name.as_ref() == b"id"
                        {
                            relation = Some(attr.unescape_value().map_err(err)?.into_owned());
                        }
                    }
                    let path = relationships
                        .get(&relation.ok_or_else(|| {
                            crate::coded(
                                "internal-missing-field",
                                "Missing private extent sheet relationship",
                            )
                        })?)
                        .and_then(Option::as_ref)
                        .ok_or_else(|| {
                            crate::coded(
                                "layout-unsupported",
                                "Unqualified private extent sheet relationship",
                            )
                        })?;
                    if !names.contains(path)
                        || !paths.insert(path.clone())
                        || !labels.insert(label.to_lowercase())
                    {
                        return Err(crate::coded(
                            "layout-ambiguous",
                            "Ambiguous private extent sheet graph",
                        ));
                    }
                    result.push(SheetBinding {
                        name: label.clone(),
                        path: path.clone(),
                    });
                    if result.len() > 1024 {
                        return Err(crate::coded(
                            "layout-limit",
                            "Private extent worksheet inventory exceeds its bound",
                        ));
                    }
                } else if sheets_depth.is_some_and(|d| depth > d + 1) {
                    return Err(crate::coded(
                        "layout-unsupported",
                        "Nested private extent sheet inventory",
                    ));
                }
                if !empty {
                    depth += 1;
                } else if sheets_depth == Some(depth) && name.as_ref() == b"sheets" {
                    sheets_depth = None;
                }
            }
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    crate::coded(
                        "internal-invalid-field",
                        "Invalid private extent workbook nesting",
                    )
                })?;
                if sheets_depth == Some(depth) {
                    sheets_depth = None;
                }
            }
            Event::Text(ref t)
                if sheets_depth.is_some() && !t.unescape().map_err(err)?.trim().is_empty() =>
            {
                return Err(crate::coded(
                    "layout-unsupported",
                    "Opaque private extent sheet text",
                ));
            }
            Event::DocType(_) => {
                return Err(crate::coded(
                    "layout-malformed",
                    "Private extent XML entities are unsupported",
                ));
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if !seen || depth != 0 || !container || result.is_empty() {
        return Err(crate::coded(
            "layout-unsupported",
            "Incomplete private extent sheet inventory",
        ));
    }
    Ok(result)
}

fn receipt_schema(db: &Connection) -> Result<()> {
    db.execute_batch("CREATE TABLE IF NOT EXISTS export_rewrite_inputs(singleton INTEGER PRIMARY KEY CHECK(singleton=1),id TEXT NOT NULL UNIQUE,version INTEGER NOT NULL,data TEXT NOT NULL);CREATE TABLE IF NOT EXISTS export_rewrite_outputs(id TEXT PRIMARY KEY,input_id TEXT NOT NULL,version INTEGER NOT NULL,output_file TEXT NOT NULL UNIQUE,data TEXT NOT NULL) WITHOUT ROWID;").map_err(err)
}
fn check_seal(record: &InputRecord) -> Result<()> {
    let source = canonical_directory(Path::new(&record.source_root))?;
    let target = canonical_directory(Path::new(&record.target_root))?;
    if source == target
        || source.parent() != target.parent()
        || owner_revision(&source, true)? != record.owner_revision
        || directory_identity(&target)? != record.target_identity
    {
        return Err(crate::coded(
            "layout-stale",
            "Private extent candidate identity is stale",
        ));
    }
    let file = local(&target, SEAL_FILE)?;
    regular(&file, Some(4096))?;
    let bytes = fs::read(&file).map_err(err)?;
    if sha(&bytes) != record.seal_revision {
        return Err(crate::coded(
            "layout-stale",
            "Private extent candidate seal changed",
        ));
    }
    let seal: Seal = serde_json::from_slice(&bytes).map_err(err)?;
    if !format_matches!(seal.schema, SCHEMA_USED_RANGE_REWRITE_SEAL_V1)
        || seal.input_id != record.id
        || !uuid(&seal.nonce)
        || seal.source_root != record.source_root
        || seal.target_root != record.target_root
        || seal.owner_revision != record.owner_revision
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Private extent candidate seal differs from native reservation",
        ));
    }
    Ok(())
}
fn check_input(root: &Path, db: &Connection, record: &InputRecord) -> Result<()> {
    if canonical_directory(root)?.to_string_lossy() != record.source_root
        || !uuid(&record.id)
        || record.input_file != "export.xlsx"
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Private extent input root differs from native receipt",
        ));
    }
    source_check(root, db, record.version, &record.source_revision)?;
    live_binding(root, db, &record.source_revision)?;
    check_seal(record)?;
    if hash(&local(root, "index.json")?)? != record.index_revision
        || hash(&local(root, "metadata.xlsx")?)? != record.metadata_revision
        || hash(&local(root, &record.input_file)?)? != record.input_revision
        || worksheet_inventory(&local(root, &record.input_file)?)? != record.sheets
    {
        return Err(crate::coded(
            "layout-stale",
            "Private extent input package/metadata changed",
        ));
    }
    Ok(())
}
impl Drop for ExportCandidate {
    fn drop(&mut self) {
        if self.seal_active {
            let file = Path::new(&self.record.target_root).join(SEAL_FILE);
            if regular(&file, Some(4096)).is_ok()
                && fs::read(&file).is_ok_and(|bytes| sha(&bytes) == self.record.seal_revision)
            {
                let _ = fs::remove_file(file);
            }
        }
    }
}
pub(crate) fn prepare_export(
    root: &Path,
    db: &Connection,
    input: &Value,
) -> Result<Option<ExportCandidate>> {
    if input
        .get("internalRewrite")
        .is_some_and(|v| !v.is_boolean())
    {
        return Err(crate::coded(
            "layout-unsupported",
            "Private extent internal rewrite flag must be boolean",
        ));
    }
    if input["internalRewrite"] != true {
        if input.get("rewriteTargetRoot").is_some() {
            return Err(crate::coded(
                "layout-unsupported",
                "Private extent target requires an internal rewrite export",
            ));
        }
        return Ok(None);
    }
    let mut fields = vec![
        "action",
        "operation",
        "root",
        "revision",
        "version",
        "calculationMetadata",
        "internalRewrite",
        "rewriteTargetRoot",
    ];
    if input.get("preservationExport").is_some() {
        fields.push("preservationExport");
    }
    closed(input, &fields)?;
    if input["action"] != "workbook" || input["operation"] != "export" {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid private extent export operation",
        ));
    }
    if input["publication"] == true || input["calculationMetadata"] != true {
        return Err(crate::coded(
            "layout-unsupported",
            "Private extent input requires prepared calculation export, separate from publication",
        ));
    }
    let selected = version(input, "version")?;
    let source_revision = field(input, "revision")?;
    source_check(root, db, selected, source_revision)?;
    live_binding(root, db, source_revision)?;
    let source = canonical_directory(root)?;
    let target = canonical_directory(Path::new(field(input, "rewriteTargetRoot")?))?;
    let owner = owner_revision(&source, true)?;
    if source == target
        || source.parent() != target.parent()
        || owner_revision(&target, true)? != owner
        || fs::read_dir(&target).map_err(err)?.next().is_some()
    {
        return Err(crate::coded(
            "layout-unsupported",
            "Private extent target must be an empty sibling under the same owner",
        ));
    }
    let id = new_id(db)?;
    let seal = Seal {
        schema: workbook_format!(SCHEMA_USED_RANGE_REWRITE_SEAL_V1).into(),
        input_id: id.clone(),
        nonce: new_id(db)?,
        source_root: source.to_string_lossy().into_owned(),
        target_root: target.to_string_lossy().into_owned(),
        owner_revision: owner.clone(),
    };
    let bytes = serde_json::to_vec(&seal).map_err(err)?;
    if bytes.len() > 4096 {
        return Err(crate::coded(
            "layout-limit",
            "Private extent candidate seal exceeds its bound",
        ));
    }
    let mut file = crate::quota::create_new(local(&target, SEAL_FILE)?).map_err(err)?;
    file.write_all(&bytes).map_err(err)?;
    file.sync_all().map_err(err)?;
    #[cfg(unix)]
    {
        File::open(&target).map_err(err)?.sync_all().map_err(err)?;
    }
    let record = InputRecord {
        id,
        source_root: seal.source_root,
        target_root: seal.target_root,
        owner_revision: owner,
        target_identity: directory_identity(&target)?,
        seal_revision: sha(&bytes),
        version: selected,
        source_revision: source_revision.into(),
        index_revision: hash(&local(root, "index.json")?)?,
        metadata_revision: hash(&local(root, "metadata.xlsx")?)?,
        input_file: "export.xlsx".into(),
        input_revision: String::new(),
        sheets: Vec::new(),
        extents: Vec::new(),
    };
    Ok(Some(ExportCandidate {
        record,
        seal_active: true,
    }))
}
pub(crate) fn finish_export(
    root: &Path,
    db: &Connection,
    mut candidate: ExportCandidate,
    output_file: &str,
    output_revision: &str,
) -> Result<String> {
    if output_file != "export.xlsx"
        || !revision(output_revision)
        || hash(&local(root, output_file)?)? != output_revision
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Private extent export bytes differ from native output",
        ));
    }
    candidate.record.input_revision = output_revision.into();
    candidate.record.sheets = worksheet_inventory(&local(root, output_file)?)?;
    if candidate.record.sheets != worksheet_inventory(&local(root, "metadata.xlsx")?)? {
        return Err(crate::coded(
            "layout-stale",
            "Private extent export changed canonical worksheet bindings",
        ));
    }
    for sheet in &candidate.record.sheets {
        crate::cancel::check()?;
        let reply = crate::used_range::command(
            root,
            &json!({"action":"workbook","operation":"used-range","root":root,"version":candidate.record.version,"sheet":sheet.path,"sourceRevision":candidate.record.source_revision}),
        )?;
        let floor = if reply["status"] == "complete" {
            if reply["empty"] == true {
                None
            } else {
                serde_json::from_value(reply["scopeMinMax"].clone()).map_err(err)?
            }
        } else {
            serde_json::from_value(reply["provenance"]["acceptedHistory"].clone()).map_err(err)?
        };
        candidate.record.extents.push(Extent {
            sheet: sheet.path.clone(),
            floor,
            reasons: reply["reasons"]
                .as_array()
                .ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing native extent guard")
                })?
                .clone(),
        });
    }
    check_input(root, db, &candidate.record)?;
    let data = encoded(&candidate.record)?;
    if !db.is_autocommit() {
        return Err(crate::coded(
            "layout-unsupported",
            "Private extent input receipt requires a finished export transaction",
        ));
    }
    let tx = db.unchecked_transaction().map_err(err)?;
    receipt_schema(&tx)?;
    tx.execute("DELETE FROM export_rewrite_outputs", [])
        .map_err(err)?;
    tx.execute("DELETE FROM export_rewrite_inputs", [])
        .map_err(err)?;
    if table(&tx, "export_rewrite_insertion_plans")? {
        tx.execute("DELETE FROM export_rewrite_insertion_plans", [])
            .map_err(err)?;
    }
    crate::quota::reserve_artifact(data.len() as u64 + 4096)?;
    tx.execute(
        "INSERT INTO export_rewrite_inputs VALUES(1,?1,?2,?3)",
        params![candidate.record.id, candidate.record.version, data],
    )
    .map_err(err)?;
    tx.commit().map_err(err)?;
    candidate.seal_active = false;
    Ok(candidate.record.id.clone())
}
fn input_record(db: &Connection, id: &str) -> Result<InputRecord> {
    if !uuid(id) || !table(db, "export_rewrite_inputs")? {
        return Err(crate::coded(
            "layout-unsupported",
            "Private extent writer has no native input receipt",
        ));
    }
    let text: Option<String> = db
        .query_row(
            "SELECT data FROM export_rewrite_inputs WHERE id=?1",
            [id],
            |r| r.get(0),
        )
        .optional()
        .map_err(err)?;
    let text = text.ok_or_else(|| {
        crate::coded(
            "layout-stale",
            "Private extent input receipt was superseded",
        )
    })?;
    if text.len() > MAX_RECORD {
        return Err(crate::coded(
            "layout-limit",
            "Private extent input receipt exceeds its bound",
        ));
    }
    serde_json::from_str(&text).map_err(err)
}
pub(crate) fn admit_writer(
    root: &Path,
    db: &Connection,
    input: &Value,
    kind: WriterKind,
) -> Result<Option<WriterInput>> {
    let Some(id) = input.get("rewriteInputId") else {
        return Ok(None);
    };
    if input.get("diagnostics").is_some_and(|v| !v.is_boolean()) {
        return Err(crate::coded(
            "layout-unsupported",
            "Private extent writer diagnostics must be boolean",
        ));
    }
    let operation = field(input, "operation")?;
    let valid = match kind {
        WriterKind::Package => operation == "package-edit",
        WriterKind::Insertion => ["plan-references", "insert-references"].contains(&operation),
        WriterKind::Transform => operation == "transform",
        WriterKind::Layout => operation == "layout",
        WriterKind::Dedup => operation == "dedup-apply",
    };
    if input["action"] != "workbook" || !valid {
        return Err(crate::coded(
            "layout-unsupported",
            "Private extent token is outside its native writer operation",
        ));
    }
    let record = input_record(
        db,
        id.as_str().ok_or_else(|| {
            crate::coded(
                "internal-invalid-field",
                "Invalid private extent input token",
            )
        })?,
    )?;
    if version(input, "version")? != record.version || field(input, "source")? != record.input_file
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Private extent writer input differs from selected export",
        ));
    }
    check_input(root, db, &record)?;
    Ok(Some(WriterInput { record, kind }))
}

fn union(result: &mut Option<Bounds>, value: Bounds) {
    match result {
        Some(old) => {
            old.min_row = old.min_row.min(value.min_row);
            old.min_column = old.min_column.min(value.min_column);
            old.max_row = old.max_row.max(value.max_row);
            old.max_column = old.max_column.max(value.max_column);
        }
        None => *result = Some(value),
    }
}
fn axis_interval(
    first: u32,
    last: u32,
    index: u32,
    count: u32,
    remove: bool,
    limit: u32,
) -> Result<Option<(u32, u32)>> {
    if first > last || last >= limit || count == 0 || index >= limit || count > limit - index {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid native extent axis mapping",
        ));
    }
    if !remove {
        let shift = |value: u32| {
            if value < index {
                Some(value)
            } else {
                value.checked_add(count).filter(|v| *v < limit)
            }
        };
        let Some(min) = shift(first) else {
            return Ok(None);
        };
        let max = shift(last).unwrap_or(limit - 1);
        return Ok(Some((min, max)));
    }
    let end = index + count;
    let mut result = None;
    if first < index {
        result = Some((first, last.min(index - 1)));
    }
    if last >= end {
        let next = (first.max(end) - count, last - count);
        result = Some(match result {
            Some((a, b)) => (a.min(next.0), b.max(next.1)),
            None => next,
        });
    }
    Ok(result)
}
fn mapped_floor(
    floor: Option<Bounds>,
    mapping: &NativeMapping,
    sheet: &str,
) -> Result<Option<Bounds>> {
    let Some(mut b) = floor else { return Ok(None) };
    match mapping {
        NativeMapping::Axis {
            sheet: target,
            axis,
            index,
            count,
            remove,
        } if target == sheet => {
            let interval = match axis {
                Axis::Row => axis_interval(b.min_row, b.max_row, *index, *count, *remove, 1048576)?,
                Axis::Column => {
                    axis_interval(b.min_column, b.max_column, *index, *count, *remove, 16384)?
                }
            };
            let Some((a, z)) = interval else {
                return Ok(None);
            };
            match axis {
                Axis::Row => {
                    b.min_row = a;
                    b.max_row = z;
                }
                Axis::Column => {
                    b.min_column = a;
                    b.max_column = z;
                }
            }
            Ok(Some(b))
        }
        NativeMapping::Rows {
            sheet: target,
            first,
            last,
            rows,
        } if target == sheet => {
            if first > last || *last >= 1048576 {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Invalid native extent row permutation",
                ));
            }
            let expected = (*last - *first + 1) as usize;
            if rows.len() != expected {
                return Err(crate::coded(
                    "layout-unsupported",
                    "Incomplete native extent row permutation",
                ));
            }
            let (mut sources, mut destinations) = (BTreeSet::new(), BTreeSet::new());
            let mut result = None;
            if b.min_row < *first {
                union(
                    &mut result,
                    Bounds {
                        max_row: b.max_row.min(first - 1),
                        ..b
                    },
                );
            }
            if b.max_row > *last {
                union(
                    &mut result,
                    Bounds {
                        min_row: b.min_row.max(last + 1),
                        ..b
                    },
                );
            }
            for (a, z) in rows {
                crate::cancel::check()?;
                if a < first
                    || a > last
                    || z < first
                    || z > last
                    || !sources.insert(*a)
                    || !destinations.insert(*z)
                {
                    return Err(crate::coded(
                        "internal-invalid-field",
                        "Invalid native extent row permutation",
                    ));
                }
                if *a >= b.min_row && *a <= b.max_row {
                    union(
                        &mut result,
                        Bounds {
                            min_row: *z,
                            max_row: *z,
                            ..b
                        },
                    );
                }
            }
            Ok(result)
        }
        NativeMapping::Compact {
            sheet: target,
            first,
            last,
            body,
            rows,
        } if target == sheet => {
            if first.0 > last.0
                || first.1 > last.1
                || last.0 >= 1048576
                || last.1 >= 16384
                || *body < first.0
                || *body > last.0
            {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Invalid native extent cell compaction",
                ));
            }
            let mut result = None;
            // Only selected columns' cell coordinates compact. Outside columns
            // and physical row effects retain their coordinates in target facts.
            if b.min_row < *body {
                union(
                    &mut result,
                    Bounds {
                        max_row: b.max_row.min(body - 1),
                        ..b
                    },
                );
            }
            if b.max_row > last.0 {
                union(
                    &mut result,
                    Bounds {
                        min_row: b.min_row.max(last.0 + 1),
                        ..b
                    },
                );
            }
            let r0 = b.min_row.max(*body);
            let r1 = b.max_row.min(last.0);
            if r0 <= r1 {
                if b.min_column < first.1 {
                    union(
                        &mut result,
                        Bounds {
                            min_row: r0,
                            max_row: r1,
                            max_column: b.max_column.min(first.1 - 1),
                            ..b
                        },
                    );
                }
                if b.max_column > last.1 {
                    union(
                        &mut result,
                        Bounds {
                            min_row: r0,
                            max_row: r1,
                            min_column: b.min_column.max(last.1 + 1),
                            ..b
                        },
                    );
                }
            }
            let c0 = b.min_column.max(first.1);
            let c1 = b.max_column.min(last.1);
            let mut sources = BTreeSet::new();
            let mut destinations = BTreeSet::new();
            if rows.len() != (last.0 - *body + 1) as usize {
                return Err(crate::coded(
                    "layout-unsupported",
                    "Incomplete native extent compaction map",
                ));
            }
            for (a, z) in rows {
                crate::cancel::check()?;
                if *a < *body || *a > last.0 || !sources.insert(*a) {
                    return Err(crate::coded(
                        "internal-invalid-field",
                        "Invalid native extent compaction source",
                    ));
                }
                if let Some(z) = z {
                    if *z < *body || *z > last.0 || !destinations.insert(*z) {
                        return Err(crate::coded(
                            "internal-invalid-field",
                            "Invalid native extent compaction destination",
                        ));
                    }
                    if *a >= b.min_row && *a <= b.max_row && c0 <= c1 {
                        union(
                            &mut result,
                            Bounds {
                                min_row: *z,
                                max_row: *z,
                                min_column: c0,
                                max_column: c1,
                            },
                        );
                    }
                }
            }
            Ok(result)
        }
        _ => Ok(Some(b)),
    }
}
pub(crate) fn prepare_output(
    root: &Path,
    db: &Connection,
    input: WriterInput,
    mapping: NativeMapping,
    output_file: &str,
) -> Result<PreparedOutput> {
    if output_file != input.kind.file() {
        return Err(crate::coded(
            "layout-unsupported",
            "Private extent writer output artifact is unqualified",
        ));
    }
    check_input(root, db, &input.record)?;
    let sheets = worksheet_inventory(&local(root, output_file)?)?;
    let old = input
        .record
        .sheets
        .iter()
        .map(|s| s.path.clone())
        .collect::<BTreeSet<_>>();
    let new = sheets
        .iter()
        .map(|s| s.path.clone())
        .collect::<BTreeSet<_>>();
    let predecessors = match &mapping {
        NativeMapping::Sheets { predecessors } => predecessors.clone(),
        _ => new
            .iter()
            .map(|s| (s.clone(), Some(s.clone())))
            .collect::<BTreeMap<_, _>>(),
    };
    if predecessors.keys().cloned().collect::<BTreeSet<_>>() != new
        || predecessors.values().flatten().any(|s| !old.contains(s))
        || old
            .iter()
            .any(|s| !predecessors.values().any(|v| v.as_ref() == Some(s)))
    {
        return Err(crate::coded(
            "layout-stale",
            "Private extent writer changed undeclared worksheet bindings",
        ));
    }
    if !matches!(mapping, NativeMapping::Sheets { .. }) && old != new {
        return Err(crate::coded(
            "layout-stale",
            "Private extent coordinate rewrite changed worksheet inventory",
        ));
    }
    let target = match &mapping {
        NativeMapping::Axis { sheet, .. }
        | NativeMapping::Rows { sheet, .. }
        | NativeMapping::Compact { sheet, .. } => Some(sheet),
        _ => None,
    };
    if target.is_some_and(|s| !old.contains(s)) {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown native extent mapping worksheet",
        ));
    }
    let mut extents = Vec::new();
    for sheet in &sheets {
        let predecessor = predecessors.get(&sheet.path).ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing native extent predecessor",
            )
        })?;
        let extent = if let Some(source) = predecessor {
            let source_extent = input
                .record
                .extents
                .iter()
                .find(|e| &e.sheet == source)
                .ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing native source extent")
                })?;
            Extent {
                sheet: sheet.path.clone(),
                floor: mapped_floor(source_extent.floor, &mapping, source)?,
                reasons: source_extent.reasons.clone(),
            }
        } else {
            Extent {
                sheet: sheet.path.clone(),
                floor: None,
                reasons: Vec::new(),
            }
        };
        extents.push(extent);
    }
    let record = OutputRecord {
        id: new_id(db)?,
        input: input.record,
        writer: input.kind,
        output_file: output_file.into(),
        output_revision: hash(&local(root, output_file)?)?,
        mapping_revision: sha(encoded(&mapping)?.as_bytes()),
        clear_hyperlinks_command: None,
        hyperlink_authoring_command: None,
        page_theme_command: None,
        background_command: None,
        note_authoring_command: None,
        sheets,
        predecessors,
        extents,
    };
    Ok(PreparedOutput { record })
}
/// Only the qualified native clear writer calls this after its unchanged-member proof.
pub(crate) fn mark_clear_hyperlinks_output(
    output: &mut PreparedOutput,
    command: &Value,
) -> Result<()> {
    if output.record.writer != WriterKind::Package
        || output.record.mapping_revision != sha(encoded(&NativeMapping::Identity)?.as_bytes())
        || command["type"] != "clearHyperlinks"
    {
        return Err(crate::coded(
            "layout-bad-input",
            "Clear Hyperlinks receipt requires an identity native package writer",
        ));
    }
    output.record.clear_hyperlinks_command = Some(command.clone());
    Ok(())
}
/// Revalidate the native-issued clear receipt, the sealed selected export and
/// the complete package transformation before granting calculation authority.
pub(crate) fn validate_clear_hyperlinks_transfer(
    previous: &Path,
    root: &Path,
    selected: u64,
    current: u64,
    target_revision: &str,
    id: &str,
) -> Result<()> {
    validate_link_transfer(previous, root, selected, current, target_revision, id, 0)
}
pub(crate) fn validate_hyperlink_transfer(
    previous: &Path,
    root: &Path,
    selected: u64,
    current: u64,
    target_revision: &str,
    id: &str,
) -> Result<()> {
    validate_link_transfer(previous, root, selected, current, target_revision, id, 1)
}
pub(crate) fn mark_hyperlink_output(output: &mut PreparedOutput, command: &Value) -> Result<()> {
    if output.record.writer != WriterKind::Package
        || output.record.mapping_revision != sha(encoded(&NativeMapping::Identity)?.as_bytes())
        || !["setHyperlink", "removeHyperlink"]
            .iter()
            .any(|kind| command["type"] == *kind)
    {
        return Err(crate::coded(
            "layout-bad-input",
            "Hyperlink authoring receipt requires its native identity writer",
        ));
    }
    output.record.hyperlink_authoring_command = Some(command.clone());
    Ok(())
}
pub(crate) fn mark_page_theme_output(output: &mut PreparedOutput, command: &Value) -> Result<()> {
    if output.record.writer != WriterKind::Package || command["type"] != "applyPageTheme" {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid page theme writer mark",
        ));
    }
    output.record.page_theme_command = Some(command.clone());
    Ok(())
}
pub(crate) fn validate_page_theme_transfer(
    previous: &Path,
    root: &Path,
    selected: u64,
    current: u64,
    target_revision: &str,
    id: &str,
) -> Result<()> {
    validate_link_transfer(previous, root, selected, current, target_revision, id, 2)
}
pub(crate) fn mark_background_output(output: &mut PreparedOutput, command: &Value) -> Result<()> {
    if output.record.writer != WriterKind::Package
        || !["setSheetBackground", "removeSheetBackground"]
            .iter()
            .any(|kind| command["type"] == *kind)
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid background writer mark",
        ));
    }
    output.record.background_command = Some(command.clone());
    Ok(())
}
pub(crate) fn validate_background_transfer(
    previous: &Path,
    root: &Path,
    selected: u64,
    current: u64,
    target_revision: &str,
    id: &str,
) -> Result<()> {
    validate_link_transfer(previous, root, selected, current, target_revision, id, 3)
}
pub(crate) fn mark_note_output(output: &mut PreparedOutput, command: &Value) -> Result<()> {
    if output.record.writer != WriterKind::Package
        || output.record.mapping_revision != sha(encoded(&NativeMapping::Identity)?.as_bytes())
        || !["setNote", "removeNote"]
            .iter()
            .any(|k| command["type"] == *k)
    {
        return Err(crate::coded(
            "layout-bad-input",
            "Note authoring requires its native identity writer",
        ));
    }
    output.record.note_authoring_command = Some(command.clone());
    Ok(())
}
pub(crate) fn validate_note_transfer(
    previous: &Path,
    root: &Path,
    selected: u64,
    current: u64,
    target_revision: &str,
    id: &str,
) -> Result<()> {
    validate_link_transfer(previous, root, selected, current, target_revision, id, 4)
}
fn validate_link_transfer(
    previous: &Path,
    root: &Path,
    selected: u64,
    current: u64,
    target_revision: &str,
    id: &str,
    purpose: u8,
) -> Result<()> {
    let previous = canonical_directory(previous)?;
    let root = canonical_directory(root)?;
    if previous == root || previous.parent() != root.parent() || !uuid(id) {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid Clear Hyperlinks calculation candidate",
        ));
    }
    let old = read_db(&previous)?;
    if !table(&old, "export_rewrite_outputs")? {
        return Err(crate::coded(
            "layout-bad-input",
            "Clear Hyperlinks has no native writer receipt",
        ));
    }
    let text: Option<String> = old
        .prepare_cached("SELECT data FROM export_rewrite_outputs WHERE id=?1")
        .map_err(crate::err)?
        .query_row([id], |r| r.get(0))
        .optional()
        .map_err(err)?;
    let text = text.ok_or_else(|| {
        crate::coded(
            "layout-stale",
            "Clear Hyperlinks native writer receipt was superseded",
        )
    })?;
    if text.len() > MAX_RECORD {
        return Err(crate::coded(
            "layout-limit",
            "Clear Hyperlinks writer receipt exceeds its bound",
        ));
    }
    let output: OutputRecord = serde_json::from_str(&text).map_err(err)?;
    let command = match purpose {
        4 => output.note_authoring_command.as_ref(),
        1 => output.hyperlink_authoring_command.as_ref(),
        2 => output.page_theme_command.as_ref(),
        3 => output.background_command.as_ref(),
        _ => output.clear_hyperlinks_command.as_ref(),
    }
    .ok_or_else(|| {
        crate::coded(
            "layout-bad-input",
            "Writer receipt is not the requested native hyperlink operation",
        )
    })?;
    if output.id != id
        || output.writer != WriterKind::Package
        || output.output_file != "package.xlsx"
        || output.input.version != selected
        || output.input.target_root != root.to_string_lossy()
        || output.output_revision != target_revision
        || output.mapping_revision != sha(encoded(&NativeMapping::Identity)?.as_bytes())
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Clear Hyperlinks transfer differs from native writer receipt",
        ));
    }
    if encoded(&input_record(&old, &output.input.id)?)? != encoded(&output.input)? {
        return Err(crate::coded(
            "layout-stale",
            "Clear Hyperlinks selected input was superseded",
        ));
    }
    check_input(&previous, &old, &output.input)?;
    if hash(&local(&previous, &output.output_file)?)? != target_revision
        || worksheet_inventory(&local(&previous, &output.output_file)?)? != output.sheets
    {
        return Err(crate::coded(
            "layout-stale",
            "Clear Hyperlinks writer output changed",
        ));
    }
    let db = crate::database(&root)?;
    source_check(&root, &db, current, target_revision)?;
    live_binding(&root, &db, target_revision)?;
    if worksheet_inventory(&local(&root, "source.xlsx")?)? != output.sheets
        || worksheet_inventory(&local(&root, "metadata.xlsx")?)? != output.sheets
    {
        return Err(crate::coded(
            "layout-stale",
            "Clear Hyperlinks candidate bindings changed",
        ));
    }
    if purpose == 4 {
        crate::package_edit::verify_note_rewrite(&previous, command, &output.output_file)?;
    } else if purpose == 3 {
        crate::package_edit::verify_background_rewrite(&previous, command, &output.output_file)?;
    } else if purpose == 2 {
        crate::package_edit::verify_page_theme_rewrite(&previous, command, &output.output_file)?;
    } else if purpose == 1 {
        crate::package_edit::verify_hyperlink_rewrite(&previous, command, &output.output_file)?;
    } else {
        crate::package_edit::verify_clear_hyperlinks_rewrite(
            &previous,
            command,
            &output.output_file,
        )?;
    }
    Ok(())
}
pub(crate) fn commit_output(root: &Path, output: PreparedOutput) -> Result<String> {
    let db = crate::database(root)?;
    let record = output.record;
    let current = input_record(&db, &record.input.id)?;
    if encoded(&current)? != encoded(&record.input)? {
        return Err(crate::coded(
            "layout-stale",
            "Private extent writer input receipt changed",
        ));
    }
    check_input(root, &db, &record.input)?;
    if hash(&local(root, &record.output_file)?)? != record.output_revision
        || worksheet_inventory(&local(root, &record.output_file)?)? != record.sheets
    {
        return Err(crate::coded(
            "layout-stale",
            "Private extent writer output changed before receipt commit",
        ));
    }
    let data = encoded(&record)?;
    let tx = db.unchecked_transaction().map_err(err)?;
    receipt_schema(&tx)?;
    crate::quota::reserve_artifact(data.len() as u64 + 4096)?;
    tx.execute(
        "DELETE FROM export_rewrite_outputs WHERE output_file=?1",
        [&record.output_file],
    )
    .map_err(err)?;
    tx.execute(
        "INSERT INTO export_rewrite_outputs VALUES(?1,?2,?3,?4,?5)",
        params![
            record.id,
            record.input.id,
            record.input.version,
            record.output_file,
            data
        ],
    )
    .map_err(err)?;
    tx.commit().map_err(err)?;
    Ok(record.id)
}

pub(crate) fn command(root: &Path, input: &Value) -> Result<Value> {
    closed(
        input,
        &[
            "action",
            "operation",
            "root",
            "version",
            "sourceRevision",
            "previousRoot",
            "previousVersion",
            "previousSourceRevision",
            "rewriteId",
        ],
    )?;
    if input["action"] != "workbook" || input["operation"] != "used-range-transfer" {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid private extent transfer operation",
        ));
    }
    let root = canonical_directory(root)?;
    let previous = canonical_directory(Path::new(field(input, "previousRoot")?))?;
    if previous == root || previous.parent() != root.parent() {
        return Err(crate::coded(
            "layout-unsupported",
            "Private extent transfer requires distinct sibling indices",
        ));
    }
    let old = read_db(&previous)?;
    let selected = version(input, "previousVersion")?;
    let current = version(input, "version")?;
    let id = field(input, "rewriteId")?;
    if !uuid(id) || !table(&old, "export_rewrite_outputs")? {
        return Err(crate::coded(
            "layout-unsupported",
            "Private extent transfer has no native writer receipt",
        ));
    }
    let text: Option<String> = old
        .query_row(
            "SELECT data FROM export_rewrite_outputs WHERE id=?1",
            [id],
            |r| r.get(0),
        )
        .optional()
        .map_err(err)?;
    let text = text.ok_or_else(|| {
        crate::coded(
            "layout-stale",
            "Private extent writer receipt was superseded",
        )
    })?;
    if text.len() > MAX_RECORD {
        return Err(crate::coded(
            "layout-limit",
            "Private extent writer receipt exceeds its bound",
        ));
    }
    let output: OutputRecord = serde_json::from_str(&text).map_err(err)?;
    if output.id != id
        || output.input.version != selected
        || output.input.source_revision != field(input, "previousSourceRevision")?
        || output.input.target_root != root.to_string_lossy()
        || output.output_file != output.writer.file()
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Private extent transfer differs from its native writer receipt",
        ));
    }
    let active = input_record(&old, &output.input.id)?;
    if encoded(&active)? != encoded(&output.input)? {
        return Err(crate::coded(
            "layout-stale",
            "Private extent transfer input was superseded",
        ));
    }
    check_input(&previous, &old, &output.input)?;
    if hash(&local(&previous, &output.output_file)?)? != output.output_revision
        || worksheet_inventory(&local(&previous, &output.output_file)?)? != output.sheets
    {
        return Err(crate::coded(
            "layout-stale",
            "Private extent transfer writer output changed",
        ));
    }
    let db = crate::database(&root)?;
    let target_revision = field(input, "sourceRevision")?;
    let target_index = source_check(&root, &db, current, target_revision)?;
    live_binding(&root, &db, target_revision)?;
    if target_revision != output.output_revision
        || worksheet_inventory(&local(&root, "metadata.xlsx")?)? != output.sheets
        || worksheet_inventory(&local(&root, "source.xlsx")?)? != output.sheets
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Private extent transfer target differs from native writer output",
        ));
    }
    let paths = target_index["sheets"]
        .as_array()
        .ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing private extent target worksheets",
            )
        })?
        .iter()
        .map(|s| field(s, "path").map(str::to_owned))
        .collect::<Result<BTreeSet<_>>>()?;
    if paths != output.sheets.iter().map(|s| s.path.clone()).collect() {
        return Err(crate::coded(
            "layout-unsupported",
            "Private extent imported target worksheet inventory is incomplete",
        ));
    }
    let tx = db.unchecked_transaction().map_err(err)?;
    known_version(&tx, current)?;
    tx.execute_batch("CREATE TABLE IF NOT EXISTS export_rewrite_acceptance(singleton INTEGER PRIMARY KEY CHECK(singleton=1),rewrite_id TEXT NOT NULL,target_version INTEGER NOT NULL,accepted_version INTEGER NOT NULL,source_revision TEXT NOT NULL);").map_err(err)?;
    if tx
        .query_row(
            "SELECT EXISTS(SELECT 1 FROM export_rewrite_acceptance)",
            [],
            |r| r.get::<_, bool>(0),
        )
        .map_err(err)?
    {
        return Err(crate::coded(
            "layout-unsupported",
            "Private extent candidate has already accepted its rewrite receipt",
        ));
    }
    tx.execute("INSERT INTO versions(parent)VALUES(?1)", [current])
        .map_err(err)?;
    let accepted = tx.last_insert_rowid() as u64;
    let mut incomplete = 0usize;
    crate::used_range::guard_schema(&tx)?;
    for extent in &output.extents {
        crate::cancel::check()?;
        let mut floor = extent.floor;
        if let Some(target_history) = crate::used_range::history(&tx, current, &extent.sheet)? {
            union(&mut floor, target_history);
        }
        if let Some(b) = floor {
            tx.execute(
                "INSERT OR REPLACE INTO used_range_history VALUES(?1,?2,?3,?4,?5,?6)",
                params![
                    accepted,
                    extent.sheet,
                    b.min_row,
                    b.min_column,
                    b.max_row,
                    b.max_column
                ],
            )
            .map_err(err)?;
        }
        let mut reasons = extent.reasons.clone();
        reasons.extend(crate::used_range::guard_reasons(
            &tx,
            current,
            &extent.sheet,
        )?);
        if !reasons.is_empty() {
            incomplete += 1;
            crate::used_range::store_guards(&tx, accepted, &extent.sheet, &reasons)?;
        }
    }
    tx.execute(
        "INSERT INTO export_rewrite_acceptance VALUES(1,?1,?2,?3,?4)",
        params![id, current, accepted, target_revision],
    )
    .map_err(err)?;
    check_seal(&output.input)?;
    if hash(&local(&root, "source.xlsx")?)? != target_revision
        || hash(&local(&previous, &output.output_file)?)? != target_revision
    {
        return Err(crate::coded(
            "layout-stale",
            "Private extent transfer source changed before acceptance",
        ));
    }
    tx.commit().map_err(err)?;
    Ok(
        json!({"schema":workbook_format!(SCHEMA_USED_RANGE_TRANSFER_V1),"version":accepted,"sourceRevision":target_revision,"previousVersion":selected,"previousSourceRevision":output.input.source_revision,"rewriteId":id,"transferredSheets":output.extents.len(),"incompleteSheets":incomplete}),
    )
}

/// A cache snapshot represents the physical source, never an accepted live
/// owner. Only a foreign immutable binding or a native snapshot marker permits
/// this reset; it cannot erase the current owner's live ancestry in place.
pub(crate) fn source_snapshot(root: &Path, input: &Value) -> Result<Value> {
    closed(
        input,
        &["action", "operation", "root", "version", "sourceRevision"],
    )?;
    if input["action"] != "workbook" || input["operation"] != "used-range-source-snapshot" {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid private extent snapshot operation",
        ));
    }
    let root = canonical_directory(root)?;
    let db = crate::database(&root)?;
    let selected = version(input, "version")?;
    let source_revision = field(input, "sourceRevision")?;
    source_check(&root, &db, selected, source_revision)?;
    let owner = owner_identity(&root, false)?;
    let prior = if table(&db, "used_range_index_binding")? {
        Some(binding(&db)?)
    } else {
        None
    };
    if let Some(prior) = prior.as_ref() {
        if prior.2 != source_revision
            || prior.0 == root.to_string_lossy()
            || (!owner.is_empty() && prior.1 == owner)
        {
            return Err(crate::coded(
                "layout-unsupported",
                "Live accepted extent cannot be reset as a source snapshot",
            ));
        }
        if prior.3 != "live" && prior.3 != "source-snapshot" {
            return Err(crate::coded(
                "internal-unknown-reference",
                "Unknown private extent index binding state",
            ));
        }
        if !owner.is_empty() && prior.3 != "source-snapshot" && prior.1.is_empty() {
            return Err(crate::coded(
                "layout-unsupported",
                "Owned restored extent requires a source snapshot or foreign prior owner",
            ));
        }
    } else if !owner.is_empty() {
        return Err(crate::coded(
            "layout-refused",
            "Owned extent snapshot lacks an immutable prior binding; reimport the workbook",
        ));
    }
    let tx = db.unchecked_transaction().map_err(err)?;
    if table(&tx, "used_range_history")? {
        tx.execute("DELETE FROM used_range_history", [])
            .map_err(err)?;
    }
    if table(&tx, "used_range_guards")? {
        tx.execute("DELETE FROM used_range_guards", [])
            .map_err(err)?;
    }
    for name in [
        "export_rewrite_outputs",
        "export_rewrite_inputs",
        "export_rewrite_insertion_plans",
        "export_rewrite_acceptance",
    ] {
        if table(&tx, name)? {
            tx.execute(&format!("DELETE FROM {name}"), [])
                .map_err(err)?;
        }
    }
    binding_schema(&tx)?;
    tx.execute(
        "INSERT OR REPLACE INTO used_range_index_binding VALUES(1,?1,?2,?3,?4)",
        params![
            root.to_string_lossy(),
            owner,
            source_revision,
            if owner.is_empty() {
                "source-snapshot"
            } else {
                "live"
            }
        ],
    )
    .map_err(err)?;
    if hash(&local(&root, "source.xlsx")?)? != source_revision {
        return Err(crate::coded(
            "layout-stale",
            "Private extent snapshot source changed before reset",
        ));
    }
    tx.commit().map_err(err)?;
    Ok(
        json!({"schema":workbook_format!(SCHEMA_USED_RANGE_SOURCE_SNAPSHOT_V1),"version":selected,"sourceRevision":source_revision}),
    )
}
pub(crate) fn prune(db: &Connection) -> Result<()> {
    for name in ["export_rewrite_outputs", "export_rewrite_inputs"] {
        if table(db, name)? {
            db.execute(
                &format!("DELETE FROM {name} WHERE version NOT IN(SELECT id FROM versions)"),
                [],
            )
            .map_err(err)?;
        }
    }
    if table(db, "export_rewrite_insertion_plans")? {
        if table(db, "export_rewrite_inputs")? {
            db.execute("DELETE FROM export_rewrite_insertion_plans WHERE input_id NOT IN(SELECT id FROM export_rewrite_inputs)",[]).map_err(err)?;
        } else {
            db.execute("DELETE FROM export_rewrite_insertion_plans", [])
                .map_err(err)?;
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    const SHEET: &str = "xl/worksheets/main.xml";
    const BASIC: &str = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:Z80"/><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>RAND()</f><v>0.125</v></c></row><row r="80"><c r="Z80"/></row></sheetData></worksheet>"#;
    fn new_owner() -> tempfile::TempDir {
        let parent = tempfile::tempdir().unwrap();
        fs::write(parent.path().join("owner.json"),json!({"schema":workbook_format!(WORKBOOK_INDEX_OWNER_V1),"pid":std::process::id(),"sessionId":"native-extent-fixture","maxSessionBytes":4u64*1024*1024*1024}).to_string()).unwrap();
        parent
    }
    fn source(parent: &Path, name: &str, xml: &str) -> PathBuf {
        let root = parent.join(name);
        fs::create_dir(&root).unwrap();
        let mut zip = zip::ZipWriter::new(File::create(root.join("source.xlsx")).unwrap());
        let options = zip::write::SimpleFileOptions::default();
        for (part, text) in [
            (
                "[Content_Types].xml",
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/main.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#,
            ),
            (
                "xl/workbook.xml",
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:q="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Main" sheetId="1" q:id="sheet1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="sheet1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/main.xml"></Relationship></Relationships>"#,
            ),
            (SHEET, xml),
        ] {
            zip.start_file(part, options).unwrap();
            zip.write_all(text.as_bytes()).unwrap();
        }
        zip.finish().unwrap();
        crate::import::index(&root).unwrap();
        root
    }
    fn view() -> Value {
        json!({"sheets":[{"name":"Main","path":SHEET,"storedFormulaCount":1}],"namedRanges":[],"tables":[]})
    }
    fn export_input(root: &Path, target: &Path, selected: u64, internal: bool) -> Value {
        let plan =
            crate::export_state::prepare(root, &json!({"version":selected,"view":view()})).unwrap();
        assert_eq!(plan["styleCount"], 0);
        let mut zip = zip::ZipArchive::new(File::open(root.join("source.xlsx")).unwrap()).unwrap();
        let mut workbook = String::new();
        zip.by_name("xl/workbook.xml")
            .unwrap()
            .read_to_string(&mut workbook)
            .unwrap();
        let refresh = if plan["incomplete"] == true { "1" } else { "0" };
        let replacement=workbook.replace("</workbook>",&format!("<calcPr calcMode=\"auto\" calcOnSave=\"1\" fullCalcOnLoad=\"{refresh}\" forceFullCalc=\"{refresh}\"/></workbook>"));
        fs::write(
            root.join("export-metadata.json"),
            json!({"styleIds":[],"replacements":{"xl/workbook.xml":replacement}}).to_string(),
        )
        .unwrap();
        let mut request = json!({"action":"workbook","operation":"export","root":root,"revision":index(root).unwrap()["revision"],"version":selected,"calculationMetadata":true});
        if internal {
            request["internalRewrite"] = json!(true);
            request["rewriteTargetRoot"] = json!(target);
        }
        crate::export::write(root, &request).unwrap()
    }
    fn resize(root: &Path, input_id: &Value, row: u32) -> Value {
        crate::layout::apply(root,&json!({"action":"workbook","operation":"layout","root":root,"source":"export.xlsx","version":0,"rewriteInputId":input_id,"sheet":SHEET,"row":row,"height":22,"command":{"operation":"resizeRow","sheet":"Main","index":row,"height":22}})).unwrap()
    }
    fn index_output(source: &Path, target: &Path, file: &str) {
        fs::copy(source.join(file), target.join("source.xlsx")).unwrap();
        crate::import::index(target).unwrap();
    }
    fn request(source: &Path, target: &Path, writer: &Value, current: u64) -> Value {
        json!({"action":"workbook","operation":"used-range-transfer","root":target,"version":current,"sourceRevision":index(target).unwrap()["revision"],"previousRoot":source,"previousVersion":0,"previousSourceRevision":index(source).unwrap()["revision"],"rewriteId":writer["rewriteId"]})
    }
    fn query(root: &Path, selected: u64) -> Value {
        crate::used_range::command(root,&json!({"action":"workbook","operation":"used-range","root":root,"version":selected,"sourceRevision":index(root).unwrap()["revision"],"sheet":SHEET})).unwrap()
    }
    fn snapshot_request(root: &Path, selected: u64) -> Value {
        json!({"action":"workbook","operation":"used-range-source-snapshot","root":root,"version":selected,"sourceRevision":index(root).unwrap()["revision"]})
    }
    fn copy_index(source: &Path, target: &Path) {
        fs::create_dir(target).unwrap();
        for name in ["source.xlsx", "cells.sqlite", "index.json", "metadata.xlsx"] {
            fs::copy(source.join(name), target.join(name)).unwrap();
        }
    }
    fn count(db: &Connection, name: &str) -> u64 {
        db.query_row(&format!("SELECT count(*) FROM {name}"), [], |r| r.get(0))
            .unwrap()
    }

    #[test]
    fn actual_layout_receipt_preserves_live_floor_target_facts_and_manual_pending_policy() {
        let _scope = crate::shared::begin();
        let owner = new_owner();
        let old = source(owner.path(), "base", BASIC);
        let target = owner.path().join("candidate");
        fs::create_dir(&target).unwrap();
        let input = export_input(&old, &target, 0, true);
        assert!(uuid(input["rewriteInputId"].as_str().unwrap()));
        let writer = resize(&old, &input["rewriteInputId"], 0);
        assert!(uuid(writer["rewriteId"].as_str().unwrap()));
        index_output(&old, &target, "layout.xlsx");
        assert_eq!(query(&old, 0)["lastCell"], "Z80");
        assert_eq!(query(&target, 0)["lastCell"], "B1");
        crate::calculation_policy::command(&target,&json!({"method":"initialize","version":0,"view":view(),"initialPolicy":{"mode":"manual","recalculateBeforeSave":false}})).unwrap();
        let before = index(&target).unwrap();
        let old_db = hash(&old.join("cells.sqlite")).unwrap();
        let db = crate::database(&target).unwrap();
        let policies = count(&db, "calculation_policies");
        let pending = count(&db, "calculation_pending");
        let changes = count(&db, "changes");
        drop(db);
        let result = command(&target, &request(&old, &target, &writer, 0)).unwrap();
        let version = result["version"].as_u64().unwrap();
        assert_eq!(result["transferredSheets"], 1);
        assert_eq!(result["incompleteSheets"], 0);
        assert_eq!(query(&target, version)["lastCell"], "Z80");
        assert_eq!(query(&target, 0)["lastCell"], "B1");
        assert_eq!(index(&target).unwrap(), before);
        assert_eq!(hash(&old.join("cells.sqlite")).unwrap(), old_db);
        let db = crate::database(&target).unwrap();
        assert_eq!(
            crate::calculation_policy::options(&db, version).unwrap(),
            Some(("manual".into(), false))
        );
        assert!(crate::calculation_policy::pending(&db, version, SHEET, (0, 1)).unwrap());
        assert_eq!(count(&db, "calculation_policies"), policies);
        assert_eq!(count(&db, "calculation_pending"), pending);
        assert_eq!(count(&db, "changes"), changes);
        let cells = crate::range::read_db(
            &db,
            &json!({"version":version,"sheet":SHEET,"range":"B1","limit":2}),
        )
        .unwrap();
        assert_eq!(cells["cells"][0][1]["cachedValue"], 0.125);
        assert_eq!(
            db.prepare_cached("SELECT parent FROM versions WHERE id=?1")
                .unwrap()
                .query_row([version], |r| { r.get::<_, u64>(0) })
                .unwrap(),
            0
        );
    }
    #[test]
    fn transfer_refuses_other_equal_byte_candidate_and_foreign_owner_without_mutation() {
        let _scope = crate::shared::begin();
        let owner = new_owner();
        let old = source(owner.path(), "base", BASIC);
        let target = owner.path().join("candidate");
        fs::create_dir(&target).unwrap();
        let input = export_input(&old, &target, 0, true);
        let writer = resize(&old, &input["rewriteInputId"], 0);
        index_output(&old, &target, "layout.xlsx");
        let peer = owner.path().join("peer");
        fs::create_dir(&peer).unwrap();
        index_output(&old, &peer, "layout.xlsx");
        let before = hash(&peer.join("cells.sqlite")).unwrap();
        assert!(
            command(&peer, &request(&old, &peer, &writer, 0))
                .unwrap_err()
                .contains("native writer receipt")
        );
        assert_eq!(hash(&peer.join("cells.sqlite")).unwrap(), before);
        let foreign = new_owner();
        let outside = foreign.path().join("candidate");
        fs::create_dir(&outside).unwrap();
        index_output(&old, &outside, "layout.xlsx");
        let before = hash(&outside.join("cells.sqlite")).unwrap();
        assert!(command(&outside, &request(&old, &outside, &writer, 0)).is_err());
        assert_eq!(hash(&outside.join("cells.sqlite")).unwrap(), before);
    }
    #[test]
    fn strict_transfer_fields_stale_versions_and_changed_seal_refuse_atomically() {
        let _scope = crate::shared::begin();
        let owner = new_owner();
        let old = source(owner.path(), "base", BASIC);
        let target = owner.path().join("candidate");
        fs::create_dir(&target).unwrap();
        let input = export_input(&old, &target, 0, true);
        let writer = resize(&old, &input["rewriteInputId"], 0);
        index_output(&old, &target, "layout.xlsx");
        let base = request(&old, &target, &writer, 0);
        let before = hash(&target.join("cells.sqlite")).unwrap();
        for (key, value) in [
            ("previousVersion", json!(1)),
            ("version", json!(99)),
            ("previousSourceRevision", json!("a".repeat(64))),
            ("sourceRevision", json!("b".repeat(64))),
            ("diagnostics", Value::Null),
            ("mapping", json!({"lastCell":"XFD1048576"})),
        ] {
            let mut bad = base.clone();
            bad[key] = value;
            assert!(command(&target, &bad).is_err(), "{key}");
            assert_eq!(hash(&target.join("cells.sqlite")).unwrap(), before);
        }
        let seal = fs::read(target.join(SEAL_FILE)).unwrap();
        fs::write(target.join(SEAL_FILE), b"{}").unwrap();
        assert!(command(&target, &base).unwrap_err().contains("seal"));
        assert_eq!(hash(&target.join("cells.sqlite")).unwrap(), before);
        fs::write(target.join(SEAL_FILE), seal).unwrap();
        let mut traced = base;
        traced["diagnostics"] = json!(true);
        assert!(command(&target, &traced).is_ok());
    }
    #[test]
    fn changed_writer_output_and_source_bytes_never_grant_transfer() {
        let _scope = crate::shared::begin();
        let owner = new_owner();
        let old = source(owner.path(), "base", BASIC);
        let target = owner.path().join("candidate");
        fs::create_dir(&target).unwrap();
        let input = export_input(&old, &target, 0, true);
        let writer = resize(&old, &input["rewriteInputId"], 0);
        index_output(&old, &target, "layout.xlsx");
        let base = request(&old, &target, &writer, 0);
        let before = hash(&target.join("cells.sqlite")).unwrap();
        for file in ["layout.xlsx", "export.xlsx", "source.xlsx"] {
            let bytes = fs::read(old.join(file)).unwrap();
            let mut bad = bytes.clone();
            bad.push(0);
            fs::write(old.join(file), bad).unwrap();
            assert!(command(&target, &base).is_err(), "{file}");
            assert_eq!(hash(&target.join("cells.sqlite")).unwrap(), before);
            fs::write(old.join(file), bytes).unwrap();
        }
    }
    #[test]
    fn ordinary_copy_has_no_input_receipt_and_cancelled_candidates_are_bounded() {
        let _scope = crate::shared::begin();
        let owner = new_owner();
        let old = source(owner.path(), "base", BASIC);
        let normal = export_input(&old, owner.path(), 0, false);
        assert!(normal.get("rewriteInputId").is_none());
        let db = crate::database(&old).unwrap();
        assert!(!table(&db, "export_rewrite_inputs").unwrap());
        drop(db);
        let mut previous: Option<String> = None;
        for n in 0..12 {
            let target = owner.path().join(format!("candidate-{n}"));
            fs::create_dir(&target).unwrap();
            let input = export_input(&old, &target, 0, true);
            if let Some(previous) = previous.take() {
                let db = crate::database(&old).unwrap();
                assert!(input_record(&db, &previous).is_err());
            }
            let writer = resize(&old, &input["rewriteInputId"], 0);
            assert!(uuid(writer["rewriteId"].as_str().unwrap()));
            let db = crate::database(&old).unwrap();
            assert_eq!(count(&db, "export_rewrite_inputs"), 1);
            assert_eq!(count(&db, "export_rewrite_outputs"), 1);
            previous = Some(input["rewriteInputId"].as_str().unwrap().to_owned());
        }
        assert_eq!(query(&old, 0)["lastCell"], "Z80");
    }
    #[test]
    fn repaired_output_cannot_launder_an_incomplete_source_guard_and_compaction_keeps_it() {
        let _scope = crate::shared::begin();
        let owner = new_owner();
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A100"/><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="100" ht="0" customHeight="1"/></sheetData></worksheet>"#;
        let old = source(owner.path(), "base", xml);
        assert_eq!(query(&old, 0)["status"], "incomplete");
        let target = owner.path().join("candidate");
        fs::create_dir(&target).unwrap();
        let input = export_input(&old, &target, 0, true);
        let writer = resize(&old, &input["rewriteInputId"], 99);
        index_output(&old, &target, "layout.xlsx");
        assert_eq!(query(&target, 0)["status"], "complete");
        let result = command(&target, &request(&old, &target, &writer, 0)).unwrap();
        let accepted = result["version"].as_u64().unwrap();
        assert_eq!(result["incompleteSheets"], 1);
        assert_eq!(query(&target, accepted)["status"], "incomplete");
        let child=crate::range::patch(&target,&json!({"version":accepted,"changes":[{"sheet":SHEET,"cell":"A1","value":{"value":2,"dataType":"number","rawValue":"2","cachedValue":null},"edit":{"value":2}}]})).unwrap()["version"].as_u64().unwrap();
        crate::compact::compact(&target, &json!({"keep":[child],"vacuum":false})).unwrap();
        assert_eq!(query(&target, child)["status"], "incomplete");
        assert_eq!(query(&target, 0)["status"], "complete");
    }
    #[test]
    fn source_snapshot_refuses_live_root_changed_marker_and_same_owner_peer() {
        let _scope = crate::shared::begin();
        let owner = new_owner();
        let old = source(owner.path(), "base", BASIC);
        let db = crate::database(&old).unwrap();
        db.execute(
            "INSERT INTO used_range_history VALUES(0,?1,0,0,99,30)",
            [SHEET],
        )
        .unwrap();
        drop(db);
        let before = hash(&old.join("cells.sqlite")).unwrap();
        assert!(
            source_snapshot(&old, &snapshot_request(&old, 0))
                .unwrap_err()
                .contains("Live accepted")
        );
        assert_eq!(hash(&old.join("cells.sqlite")).unwrap(), before);
        let marker = owner.path().join("owner.json");
        let bytes = fs::read(&marker).unwrap();
        let mut value: Value = serde_json::from_slice(&bytes).unwrap();
        value["sessionId"] = json!("changed-owner");
        fs::write(&marker, value.to_string()).unwrap();
        assert!(
            source_snapshot(&old, &snapshot_request(&old, 0))
                .unwrap_err()
                .contains("Live accepted")
        );
        assert_eq!(hash(&old.join("cells.sqlite")).unwrap(), before);
        fs::write(&marker, bytes).unwrap();
        let peer = owner.path().join("peer");
        copy_index(&old, &peer);
        let before = hash(&peer.join("cells.sqlite")).unwrap();
        assert!(
            source_snapshot(&peer, &snapshot_request(&peer, 0))
                .unwrap_err()
                .contains("Live accepted")
        );
        assert_eq!(hash(&peer.join("cells.sqlite")).unwrap(), before);
    }
    #[test]
    fn anonymous_snapshot_and_foreign_owner_restore_reset_extent_only() {
        let _scope = crate::shared::begin();
        let owner = new_owner();
        let old = source(owner.path(), "base", BASIC);
        crate::calculation_policy::command(&old,&json!({"method":"initialize","version":0,"view":view(),"initialPolicy":{"mode":"manual","recalculateBeforeSave":false}})).unwrap();
        let db = crate::database(&old).unwrap();
        db.execute(
            "INSERT INTO used_range_history VALUES(0,?1,0,0,99,30)",
            [SHEET],
        )
        .unwrap();
        crate::used_range::store_guards(
            &db,
            0,
            SHEET,
            &[json!({"code":"invalid-source","message":"Fixture predecessor is incomplete."})],
        )
        .unwrap();
        let pending = count(&db, "calculation_pending");
        let axes = count(&db, "used_range_axes");
        let cells = count(&db, "cells");
        drop(db);
        let anonymous = tempfile::tempdir().unwrap();
        let clone = anonymous.path().join("clone");
        copy_index(&old, &clone);
        let index_before = index(&clone).unwrap();
        source_snapshot(&clone, &snapshot_request(&clone, 0)).unwrap();
        let db = crate::database(&clone).unwrap();
        assert_eq!(count(&db, "used_range_history"), 0);
        assert_eq!(count(&db, "used_range_guards"), 0);
        assert_eq!(count(&db, "calculation_pending"), pending);
        assert_eq!(count(&db, "used_range_axes"), axes);
        assert_eq!(count(&db, "cells"), cells);
        assert_eq!(binding(&db).unwrap().3, "source-snapshot");
        drop(db);
        assert_eq!(index(&clone).unwrap(), index_before);
        let restored_owner = new_owner();
        let restored = restored_owner.path().join("restored");
        copy_index(&clone, &restored);
        source_snapshot(&restored, &snapshot_request(&restored, 0)).unwrap();
        let db = crate::database(&restored).unwrap();
        assert_eq!(binding(&db).unwrap().3, "live");
        assert_eq!(
            crate::calculation_policy::options(&db, 0).unwrap(),
            Some(("manual".into(), false))
        );
        drop(db);
        assert!(source_snapshot(&restored, &snapshot_request(&restored, 0)).is_err());
        assert_eq!(query(&restored, 0)["lastCell"], "Z80");
    }
    #[test]
    fn unbound_legacy_index_reset_is_anonymous_only_and_snapshot_fields_are_closed() {
        let _scope = crate::shared::begin();
        let owner = new_owner();
        let old = source(owner.path(), "base", BASIC);
        let db = crate::database(&old).unwrap();
        db.execute("DROP TABLE used_range_index_binding", [])
            .unwrap();
        drop(db);
        let before = hash(&old.join("cells.sqlite")).unwrap();
        assert!(source_snapshot(&old, &snapshot_request(&old, 0)).is_err());
        assert_eq!(hash(&old.join("cells.sqlite")).unwrap(), before);
        let anonymous = tempfile::tempdir().unwrap();
        let clone = anonymous.path().join("clone");
        copy_index(&old, &clone);
        let base = snapshot_request(&clone, 0);
        let before = hash(&clone.join("cells.sqlite")).unwrap();
        for (key, value) in [
            ("version", json!(999)),
            ("sourceRevision", json!("a".repeat(64))),
            ("diagnostics", json!("true")),
            ("purpose", json!("cache")),
        ] {
            let mut bad = base.clone();
            bad[key] = value;
            assert!(source_snapshot(&clone, &bad).is_err());
            assert_eq!(hash(&clone.join("cells.sqlite")).unwrap(), before);
        }
        source_snapshot(&clone, &base).unwrap();
        assert_eq!(
            binding(&crate::database(&clone).unwrap()).unwrap().3,
            "source-snapshot"
        );
    }
    #[test]
    fn native_axis_permutation_and_rectangle_compaction_math_is_bounded() {
        assert_eq!(axis_interval(2, 8, 4, 2, false, 20).unwrap(), Some((2, 10)));
        assert_eq!(axis_interval(2, 8, 4, 2, true, 20).unwrap(), Some((2, 6)));
        assert_eq!(axis_interval(4, 5, 4, 2, true, 20).unwrap(), None);
        assert!(axis_interval(0, 8, 4, 0, false, 20).is_err());
        let b = Bounds {
            min_row: 1,
            max_row: 2,
            min_column: 1,
            max_column: 2,
        };
        let rows = NativeMapping::Rows {
            sheet: SHEET.into(),
            first: 1,
            last: 3,
            rows: vec![(1, 3), (2, 1), (3, 2)],
        };
        assert_eq!(
            mapped_floor(Some(b), &rows, SHEET).unwrap(),
            Some(Bounds {
                min_row: 1,
                max_row: 3,
                ..b
            })
        );
        let bad = NativeMapping::Rows {
            sheet: SHEET.into(),
            first: 1,
            last: 3,
            rows: vec![(1, 3), (2, 3), (3, 2)],
        };
        assert!(mapped_floor(Some(b), &bad, SHEET).is_err());
        let b = Bounds {
            min_row: 1,
            max_row: 3,
            min_column: 0,
            max_column: 2,
        };
        let compact = NativeMapping::Compact {
            sheet: SHEET.into(),
            first: (0, 1),
            last: (3, 2),
            body: 1,
            rows: vec![(1, Some(1)), (2, None), (3, Some(2))],
        };
        assert_eq!(mapped_floor(Some(b), &compact, SHEET).unwrap(), Some(b));
        let inside = Bounds { min_column: 1, ..b };
        assert_eq!(
            mapped_floor(Some(inside), &compact, SHEET).unwrap(),
            Some(Bounds {
                max_row: 2,
                ..inside
            })
        );
    }
    #[test]
    fn transfer_keeps_farther_target_history_and_guards_and_candidate_is_single_use() {
        let _scope = crate::shared::begin();
        let owner = new_owner();
        let old = source(owner.path(), "base", BASIC);
        let db = crate::database(&old).unwrap();
        crate::used_range::store_guards(&db,0,SHEET,&[json!({"code":"legacy-row-facts","message":"The source predecessor has incomplete row provenance."})]).unwrap();
        drop(db);
        let target = owner.path().join("candidate");
        fs::create_dir(&target).unwrap();
        let input = export_input(&old, &target, 0, true);
        let writer = resize(&old, &input["rewriteInputId"], 0);
        index_output(&old, &target, "layout.xlsx");
        let peak=crate::range::patch(&target,&json!({"version":0,"changes":[{"sheet":SHEET,"cell":"Z5000","value":{"value":5,"dataType":"number","rawValue":"5","cachedValue":null},"edit":{"value":5}}]})).unwrap()["version"].as_u64().unwrap();
        let db = crate::database(&target).unwrap();
        crate::used_range::store_guards(&db,peak,SHEET,&[json!({"code":"invalid-cell-style","message":"The target ancestor has an independent incomplete style profile."})]).unwrap();
        drop(db);
        let accepted = command(&target, &request(&old, &target, &writer, peak)).unwrap()["version"]
            .as_u64()
            .unwrap();
        let db = crate::database(&target).unwrap();
        assert_eq!(
            crate::used_range::history(&db, accepted, SHEET)
                .unwrap()
                .unwrap()
                .max_row,
            4999
        );
        let reasons = crate::used_range::guard_reasons(&db, accepted, SHEET).unwrap();
        assert!(reasons.iter().any(|r| r["code"] == "legacy-row-facts"));
        assert!(reasons.iter().any(|r| r["code"] == "invalid-cell-style"));
        drop(db);
        let cleared=crate::range::patch(&target,&json!({"version":accepted,"changes":[{"sheet":SHEET,"cell":"Z5000","value":{"value":null,"dataType":"blank","rawValue":null,"cachedValue":null,"styleId":0},"edit":{"value":null}}]})).unwrap()["version"].as_u64().unwrap();
        let db = crate::database(&target).unwrap();
        assert_eq!(
            crate::used_range::history(&db, cleared, SHEET)
                .unwrap()
                .unwrap()
                .max_row,
            4999
        );
        drop(db);
        assert_eq!(query(&target, cleared)["status"], "incomplete");
        assert_eq!(query(&target, 0)["lastCell"], "B1");
        crate::compact::compact(&target, &json!({"keep":[cleared],"vacuum":false})).unwrap();
        let before = hash(&target.join("cells.sqlite")).unwrap();
        assert!(
            command(&target, &request(&old, &target, &writer, 0))
                .unwrap_err()
                .contains("already accepted")
        );
        assert_eq!(hash(&target.join("cells.sqlite")).unwrap(), before);
    }
    #[test]
    fn farther_target_peak_survives_transfer_clear_and_compaction_but_not_cache_snapshot() {
        let _scope = crate::shared::begin();
        let owner = new_owner();
        let old = source(owner.path(), "base", BASIC);
        let target = owner.path().join("candidate");
        fs::create_dir(&target).unwrap();
        let input = export_input(&old, &target, 0, true);
        let writer = resize(&old, &input["rewriteInputId"], 0);
        index_output(&old, &target, "layout.xlsx");
        let peak=crate::range::patch(&target,&json!({"version":0,"changes":[{"sheet":SHEET,"cell":"Z5000","value":{"value":5,"dataType":"number","rawValue":"5","cachedValue":null},"edit":{"value":5}}]})).unwrap()["version"].as_u64().unwrap();
        let accepted = command(&target, &request(&old, &target, &writer, peak)).unwrap()["version"]
            .as_u64()
            .unwrap();
        let cleared=crate::range::patch(&target,&json!({"version":accepted,"changes":[{"sheet":SHEET,"cell":"Z5000","value":{"value":null,"dataType":"blank","rawValue":null,"cachedValue":null,"styleId":0},"edit":{"value":null}}]})).unwrap()["version"].as_u64().unwrap();
        assert_eq!(query(&target, cleared)["lastCell"], "Z5000");
        assert_eq!(query(&target, 0)["lastCell"], "B1");
        crate::compact::compact(&target, &json!({"keep":[cleared],"vacuum":false})).unwrap();
        assert_eq!(query(&target, cleared)["lastCell"], "Z5000");
        let before = hash(&target.join("cells.sqlite")).unwrap();
        assert!(command(&target, &request(&old, &target, &writer, 0)).is_err());
        assert_eq!(hash(&target.join("cells.sqlite")).unwrap(), before);
        let anonymous = tempfile::tempdir().unwrap();
        let clone = anonymous.path().join("cache");
        copy_index(&target, &clone);
        source_snapshot(&clone, &snapshot_request(&clone, cleared)).unwrap();
        assert_eq!(query(&clone, cleared)["lastCell"], "B1");
        assert_eq!(
            count(
                &crate::database(&clone).unwrap(),
                "export_rewrite_acceptance"
            ),
            0
        );
        assert_eq!(hash(&target.join("cells.sqlite")).unwrap(), before);
    }
}
