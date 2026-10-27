//! File-backed workbook data owner. No Electron, UI or MCP dependencies.
#[macro_use]
mod workbook_format;
mod bulk;
mod calculation;
mod calculation_policy;
mod calculation_transfer;
mod cancel;
mod compact;
mod decimal;
mod dedup;
mod dependencies;
mod diagnostics;
mod display;
mod durability;
mod evidence;
mod export;
mod export_state;
mod format;
mod formula;
mod import;
mod insertion;
pub mod journal;
mod keys;
mod layout;
mod merges;
mod package_edit;
mod preservation;
mod preview;
mod query;
mod quota;
mod range;
mod raw_read;
mod scenarios;
mod shared;
mod spill;
mod stale_cache;
mod trace;
mod transform;
mod unhide;
mod used_range;
mod used_range_transfer;
mod version_summary;
mod workbook_view;
mod xml;

use rusqlite::Connection;
use serde_json::{Value, json};
use sha2::{Digest, Sha256};
use std::fs::{self, File};
use std::io::Read;
use std::path::{Path, PathBuf};

pub type Result<T> = std::result::Result<T, String>;
pub use durability::{flush_session, validate_session};
pub use merges::MergeCommand;
/// A refusal a caller may need to act on carries a stable code in front of its
/// sentence: `[code] sentence`. The helper binary sends the two separately, so
/// callers decide on the code and people read the sentence. Codes are listed
/// in `core/refusal-codes.mjs`.
pub(crate) fn coded(code: &str, message: impl std::fmt::Display) -> String {
    let message = message.to_string();
    // A message that already has a code (an error passed on by a layer that only knows a general code) keeps its own.
    match split_code(&message).0 {
        Some(_) => message,
        None => format!("[{code}] {message}"),
    }
}
/// The sentence of an error without its code, for code that reads what an error says.
pub(crate) fn sentence(error: &str) -> &str {
    split_code(error).1
}
/// Adds context to an error's sentence and keeps the code it already has; an error with none gets `code`.
pub(crate) fn rephrase(
    code: &str,
    error: impl std::fmt::Display,
    wrap: impl FnOnce(&str) -> String,
) -> String {
    let error = error.to_string();
    let (own, said) = split_code(&error);
    coded(own.unwrap_or(code), wrap(said))
}
/// Splits `[code] sentence` into its code and sentence.
pub fn split_code(error: &str) -> (Option<&str>, &str) {
    error
        .strip_prefix('[')
        .and_then(|rest| rest.split_once("] "))
        .filter(|(code, _)| {
            !code.is_empty() && code.bytes().all(|b| b.is_ascii_lowercase() || b == b'-')
        })
        .map_or((None, error), |(code, message)| (Some(code), message))
}
pub(crate) fn err(error: impl std::fmt::Display) -> String {
    error.to_string()
}
pub(crate) fn field<'a>(input: &'a Value, key: &str) -> Result<&'a str> {
    input[key]
        .as_str()
        .ok_or_else(|| crate::coded("internal-missing-field", format!("Missing workbook {key}")))
}
pub(crate) fn hash(path: &Path) -> Result<String> {
    let mut source = File::open(path).map_err(err)?;
    let mut digest = Sha256::new();
    let mut buffer = [0u8; 65536];
    loop {
        cancel::check()?;
        let count = source.read(&mut buffer).map_err(err)?;
        if count == 0 {
            break;
        }
        digest.update(&buffer[..count]);
    }
    Ok(format!("{:x}", digest.finalize()))
}
pub(crate) fn database(root: &Path) -> Result<Connection> {
    let connection = Connection::open(root.join("cells.sqlite")).map_err(err)?;
    connection
        .execute_batch("PRAGMA cache_size=-4096; PRAGMA temp_store=FILE; PRAGMA mmap_size=0;")
        .map_err(err)?;
    quota::database(&connection)?;
    // Per-cell loops reuse their compiled statements instead of re-parsing SQL.
    connection.set_prepared_statement_cache_capacity(256);
    connection.progress_handler(10000, Some(cancel::interrupted));
    shared::schema(&connection)?;
    Ok(connection)
}
pub(crate) fn local(root: &Path, name: &str) -> Result<PathBuf> {
    if name.is_empty() || name.contains(['/', '\\']) || name == "." || name == ".." {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid workbook artifact name",
        ));
    }
    let path = root.join(name);
    if fs::symlink_metadata(&path).is_ok_and(|m| m.file_type().is_symlink()) {
        return Err(crate::coded(
            "workbook-refused",
            "Workbook artifact is a symlink",
        ));
    }
    Ok(path)
}

fn session_root(supplied: &Path) -> Result<PathBuf> {
    if !supplied.is_absolute()
        || fs::symlink_metadata(supplied)
            .map_err(err)?
            .file_type()
            .is_symlink()
    {
        return Err(crate::coded(
            "session-unsupported",
            "Unsafe workbook session root",
        ));
    }
    fs::canonicalize(supplied).map_err(err)
}

/// Apply derived calculation patches as one native transaction and version.
/// Items use the `patch` command's `{sheet, cell, value, edits}` shape. Supply the
/// original cell with updated calculation metadata and `edits: []` to preserve
/// its source fields and inherited export intent.
/// Producer errors, cancellation and native validation/write failures roll back
/// every item. An empty iterator returns `parent` without creating a version.
pub fn apply_calculation_patches<I>(root: &Path, parent: u64, patches: I) -> Result<u64>
where
    I: IntoIterator<Item = Result<Value>>,
{
    let _shared_scope = shared::begin();
    let root = session_root(root)?;
    cancel::begin(&root);
    cancel::check()?;
    let _budget = quota::begin(&root, "patch")?;
    range::patch_changes(&root, parent, patches, false)
}

/// The application authorizes the source and owns this private session directory.
/// Requests can address fixed artifacts here, never arbitrary output paths.
pub fn command(input: Value) -> Result<Value> {
    // Journal checks need no session directory (journal.rs).
    if input["action"] == "journal" {
        return journal::command(&input);
    }
    let _shared_scope = shared::begin();
    let root = session_root(Path::new(field(&input, "root")?))?;
    let diagnostics = diagnostics::begin(input["diagnostics"] == true);
    cancel::begin(&root);
    // Preview is bounded read-only work; no quota lease/active artifact is written.
    if field(&input, "operation")? == "preview-formula" {
        return preview::command(&root, &input);
    }
    if field(&input, "operation")? == "used-range-export-needed" {
        let mut result = used_range::export_needed(&root, &input)?;
        if let Some(trace) = diagnostics.finish() {
            result["diagnostics"] = trace;
        }
        return Ok(result);
    }
    cancel::check()?;
    let _budget = quota::begin(&root, field(&input, "operation")?)?;
    let mut result = match field(&input, "operation")? {
        "reserve-candidate" | "release-candidate" => quota::candidate(&root, &input),
        "artifact-write" => quota::artifact(&root, &input),
        "copy-source" => quota::copy(&root, &input),
        "dependencies" => dependencies::command(&root, &input),
        "trace" => trace::command(&root, &input),
        "calculation-policy" => calculation_policy::command(&root, &input),
        "calculation-transfer" => calculation_transfer::command(&root, &input),
        "formula" => formula::command(&input),
        "calculate" => calculation::command(&root, &input),
        "spill-inspect" => spill::inspect(&root, &input),
        "spill-stage" => spill::stage_qualified(&root, &input),
        "financial-profiles" => formula::financial_profiles(&root, &input),
        "spill-consumers" => spill::consumers::command(&root, &input),
        "scalar-proofs" => spill::scalars::command(&root, &input),
        "spill-release" => spill::release_master(&root, &input),
        "spill-metadata-check" => spill::metadata::check(&root, &input),
        "compact" => compact::compact(&root, &input),
        "insert-references" => insertion::apply(&root, &input),
        "plan-references" => insertion::prepare(&root, &input),
        "index" => import::index(&root),
        "raw-index" => import::raw_index(&root, &input),
        "raw-query" => raw_read::query(&root, &input),
        "range" => range::read(&root, &input),
        "used-range" => used_range::command(&root, &input),
        "used-range-transfer" => used_range_transfer::command(&root, &input),
        "used-range-source-snapshot" => used_range_transfer::source_snapshot(&root, &input),
        "keys" => keys::query(&root, &input),
        "query" => query::search(&root, &input),
        "selection" => query::selection(&root, &input),
        "filter-values" => query::filter_values(&root, &input),
        "dedup-preview" => dedup::preview(&root, &input),
        "dedup-apply" => dedup::apply(&root, &input),
        "diff" => evidence::diff(&root, &input),
        "diff-bases" => evidence::bases(&root, &input),
        "summary" => evidence::summary(&root, &input),
        "version-summary" => version_summary::read(&root, &input),
        "patch" => range::patch(&root, &input),
        "merge" => merges::apply(&root, &input),
        "preservation-point-check" => preservation::check(&root, &input),
        "format" => format::apply(&root, &input),
        "bulk" => bulk::apply(&root, &input),
        "normalize-shared" => bulk::normalize_shared(&root, &input),
        "special-kinds" => bulk::special_kinds(&root),
        "transform" => transform::apply(&root, &input),
        "layout" => layout::apply(&root, &input),
        "package-edit" => package_edit::apply(&root, &input),
        "rules-apply" => package_edit::apply_rules(&root, &input),
        "notes" => package_edit::read_notes(&root, &input),
        "hyperlink" => package_edit::read_hyperlink(&root, &input),
        "boolean-cell-check" => package_edit::check_boolean_cell(&root, &input),
        "print-check" => package_edit::print_check(&root, &input),
        "export-plan" => export_state::prepare(&root, &input),
        "workbook-view" => workbook_view::command(&root, &input),
        "export-metadata" => package_edit::export_metadata::prepare(&root, &input),
        "export-styles" => export_state::styles(&root, &input),
        "export" => export::write(&root, &input),
        "stats" => {
            let db = database(&root)?;
            let cells: u64 = db
                .query_row("SELECT count(*) FROM cells", [], |row| row.get(0))
                .map_err(err)?;
            let changed_cells: u64 = db
                .query_row("SELECT count(*) FROM changes", [], |r| r.get(0))
                .map_err(err)?;
            let shared_groups: u64 = db
                .query_row("SELECT count(*) FROM shared_normalizations", [], |r| {
                    r.get(0)
                })
                .map_err(err)?;
            Ok(
                json!({"cells":cells,"indexBytes":fs::metadata(root.join("cells.sqlite")).map_err(err)?.len(),"storage":"rust-sqlite-v1","changedCellRecords":changed_cells,"normalizedSharedGroups":shared_groups}),
            )
        }
        _ => Err(crate::coded(
            "internal-unknown-reference",
            "Unknown workbook owner operation",
        )),
    }?;
    if let Some(trace) = diagnostics.finish() {
        result["diagnostics"] = trace;
    }
    Ok(result)
}
