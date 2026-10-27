//! Same-address, accepted-source copies. A batch keeps all native candidates in
//! its caller's transaction; Formats never writes content or calculation policy.
use super::{
    Bounds, STYLE_FIELDS, bounds, fill_guard, inherited_edit, merge_edit, set_value, source_cell,
    store_change, tile, values,
};
use crate::xml::address;
use crate::{Result, err, field};
use rusqlite::{Connection, params};
use serde_json::{Value, json};
use std::{
    collections::{BTreeSet, VecDeque},
    path::Path,
};

const POSITION_LIMIT: u64 = 10_000;
const FORMULA_LIMIT: usize = 10_000;
pub(super) struct Outcome {
    pub version: u64,
    pub formula_delta: i64,
    pub formula_touched: bool,
    pub content_changed: bool,
}
fn descriptor(command: &Value) -> Result<(&str, &str, Bounds, &str)> {
    let map = command.as_object().ok_or_else(|| {
        crate::coded(
            "internal-invalid-field",
            "Invalid Across Worksheets command",
        )
    })?;
    if map.len() != 5
        || map
            .keys()
            .any(|key| !["type", "sheet", "sourceSheet", "range", "mode"].contains(&key.as_str()))
        || command["type"] != "fillAcrossSheets"
    {
        return Err(crate::coded(
            "fill-bad-input",
            "Across Worksheets requires exactly type, sheet, sourceSheet, range and mode",
        ));
    }
    let sheet = field(command, "sheet")?;
    let source = field(command, "sourceSheet")?;
    if sheet == source {
        return Err(crate::coded(
            "internal-mismatch",
            "Across Worksheets source must differ from its destination",
        ));
    }
    let range = field(command, "range")?;
    let area = bounds(range)?;
    let canonical = if !range.contains(':') {
        address(area.0.0, area.0.1)
    } else {
        format!(
            "{}:{}",
            address(area.0.0, area.0.1),
            address(area.1.0, area.1.1)
        )
    };
    if range != canonical {
        return Err(crate::coded(
            "fill-bad-input",
            "Across Worksheets requires one canonical rectangle",
        ));
    }
    let mode = field(command, "mode")?;
    if !["all", "contents", "formats"].contains(&mode) {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown Across Worksheets mode",
        ));
    }
    Ok((sheet, source, area, mode))
}
pub(super) fn validate_runs(commands: &[Value]) -> Result<()> {
    let mut identity: Option<(&str, &str, &str)> = None;
    let mut destinations = BTreeSet::new();
    let mut positions = 0_u64;
    for command in commands {
        if command["type"] != "fillAcrossSheets" {
            identity = None;
            destinations.clear();
            positions = 0;
            continue;
        }
        let (sheet, source, area, mode) = descriptor(command)?;
        let current = (source, field(command, "range")?, mode);
        if identity.is_some_and(|old| old != current) {
            return Err(crate::coded(
                "fill-bad-input",
                "One Across Worksheets run requires the same source, rectangle and mode",
            ));
        }
        identity = Some(current);
        if !destinations.insert(sheet) {
            return Err(crate::coded(
                "fill-bad-input",
                "Across Worksheets destinations must be distinct",
            ));
        }
        positions = positions
            .checked_add(u64::from(area.1.0 - area.0.0 + 1) * u64::from(area.1.1 - area.0.1 + 1))
            .ok_or_else(|| {
                crate::coded("fill-limit", "Across Worksheets position count overflow")
            })?;
        if positions > POSITION_LIMIT {
            return Err(crate::coded(
                "fill-limit",
                "Across Worksheets exceeds 10,000 destination positions",
            ));
        }
    }
    Ok(())
}
fn axes(db: &Connection, index: &Value, sheet: &str, area: Bounds) -> Result<()> {
    if index["sourceFactSchema"] != crate::used_range::FACTS_SCHEMA {
        return Err(crate::coded(
            "fill-bad-input",
            "Across Worksheets requires current native axis source facts",
        ));
    }
    let inherited: bool = db.prepare_cached("SELECT EXISTS(SELECT 1 FROM used_range_axes WHERE sheet=?1 AND style<>0 AND ((axis='row' AND first<=?3 AND last>=?2) OR (axis='col' AND first<=?5 AND last>=?4)))").map_err(crate::err)?.query_row(params![sheet, area.0.0, area.1.0, area.0.1, area.1.1], |row| row.get(0)).map_err(err)?;
    if inherited {
        return Err(crate::coded(
            "fill-unsupported",
            "Across Worksheets inherited row or column styles are not qualified",
        ));
    }
    let facts = index["sheets"]
        .as_array()
        .ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing Across Worksheets source facts",
            )
        })?
        .iter()
        .find(|entry| entry["path"] == sheet)
        .ok_or_else(|| {
            crate::coded(
                "internal-unknown-reference",
                "Unknown Across Worksheets sheet",
            )
        })?;
    if facts["extended"] == true || facts["dataStart"].is_null() {
        return Err(crate::coded(
            "fill-unsupported",
            "Across Worksheets extended or incomplete worksheet source is not qualified",
        ));
    }
    if facts["usedRangeFacts"]["issues"]
        .as_array()
        .is_none_or(|issues| !issues.is_empty())
    {
        return Err(crate::coded(
            "fill-unsupported",
            "Across Worksheets incomplete worksheet source facts are not qualified",
        ));
    }
    Ok(())
}
fn ordinary(cell: &Value) -> Result<()> {
    if ["formulaType", "formulaOutput", "spill"]
        .iter()
        .any(|key| cell.get(*key).is_some_and(|value| !value.is_null()))
        || cell["scenarioResult"] == true
    {
        return Err(crate::coded(
            "fill-unsupported",
            "Across Worksheets special, spill or scenario cells are not qualified",
        ));
    }
    Ok(())
}
fn physical_cell(
    root: &Path,
    db: &Connection,
    sheet: &str,
    at: (u32, u32),
    formula: bool,
) -> Result<()> {
    let attributes = crate::spill::source_attributes(root, db, sheet, at)?;
    if attributes.as_ref().is_some_and(|attrs| {
        attrs.iter().any(|(key, value)| {
            !["r", "s", "t", "xmlns"].contains(&key.as_str())
                || key == "xmlns" && value.as_bytes() != crate::xml::NS
        })
    }) {
        return Err(crate::coded(
            "fill-unsupported",
            "Across Worksheets extended source cell attributes are not qualified",
        ));
    }
    let original_formula: bool = db.prepare_cached("SELECT EXISTS(SELECT 1 FROM cells WHERE sheet=?1 AND r=?2 AND c=?3 AND formula IS NOT NULL)").map_err(crate::err)?.query_row(params![sheet, at.0, at.1], |row| row.get(0)).map_err(err)?;
    crate::spill::source_content_proof(root, db, sheet, at, original_formula || formula)
}
/// The copied range is finite, but unknown formulas elsewhere can obscure the
/// affected dependency set. Refuse them instead of using a viewport inventory.
fn dependencies(db: &Connection, version: u64, metadata: &Value) -> Result<()> {
    let mut scanned = 0;
    for sheet in metadata["sheets"].as_array().ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing Across Worksheets reference inventory",
        )
    })? {
        let path = field(sheet, "path")?;
        let name = field(sheet, "name")?;
        let mut after: Option<String> = None;
        loop {
            crate::cancel::check()?;
            let page = super::read_db(
                db,
                &json!({"version":version,"sheet":path,"range":"A1:XFD1048576","formulas":true,"after":after,"limit":512}),
            )?;
            let cells = page["cells"].as_array().ok_or_else(|| {
                crate::coded(
                    "internal-invalid-field",
                    "Invalid Across Worksheets formula inventory",
                )
            })?;
            for pair in cells {
                scanned += 1;
                if scanned > FORMULA_LIMIT {
                    return Err(crate::coded(
                        "fill-limit",
                        "Across Worksheets dependency inventory exceeds its bound",
                    ));
                }
                let refs = crate::formula::auditing::references(
                    metadata,
                    name,
                    pair[0].as_str().ok_or_else(|| {
                        crate::coded(
                            "internal-invalid-field",
                            "Invalid Across Worksheets formula address",
                        )
                    })?,
                    &pair[1],
                    false,
                )?;
                // A literal broken reference has a closed error meaning. It
                // does not hide an unknown dependency target, even inside IFERROR.
                let broken_errors_only = !refs.markers.is_empty()
                    && refs.markers.iter().all(|marker| {
                        marker["kind"] == "unsupported"
                            && marker["reason"] == "Formula contains a broken#REF! reference"
                    });
                if refs.unknown && !broken_errors_only
                    || refs.ranges.iter().any(|range| range["via"] != "formula")
                {
                    return Err(crate::coded(
                        "fill-unsupported",
                        "Across Worksheets named, table, dynamic, external or unknown dependencies are not qualified",
                    ));
                }
            }
            if page["nextCursor"].is_null() {
                break;
            }
            after = Some(field(&page, "nextCursor")?.into());
        }
    }
    Ok(())
}
fn copy_style(cell: &mut Value, delta: &mut Value, source: &Value, styles: u64) -> Result<()> {
    let style = source["styleId"].as_u64().ok_or_else(|| {
        crate::coded(
            "internal-invalid-field",
            "Across Worksheets source style identity is invalid",
        )
    })?;
    if style >= styles {
        return Err(crate::coded(
            "fill-unsupported",
            "Across Worksheets source style is outside the accepted style inventory",
        ));
    }
    for key in STYLE_FIELDS {
        cell.as_object_mut()
            .ok_or_else(|| crate::coded("fill-malformed", "Invalid Across Worksheets target cell"))?
            .remove(*key);
    }
    cell.as_object_mut().unwrap().remove("format");
    cell["styleId"] = json!(style);
    delta["styleId"] = json!(style);
    for key in STYLE_FIELDS {
        if let Some(value) = source["$edit"].get(*key) {
            delta[*key] = value.clone();
        }
        if let Some(value) = source.get(*key) {
            cell[*key] = value.clone();
        }
    }
    if let Some(value) = source.get("format") {
        cell["format"] = value.clone();
    }
    Ok(())
}
pub(super) fn apply(
    root: &Path,
    db: &Connection,
    parent: u64,
    command: &Value,
    index: &Value,
) -> Result<Outcome> {
    let (sheet, source, area, mode) = descriptor(command)?;
    let styles = fill_guard::across_metadata(root, &[sheet, source])?;
    for path in [sheet, source] {
        axes(db, index, path, area)?;
        if crate::spill::present(db)? {
            crate::spill::guard(
                db,
                parent,
                path,
                &crate::spill::Geometry {
                    first: area.0,
                    last: area.1,
                },
            )?;
        }
        super::special_groups(db, parent, path, area, false)?;
    }
    let contents = mode != "formats";
    fill_guard::across_shared_strings(root, db, parent, source, area)?;
    if !contents {
        fill_guard::across_shared_strings(root, db, parent, sheet, area)?;
    }
    if contents {
        dependencies(db, parent, &crate::package_edit::auditing_metadata(root)?)?;
    }
    db.execute("INSERT INTO versions(parent) VALUES(?1)", [parent])
        .map_err(err)?;
    let version = db.last_insert_rowid();
    let mut cache = VecDeque::new();
    let mut formula_delta = 0;
    let mut formula_touched = false;
    for row in (area.0.0 / 32 * 32..=area.1.0).step_by(32) {
        crate::cancel::check()?;
        for col in (area.0.1 / 32 * 32..=area.1.1).step_by(32) {
            let mut targets = tile(db, parent, sheet, row, col)?;
            for r in row.max(area.0.0)..=(row + 31).min(area.1.0) {
                for c in col.max(area.0.1)..=(col + 31).min(area.1.1) {
                    crate::cancel::check()?;
                    let source_cell = source_cell(&mut cache, db, parent, source, r, c)?;
                    ordinary(&source_cell)?;
                    let mut cell = targets
                        .remove(&address(r, c))
                        .unwrap_or(json!({"value":null,"styleId":0}));
                    ordinary(&cell)?;
                    let mut edit = inherited_edit(db, parent, sheet, r, c)?;
                    let mut delta = json!({});
                    physical_cell(root, db, source, (r, c), source_cell["formula"].is_string())?;
                    physical_cell(root, db, sheet, (r, c), cell["formula"].is_string())?;
                    if contents {
                        let was_formula = cell["formula"].is_string();
                        if let Some(formula) = source_cell["formula"].as_str() {
                            let source_edit = &source_cell["$edit"];
                            let attributed: bool = db.prepare_cached("SELECT EXISTS(SELECT 1 FROM cells WHERE sheet=?1 AND r=?2 AND c=?3 AND formula_attrs IS NOT NULL AND formula_attrs<>'{}')").map_err(crate::err)?.query_row(params![source, r, c], |row| row.get(0)).map_err(err)?;
                            if attributed && source_edit.get("formula").is_none() {
                                return Err(crate::coded(
                                    "fill-unsupported",
                                    "Across Worksheets attributed formulas are not qualified for content copying",
                                ));
                            }
                            // Same-address copying has no coordinate displacement.
                            // Keep the complete authored expression byte for byte.
                            set_value(&mut cell, Value::Null);
                            cell["formula"] = json!(formula);
                            delta["formula"] = json!(formula);
                        } else {
                            let literal = values::literal_snapshot(&source_cell)?;
                            delta["value"] = literal.value.clone();
                            set_value(&mut cell, literal.value);
                            if literal.error {
                                delta["valueType"] = json!("error");
                                cell["dataType"] = json!("error");
                                cell["error"] = cell["value"].clone();
                            }
                        }
                        let is_formula = cell["formula"].is_string();
                        formula_delta += i64::from(is_formula) - i64::from(was_formula);
                        formula_touched |= was_formula || is_formula;
                    }
                    if mode != "contents" {
                        copy_style(&mut cell, &mut delta, &source_cell, styles)?;
                    }
                    merge_edit(&mut edit, &delta)?;
                    store_change(db, version, sheet, r, c, &cell, &edit)?;
                }
            }
        }
    }
    Ok(Outcome {
        version: version as u64,
        formula_delta,
        formula_touched,
        content_changed: contents,
    })
}

#[cfg(test)]
mod tests;
