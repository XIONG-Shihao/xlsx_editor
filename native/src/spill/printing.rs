//! Whole-worksheet array ownership admission precedes bounded print pages.
use super::*;
use std::collections::BTreeMap;
pub(crate) fn qualify(
    root: &Path,
    db: &Connection,
    version: u64,
    sheet: &str,
    view: &Value,
    print_area: &Geometry,
) -> Result<BTreeMap<String, Value>> {
    // An omitted worksheet must not skip validate_current's constructor walk
    // while borrowing retained output cells from the requested path below.
    let sheets = view["sheets"].as_array().ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing print ownership worksheets",
        )
    })?;
    if sheets.iter().filter(|item| item["path"] == sheet).count() != 1 {
        return Err(crate::coded(
            "print-unsupported",
            "PRINT_SPECIAL_FORMULA: print ownership worksheet identity is incomplete",
        ));
    }
    if any_history(db, version)? {
        validate_current(root, db, &json!({"version":version,"view":view}))?;
    } else {
        crate::calculation::snapshot(db, version)?;
    }
    let owners = effective(db, version, sheet)?;
    let mut qualified = BTreeMap::new();
    for owner in &owners {
        if owner.proof["status"] != "calculated"
            || owner.proof["qualification"] != "retained-sequence-v1"
            || !retained_arguments(&json!(validate_constructor(owner)?.arguments))
        {
            return Err(crate::coded(
                "print-unsupported",
                "PRINT_SPECIAL_FORMULA: only complete calculated retained numeric spills are qualified",
            ));
        }
        validate_output_cells(db, version, owner)?;
        // Recheck imported cm-to-XLDAPR source identity instead of trusting a
        // caller's top-left cache or a hidden master outside the page.
        if let Some(attrs) = source_attributes(root, db, sheet, owner.master)? {
            if let Some(cm) = attrs.get("cm") {
                let cm = cm.parse::<u64>().map_err(err)?;
                let mut map = serde_json::Map::new();
                map.insert(
                    format!("{sheet}\0{}", address(owner.master.0, owner.master.1)),
                    json!({"cm":cm,"range":owner.geometry.reference()}),
                );
                metadata::validate(
                    &metadata::source(root)?,
                    &Default::default(),
                    &Value::Object(map),
                )?;
            }
        }
        for at in owner.geometry.positions() {
            if print_area.contains(at) {
                let reference = address(at.0, at.1);
                qualified.insert(reference,json!({"kind":if at==owner.master{"master"}else{"follower"},"spill":owner.cell_metadata(at)?}));
            }
        }
    }
    let context = if any_history(db, version)? {
        Some(consumers::Context::new(root, db, version, view)?)
    } else {
        None
    };
    let mut scalar_context = scalars::Context::new(root, db, version)?;
    let mut after = None;
    let mut scanned = 0;
    loop {
        let page = crate::calculation::page(db, sheet, (0, 0), (1048575, 16383), after, true)?;
        for (reference, cell) in &page {
            crate::cancel::check()?;
            scanned += 1;
            if scanned > 100000 {
                return Err(crate::coded(
                    "print-limit",
                    "PRINT_FORMULA_PROOF_BOUND: worksheet formula proof exceeds100,000 positions",
                ));
            }
            let at = position(reference)?;
            if owners.iter().any(|owner| owner.master == at) {
                continue;
            }
            if context
                .as_ref()
                .map(|context| context.validate(sheet, reference, cell))
                .transpose()?
                .flatten()
                .is_some()
            {
                if print_area.contains(at) {
                    qualified.insert(
                        reference.clone(),
                        json!({"kind":"consumer","calculation":cell["calculation"]}),
                    );
                }
                continue;
            }
            if scalar_context.validate(sheet, reference, cell)?.is_some() {
                if print_area.contains(at) {
                    qualified.insert(
                        reference.clone(),
                        json!({"kind":"consumer","calculation":cell["calculation"]}),
                    );
                }
                continue;
            }
            // Closed authored consumers prove a scalar shape even if shrink
            // makes their current result unresolved. Visible unresolved cells
            // still refuse; imported singleton cm requires the full proof.
            if cell.get("formulaType").is_none()
                && cell["calculation"]["status"] == "unresolved"
                && cell["formula"].as_str().is_some_and(|formula| {
                    matches!(crate::formula::spill_consumer(formula), Ok(Some(_)))
                })
            {
                continue;
            }
            if cell.get("formulaType").is_some_and(|kind| kind != "shared")
                || cell.get("scenarioResult").is_some()
                || cell["formula"]
                    .as_str()
                    .is_none_or(|formula| !crate::formula::scalar_print_formula(formula))
            {
                if cell.get("formulaType").is_none()
                    && cell["calculation"]["status"] == "unresolved"
                {
                    return Err(crate::coded(
                        "print-unsupported",
                        format!(
                            "PRINT_UNRESOLVED_CALCULATION: unqualified current provider at {reference}"
                        ),
                    ));
                }
                return Err(crate::coded(
                    "print-unsupported",
                    format!(
                        "PRINT_SPECIAL_FORMULA: unqualified current array/name/provider at {reference}"
                    ),
                ));
            }
        }
        if page.len() < 512 {
            break;
        }
        after = Some(position(&page.last().unwrap().0)?);
    }
    Ok(qualified)
}
