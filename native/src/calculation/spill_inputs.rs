//! Evaluation-only literals from complete native-owned spill output records.
use super::*;

pub(super) fn prepare(db: &Connection, version: u64, view: &Value) -> Result<()> {
    for sheet in view["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheets"))?
    {
        let path = field(sheet, "path")?;
        for group in crate::spill::effective(db, version, path)? {
            crate::cancel::check()?;
            if group.proof["qualification"] != "retained-sequence-v1" {
                continue;
            }
            let reference = address(group.master.0, group.master.1);
            let master = cell(db, path, &reference)?;
            // accepted_cell checks the complete current group's numeric/error
            // values and authored identities once, rather than once per output.
            if !crate::spill::accepted_cell(db, version, path, group.master, &master)? {
                return Err(crate::coded(
                    "spill-bad-input",
                    "Accepted spill master has no native ownership proof",
                ));
            }
            let page = crate::range::read_db(
                db,
                &json!({"version":version,"sheet":path,"range":group.geometry.reference(),"limit":crate::spill::MAX_OUTPUTS}),
            )?;
            for pair in page["cells"].as_array().ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing accepted spill outputs")
            })? {
                let (r, c) = position(pair[0].as_str().ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing accepted spill address")
                })?)?;
                let source = &pair[1];
                let mut literal = json!({"value":source["value"],"dataType":source["dataType"]});
                for key in ["rawValue", "error"] {
                    if let Some(value) = source.get(key) {
                        literal[key] = value.clone();
                    }
                }
                db.execute(
                    "INSERT INTO calc_owned_spills VALUES(?1,?2,?3,?4)",
                    params![path, r, c, literal.to_string()],
                )
                .map_err(err)?;
            }
        }
    }
    Ok(())
}

pub(super) fn owned(db: &Connection, sheet: &str, reference: &str) -> Result<bool> {
    let (r, c) = position(reference)?;
    db.prepare_cached(
        "SELECT EXISTS(SELECT 1 FROM calc_owned_spills WHERE sheet=?1 AND r=?2 AND c=?3)",
    )
    .map_err(crate::err)?
    .query_row(params![sheet, r, c], |row| row.get(0))
    .map_err(err)
}
