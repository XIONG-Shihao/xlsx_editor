//! Accepted-version facts for native coordinators. Imported counts alone become
//! stale when a literal is replaced by a formula or an edit extends a worksheet.
use crate::{Result, calculation, database, err, field, local};
use serde_json::{Value, json};
use std::{fs::File, io::BufReader, path::Path};
pub fn read(root: &Path, input: &Value) -> Result<Value> {
    let version = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing summary version"))?;
    let db = database(root)?;
    if !db
        .query_row(
            "SELECT EXISTS(SELECT 1 FROM versions WHERE id=?1)",
            [version],
            |row| row.get::<_, bool>(0),
        )
        .map_err(err)?
    {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown summary version",
        ));
    }
    calculation::snapshot(&db, version)?;
    let index: Value = serde_json::from_reader(BufReader::new(
        File::open(local(root, "index.json")?).map_err(err)?,
    ))
    .map_err(err)?;
    let mut sheets = Vec::new();
    for sheet in index["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing indexed sheets"))?
    {
        crate::cancel::check()?;
        let path = field(sheet, "path")?;
        let formulas:u64=db.prepare_cached("SELECT (SELECT count(*) FROM cells c WHERE c.sheet=?1 AND c.formula IS NOT NULL AND NOT EXISTS(SELECT 1 FROM calc_overlay o WHERE o.sheet=c.sheet AND o.r=c.r AND o.c=c.c))+(SELECT count(*) FROM calc_overlay WHERE sheet=?1 AND (json_extract(data,'$.formula') IS NOT NULL OR json_extract(data,'$.formulaType') IS NOT NULL))").map_err(crate::err)?.query_row([path],|row|row.get(0)).map_err(err)?;
        let (row, column): (Option<u32>, Option<u32>) = db
            .query_row(
                "SELECT max(r),max(c) FROM calc_overlay WHERE sheet=?1",
                [path],
                |row| Ok((row.get(0)?, row.get(1)?)),
            )
            .map_err(err)?;
        let mut rows = sheet["rows"]
            .as_u64()
            .ok_or_else(|| crate::coded("workbook-malformed", "Invalid indexed rows"))?
            .max(row.map(|r| u64::from(r) + 1).unwrap_or(0));
        let mut columns = sheet["columns"]
            .as_u64()
            .ok_or_else(|| crate::coded("workbook-malformed", "Invalid indexed columns"))?
            .max(column.map(|c| u64::from(c) + 1).unwrap_or(0));
        let merges = crate::merges::selected_override(&db, version, path)?;
        if let Some(ranges) = &merges {
            for range in ranges {
                let (_, last) = crate::bulk::bounds(range)?;
                rows = rows.max(u64::from(last.0) + 1);
                columns = columns.max(u64::from(last.1) + 1);
            }
        }
        let mut facts = json!({"path":path,"rows":rows,"columns":columns,"formulaCount":formulas});
        if let Some(merges) = merges {
            facts["merges"] = json!(merges);
        }
        sheets.push(facts);
    }
    Ok(json!({"version":version,"sheets":sheets}))
}
