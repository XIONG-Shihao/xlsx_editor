//! What-If Data Tables: bounded virtual input substitutions; never mutate inputs.
use crate::{
    Result, err, field,
    xml::{address, position},
};
use rusqlite::{Connection, params};
use serde_json::{Value, json};
pub(crate) type Inputs = (String, String, String, Option<String>, Option<String>);
pub(crate) fn prepare(db: &Connection, view: &Value) -> Result<()> {
    db.execute_batch("CREATE TEMP TABLE calc_scenarios(sheet TEXT,r INTEGER,c INTEGER,model TEXT,input1 TEXT,value1 TEXT,input2 TEXT,value2 TEXT,PRIMARY KEY(sheet,r,c)) WITHOUT ROWID;").map_err(err)?;
    let mut total = 0u64;
    for sheet in view["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheets"))?
    {
        let path = field(sheet, "path")?;
        let present:bool=db.prepare_cached("SELECT EXISTS(SELECT 1 FROM cells WHERE sheet=?1 AND formula IS NOT NULL AND formula_attrs IS NOT NULL AND json_extract(formula_attrs,'$.t')='dataTable')").map_err(crate::err)?.query_row([path],|row|row.get(0)).map_err(err)?;
        if !present {
            continue;
        }

        let mut after = None;
        loop {
            let cells = crate::calculation::page(db, path, (0, 0), (1048575, 16383), after, true)?;
            for (reference, cell) in &cells {
                if cell["formulaType"] != "dataTable" {
                    continue;
                }
                let p = position(reference)?;
                let range = field(cell, "formulaReference")?;
                let mut parts = range.split(':');
                let first = position(parts.next().unwrap())?;
                let last = parts.next().map(position).transpose()?.unwrap_or(first);
                if parts.next().is_some()
                    || p != first
                    || first.0 == 0
                    || first.1 == 0
                    || last.0 < first.0
                    || last.1 < first.1
                {
                    return Err(crate::coded(
                        "formula-limit",
                        "Unsupported What-If table bounds",
                    ));
                }
                let raw: String = db
                    .query_row(
                        "SELECT formula_attrs FROM cells WHERE sheet=?1 AND r=?2 AND c=?3",
                        params![path, p.0, p.1],
                        |r| r.get(0),
                    )
                    .map_err(err)?;
                let attrs: Value = serde_json::from_str(&raw).map_err(err)?;
                if attrs.as_object().is_none_or(|a| {
                    a.keys().any(|k| {
                        !["t", "ref", "r1", "r2", "dt2D", "dtr", "ca", "del1", "del2"]
                            .contains(&k.as_str())
                    })
                }) {
                    return Err(crate::coded(
                        "formula-unsupported",
                        "Extended What-If table attributes",
                    ));
                }
                let flag = |key: &str| -> Result<bool> {
                    match attrs[key].as_str() {
                        None | Some("0" | "false") => Ok(false),
                        Some("1" | "true") => Ok(true),
                        _ => Err(crate::coded(
                            "internal-invalid-field",
                            "Invalid What-If flag",
                        )),
                    }
                };
                if flag("del1")? || flag("del2")? {
                    return Err(crate::coded("formula-unsupported", "Deleted What-If input"));
                }
                let two = flag("dt2D")?;
                let row = flag("dtr")?;
                let input1 = position(field(&attrs, "r1")?)?;
                let input2 = if two {
                    Some(position(field(&attrs, "r2")?)?)
                } else {
                    None
                };
                for input in [Some(input1), input2].into_iter().flatten() {
                    if input.0 >= first.0
                        && input.0 <= last.0
                        && input.1 >= first.1
                        && input.1 <= last.1
                    {
                        return Err(crate::coded(
                            "formula-conflict",
                            "What-If input overlaps results",
                        ));
                    }
                }
                if input2 == Some(input1) {
                    return Err(crate::coded(
                        "formula-bad-input",
                        "What-If inputs must be distinct",
                    ));
                }
                total += u64::from(last.0 - first.0 + 1) * u64::from(last.1 - first.1 + 1);
                if total > 100000 {
                    return Err(crate::coded(
                        "formula-limit",
                        "What-If scenario work budget",
                    ));
                }
                for r in first.0..=last.0 {
                    for c in first.1..=last.1 {
                        crate::cancel::check()?;
                        let model = if two {
                            (first.0 - 1, first.1 - 1)
                        } else if row {
                            (r, first.1 - 1)
                        } else {
                            (first.0 - 1, c)
                        };
                        let value1 = if two || row {
                            (first.0 - 1, c)
                        } else {
                            (r, first.1 - 1)
                        };
                        let value2 = two.then(|| address(r, first.1 - 1));
                        let mut output = crate::calculation::cell(db, path, &address(r, c))?;
                        if !output.is_object() {
                            return Err(crate::coded(
                                "formula-bad-input",
                                "Sparse What-If result groups require compatibility",
                            ));
                        }
                        if (r, c) != first
                            && (output["formula"].is_string()
                                || output.get("formulaType").is_some())
                        {
                            return Err(crate::coded(
                                "formula-conflict",
                                "Overlapping What-If formula groups",
                            ));
                        }
                        output["scenarioResult"] = json!(true);
                        db.prepare_cached("INSERT INTO calc_overlay(sheet,r,c,data,edit) VALUES(?1,?2,?3,?4,'{}') ON CONFLICT(sheet,r,c) DO UPDATE SET data=excluded.data").map_err(err)?.execute(params![path,r,c,output.to_string()]).map_err(err)?;
                        db.prepare_cached(
                            "INSERT INTO calc_scenarios VALUES(?1,?2,?3,?4,?5,?6,?7,?8)",
                        )
                        .map_err(err)?
                        .execute(params![
                            path,
                            r,
                            c,
                            address(model.0, model.1),
                            address(input1.0, input1.1),
                            address(value1.0, value1.1),
                            input2.map(|x| address(x.0, x.1)),
                            value2
                        ])
                        .map_err(err)?;
                    }
                }
            }
            if cells.len() < 512 {
                break;
            }
            after = Some(position(&cells.last().unwrap().0)?);
        }
    }
    Ok(())
}
pub(crate) fn dependencies(
    db: &Connection,
    view: &Value,
    sheet: &str,
    reference: &str,
) -> Result<Value> {
    let path = view["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheets"))?
        .iter()
        .find(|s| s["name"] == sheet)
        .and_then(|s| s["path"].as_str())
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing scenario sheet"))?;
    let (r, c) = position(reference)?;
    let (model,input1,value1,input2,value2):Inputs=db.prepare_cached("SELECT model,input1,value1,input2,value2 FROM calc_scenarios WHERE sheet=?1 AND r=?2 AND c=?3").map_err(crate::err)?.query_row(params![path,r,c],|row|Ok((row.get(0)?,row.get(1)?,row.get(2)?,row.get(3)?,row.get(4)?))).map_err(err)?;
    let ranges: Vec<Value> = [Some(model), Some(input1), Some(value1), input2, value2]
        .into_iter()
        .flatten()
        .map(|reference| json!({"sheet":sheet,"ref":reference}))
        .collect();
    Ok(json!({"sheet":sheet,"ref":reference,"ranges":ranges}))
}
