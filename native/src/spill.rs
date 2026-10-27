//! Versioned spill ownership primitives. Constructors remain gated separately.
//! A scalar value is never sufficient proof of an array's output ownership.
use crate::xml::{address, position};
use crate::{Result, err};
use rusqlite::{Connection, OptionalExtension, params};
use serde_json::{Value, json};
use std::{
    io::{Read, Seek, SeekFrom},
    path::Path,
};
pub(crate) mod consumers;
pub(crate) mod metadata;
pub(crate) mod printing;
pub(crate) mod scalars;

pub(crate) const SCHEMA: u32 = 1;
pub(crate) const MAX_OUTPUTS: u32 = 1_000;
const MAX_EXACT_INTEGER: f64 = 9_007_199_254_740_991.0;

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct Geometry {
    pub first: (u32, u32),
    pub last: (u32, u32),
}
impl Geometry {
    pub fn from_reference(reference: &str) -> Result<Self> {
        let mut pieces = reference.split(':');
        let first = position(
            pieces
                .next()
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing spill range"))?,
        )?;
        let last = pieces.next().map(position).transpose()?.unwrap_or(first);
        if pieces.next().is_some() || first.0 > last.0 || first.1 > last.1 {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid spill rectangle",
            ));
        }
        Ok(Self { first, last })
    }
    pub fn count(&self) -> u64 {
        u64::from(self.last.0 - self.first.0 + 1) * u64::from(self.last.1 - self.first.1 + 1)
    }
    pub fn reference(&self) -> String {
        let first = address(self.first.0, self.first.1);
        if self.first == self.last {
            first
        } else {
            format!("{first}:{}", address(self.last.0, self.last.1))
        }
    }
    pub fn contains(&self, at: (u32, u32)) -> bool {
        self.first.0 <= at.0 && at.0 <= self.last.0 && self.first.1 <= at.1 && at.1 <= self.last.1
    }
    pub fn intersects(&self, other: &Self) -> bool {
        self.first.0 <= other.last.0
            && other.first.0 <= self.last.0
            && self.first.1 <= other.last.1
            && other.first.1 <= self.last.1
    }
    pub fn positions(&self) -> impl Iterator<Item = (u32, u32)> + '_ {
        (self.first.0..=self.last.0)
            .flat_map(move |r| (self.first.1..=self.last.1).map(move |c| (r, c)))
    }
}

#[derive(Clone, Debug)]
pub(crate) struct Sequence {
    pub geometry: Geometry,
    pub arguments: [f64; 4],
    pub values: Vec<Value>,
}
impl Sequence {
    /// Exact integer slice only. This does not implement Excel coercions or
    /// certify arbitrary numeric arguments outside independently proved inputs.
    pub fn new(master: (u32, u32), supplied: &[f64]) -> Result<Self> {
        if !(1..=4).contains(&supplied.len()) {
            return Err(crate::coded(
                "spill-unsupported",
                "SEQUENCE requires one through four qualified scalar arguments",
            ));
        }
        let mut arguments = [1.0; 4];
        for (index, value) in supplied.iter().copied().enumerate() {
            if !value.is_finite()
                || value.fract() != 0.0
                || value.abs() > MAX_EXACT_INTEGER
                || value == 0.0 && value.is_sign_negative()
            {
                return Err(crate::coded(
                    "spill-unsupported",
                    "SEQUENCE argument is outside the exact integer slice",
                ));
            }
            arguments[index] = value;
        }
        if arguments[0] < 1.0
            || arguments[1] < 1.0
            || arguments[0] > f64::from(MAX_OUTPUTS)
            || arguments[1] > f64::from(MAX_OUTPUTS)
            || arguments[0] * arguments[1] > f64::from(MAX_OUTPUTS)
        {
            return Err(crate::coded(
                "spill-limit",
                "SEQUENCE shape is outside the qualified output bound",
            ));
        }
        let (rows, columns) = (arguments[0] as u32, arguments[1] as u32);
        let last = (
            master
                .0
                .checked_add(rows - 1)
                .ok_or_else(|| crate::coded("spill-limit", "Spill row overflow"))?,
            master
                .1
                .checked_add(columns - 1)
                .ok_or_else(|| crate::coded("spill-limit", "Spill column overflow"))?,
        );
        if last.0 >= 1_048_576 || last.1 >= 16_384 {
            return Err(crate::coded(
                "spill-limit",
                "SEQUENCE output exceeds worksheet bounds",
            ));
        }
        let geometry = Geometry {
            first: master,
            last,
        };
        let mut values = Vec::with_capacity(geometry.count() as usize);
        // i128 verifies the exact integer progression before conversion; f64
        // arithmetic alone could silently round a boundary value back in-range.
        let (start, step) = (arguments[2] as i128, arguments[3] as i128);
        for index in 0..geometry.count() {
            let delta = i128::from(index) * step;
            if delta.abs() > MAX_EXACT_INTEGER as i128 {
                return Err(crate::coded(
                    "spill-unsupported",
                    "SEQUENCE progression requires an unqualified inexact intermediate",
                ));
            }
            let exact = start + delta;
            if exact.abs() > MAX_EXACT_INTEGER as i128 {
                return Err(crate::coded(
                    "spill-bad-input",
                    "SEQUENCE output is not an exact finite integer",
                ));
            }
            values.push(json!(exact as f64));
        }
        Ok(Self {
            geometry,
            arguments,
            values,
        })
    }
}

#[derive(Clone, Debug)]
pub(crate) struct Group {
    pub version: u64,
    pub sheet: String,
    pub master: (u32, u32),
    pub geometry: Geometry,
    pub formula: String,
    pub proof: Value,
}
impl Group {
    pub fn cell_metadata(&self, at: (u32, u32)) -> Result<Value> {
        if !self.geometry.contains(at) {
            return Err(crate::coded(
                "spill-unsupported",
                "Cell is outside its spill owner",
            ));
        }
        let mut metadata = json!({"schema":SCHEMA,"master":address(self.master.0,self.master.1),"range":self.geometry.reference(),"role":if at==self.master{"master"}else{"follower"}});
        if self.proof["status"] == "blocked" {
            metadata["intendedRange"] = self.proof["intendedRange"].clone();
        }
        Ok(metadata)
    }
}

pub(crate) fn schema(db: &Connection) -> Result<()> {
    db.execute_batch("CREATE TABLE IF NOT EXISTS spill_groups(version INTEGER NOT NULL,sheet TEXT NOT NULL,mr INTEGER NOT NULL,mc INTEGER NOT NULL,active INTEGER NOT NULL CHECK(active IN(0,1)),r0 INTEGER NOT NULL,c0 INTEGER NOT NULL,r1 INTEGER NOT NULL,c1 INTEGER NOT NULL,formula TEXT NOT NULL,proof TEXT NOT NULL,schema INTEGER NOT NULL,PRIMARY KEY(version,sheet,mr,mc)) WITHOUT ROWID; CREATE INDEX IF NOT EXISTS spill_version_sheet ON spill_groups(version,sheet);").map_err(err)
}
pub(crate) fn present(db: &Connection) -> Result<bool> {
    db.prepare_cached(
        "SELECT EXISTS(SELECT 1 FROM sqlite_master WHERE type='table' AND name='spill_groups')",
    )
    .map_err(crate::err)?
    .query_row([], |row| row.get(0))
    .map_err(err)
}
pub(crate) fn any_active(db: &Connection, version: u64) -> Result<bool> {
    known_version(db, version)?;
    if !present(db)? {
        return Ok(false);
    }
    db.prepare_cached("WITH RECURSIVE ancestry(id,depth) AS (SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN ancestry a ON v.id=a.id WHERE v.parent IS NOT NULL), ranked AS (SELECT g.active,row_number() OVER(PARTITION BY g.sheet,g.mr,g.mc ORDER BY a.depth) priority FROM ancestry a JOIN spill_groups g ON g.version=a.id) SELECT EXISTS(SELECT 1 FROM ranked WHERE priority=1 AND active=1)").map_err(crate::err)?.query_row([version],|row|row.get(0)).map_err(err)
}
pub(crate) fn any_history(db: &Connection, version: u64) -> Result<bool> {
    known_version(db, version)?;
    if !present(db)? {
        return Ok(false);
    }
    db.prepare_cached("WITH RECURSIVE ancestry(id) AS (SELECT ?1 UNION ALL SELECT v.parent FROM versions v JOIN ancestry a ON v.id=a.id WHERE v.parent IS NOT NULL) SELECT EXISTS(SELECT 1 FROM ancestry a JOIN spill_groups g ON g.version=a.id)").map_err(crate::err)?.query_row([version],|row|row.get(0)).map_err(err)
}
pub(crate) fn protected_input(db: &Connection, sheet: &str, at: (u32, u32)) -> Result<bool> {
    db.prepare_cached("SELECT EXISTS(SELECT 1 FROM spill_protected WHERE sheet=?1 AND r0<=?2 AND r1>=?2 AND c0<=?3 AND c1>=?3)").map_err(crate::err)?.query_row(params![sheet,at.0,at.1],|row|row.get(0)).map_err(err)
}
fn known_version(db: &Connection, version: u64) -> Result<()> {
    if db
        .prepare_cached("SELECT EXISTS(SELECT 1 FROM versions WHERE id=?1)")
        .map_err(crate::err)?
        .query_row([version], |r| r.get::<_, bool>(0))
        .map_err(err)?
    {
        Ok(())
    } else {
        Err(crate::coded(
            "internal-unknown-reference",
            "Unknown spill version",
        ))
    }
}
pub(crate) fn effective(db: &Connection, version: u64, sheet: &str) -> Result<Vec<Group>> {
    known_version(db, version)?;
    if !present(db)? {
        return Ok(Vec::new());
    }
    let mut statement=db.prepare_cached("WITH RECURSIVE ancestry(id,depth) AS (SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN ancestry a ON v.id=a.id WHERE v.parent IS NOT NULL), ranked AS (SELECT g.*,row_number() OVER(PARTITION BY g.sheet,g.mr,g.mc ORDER BY a.depth) priority FROM ancestry a JOIN spill_groups g ON g.version=a.id WHERE g.sheet=?2) SELECT version,mr,mc,r0,c0,r1,c1,formula,proof,schema FROM ranked WHERE priority=1 AND active=1 ORDER BY mr,mc").map_err(err)?;
    let mut rows = statement.query(params![version, sheet]).map_err(err)?;
    let mut result = Vec::new();
    while let Some(row) = rows.next().map_err(err)? {
        crate::cancel::check()?;
        if row.get::<_, u32>(9).map_err(err)? != SCHEMA {
            return Err(crate::coded(
                "spill-unsupported",
                "Unsupported spill ownership schema",
            ));
        }
        let master = (
            row.get::<_, u32>(1).map_err(err)?,
            row.get::<_, u32>(2).map_err(err)?,
        );
        let geometry = Geometry {
            first: (row.get(3).map_err(err)?, row.get(4).map_err(err)?),
            last: (row.get(5).map_err(err)?, row.get(6).map_err(err)?),
        };
        if geometry.first != master
            || geometry.first.0 > geometry.last.0
            || geometry.first.1 > geometry.last.1
            || geometry.last.0 >= 1_048_576
            || geometry.last.1 >= 16_384
            || geometry.count() > u64::from(MAX_OUTPUTS)
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid persisted spill geometry",
            ));
        }
        let group = Group {
            version: row.get(0).map_err(err)?,
            sheet: sheet.into(),
            master,
            geometry,
            formula: row.get(7).map_err(err)?,
            proof: serde_json::from_str(&row.get::<_, String>(8).map_err(err)?).map_err(err)?,
        };
        validate_constructor(&group)?;
        if result
            .iter()
            .any(|old: &Group| old.geometry.intersects(&group.geometry))
        {
            return Err(crate::coded(
                "spill-conflict",
                "Overlapping persisted spill ownership",
            ));
        }
        result.push(group);
    }
    Ok(result)
}

/// Caller owns the transaction and output-cell persistence. There is no API
/// that commits ownership separately from its complete output changes.
pub(crate) fn record(db: &Connection, group: &Group, active: bool) -> Result<()> {
    if db.is_autocommit() {
        return Err(crate::coded(
            "spill-bad-input",
            "Spill ownership records require an owned transaction",
        ));
    }
    known_version(db, group.version)?;
    if group.geometry.first != group.master
        || Geometry::from_reference(&group.geometry.reference())? != group.geometry
        || group.geometry.count() > u64::from(MAX_OUTPUTS)
        || !group.proof.is_object()
        || group.formula.is_empty()
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid spill publication proof",
        ));
    }
    if active {
        let sequence = validate_constructor(group)?;
        let blocked = group.proof["status"] == "blocked";
        if blocked {
            let blocker = position(group.proof["blocker"].as_str().ok_or_else(|| {
                crate::coded(
                    "spill-conflict",
                    "Blocked spill proof has no literal collision",
                )
            })?)?;
            if blocker == group.master || !sequence.geometry.contains(blocker) {
                return Err(crate::coded(
                    "spill-conflict",
                    "Blocked spill proof has an invalid literal collision position",
                ));
            }
            let parent: u64 = db
                .prepare_cached("SELECT parent FROM versions WHERE id=?1")
                .map_err(crate::err)?
                .query_row([group.version], |row| row.get(0))
                .map_err(err)?;
            let page = crate::range::read_db(
                db,
                &json!({"version":parent,"sheet":group.sheet,"range":address(blocker.0,blocker.1),"limit":1}),
            )?;
            let cell = &page["cells"][0][1];
            if cell["value"].is_null()
                || ["formula", "formulaType", "spill"]
                    .iter()
                    .any(|key| cell.get(*key).is_some())
            {
                return Err(crate::coded(
                    "spill-conflict",
                    "Blocked spill proof has no occupied literal in its parent version",
                ));
            }
        }
        for owner in effective(db, group.version, &group.sheet)? {
            if owner.master != group.master && owner.geometry.intersects(&sequence.geometry) {
                return Err(crate::coded(
                    "spill-conflict",
                    "A spill constructor overlaps another active owner",
                ));
            }
        }
        for (index, at) in group.geometry.positions().enumerate() {
            let raw: Option<String> = db
                .prepare_cached(
                    "SELECT data FROM changes WHERE version=?1 AND sheet=?2 AND r=?3 AND c=?4",
                )
                .map_err(crate::err)?
                .query_row(params![group.version, group.sheet, at.0, at.1], |row| {
                    row.get(0)
                })
                .optional()
                .map_err(err)?;
            let cell: Value = serde_json::from_str(&raw.ok_or_else(|| {
                crate::coded(
                    "spill-refused",
                    "Spill ownership has an unpublished output position",
                )
            })?)
            .map_err(err)?;
            let typed = if blocked {
                cell["calculation"]["status"] == "error"
                    && cell["calculation"]["error"] == "#SPILL!"
                    && cell["value"] == "#SPILL!"
                    && cell["dataType"] == "error"
                    && cell["rawValue"] == "#VALUE!"
            } else {
                cell["calculation"]["status"] == "calculated"
                    && cell["calculation"]["value"] == sequence.values[index]
                    && cell["value"] == sequence.values[index]
                    && cell["dataType"] == "number"
                    && cell.get("error").is_none()
            };
            let authored = if at == group.master {
                cell["formula"] == group.formula
                    && cell["formulaType"] == "array"
                    && cell["formulaReference"] == group.geometry.reference()
            } else {
                cell.get("formula").is_none()
                    && cell.get("formulaType").is_none()
                    && cell["value"] == sequence.values[index]
                    && cell["dataType"] == "number"
            };
            if cell["spill"] != group.cell_metadata(at)? || !typed || !authored {
                return Err(crate::coded(
                    "spill-unsupported",
                    "Spill ownership output proof is incomplete or inconsistent",
                ));
            }
        }
    } else {
        let parent: u64 = db
            .prepare_cached("SELECT parent FROM versions WHERE id=?1")
            .map_err(crate::err)?
            .query_row([group.version], |row| row.get(0))
            .map_err(err)?;
        let previous = effective(db, parent, &group.sheet)?
            .into_iter()
            .find(|owner| owner.master == group.master)
            .ok_or_else(|| {
                crate::coded(
                    "spill-bad-input",
                    "Spill release has no active parent owner",
                )
            })?;
        if previous.geometry != group.geometry
            || previous.formula != group.formula
            || previous.proof != group.proof
        {
            return Err(crate::coded(
                "spill-bad-input",
                "Spill release does not identify the complete previous owner",
            ));
        }
        for at in group.geometry.positions() {
            let raw: Option<(String, String)> = db
                .prepare_cached(
                    "SELECT data,edit FROM changes WHERE version=?1 AND sheet=?2 AND r=?3 AND c=?4",
                )
                .map_err(crate::err)?
                .query_row(params![group.version, group.sheet, at.0, at.1], |row| {
                    Ok((row.get(0)?, row.get(1)?))
                })
                .optional()
                .map_err(err)?;
            let (raw, edit) = raw.ok_or_else(|| {
                crate::coded(
                    "spill-refused",
                    "Spill release has an unpublished output tombstone",
                )
            })?;
            let cell: Value = serde_json::from_str(&raw).map_err(err)?;
            let edit: Value = serde_json::from_str(&edit).map_err(err)?;
            if !cell["value"].is_null()
                || cell["dataType"] != "blank"
                || ["formula", "formulaType", "spill", "calculation"]
                    .iter()
                    .any(|key| cell.get(*key).is_some())
                || edit["$spillRelease"] != true
            {
                return Err(crate::coded(
                    "spill-unsupported",
                    "Spill release has an incomplete output tombstone proof",
                ));
            }
        }
    }
    db.execute(
        "INSERT INTO spill_groups VALUES(?1,?2,?3,?4,?5,?6,?7,?8,?9,?10,?11,?12)",
        params![
            group.version,
            group.sheet,
            group.master.0,
            group.master.1,
            u32::from(active),
            group.geometry.first.0,
            group.geometry.first.1,
            group.geometry.last.0,
            group.geometry.last.1,
            group.formula,
            group.proof.to_string(),
            SCHEMA
        ],
    )
    .map_err(err)?;
    Ok(())
}
fn validate_constructor(group: &Group) -> Result<Sequence> {
    if group.proof["constructor"] != "sequence-integer-v1" {
        return Err(crate::coded(
            "spill-unsupported",
            "Unsupported spill constructor proof",
        ));
    }
    let arguments = group.proof["arguments"]
        .as_array()
        .ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing exact spill constructor proof",
            )
        })?
        .iter()
        .map(|arg| {
            arg.as_f64().ok_or_else(|| {
                crate::coded(
                    "internal-invalid-field",
                    "Invalid spill constructor argument proof",
                )
            })
        })
        .collect::<Result<Vec<_>>>()?;
    let sequence = Sequence::new(group.master, &arguments)?;
    if group.proof["status"] == "blocked" {
        if group.geometry.first != group.geometry.last
            || group.proof["intendedRange"] != sequence.geometry.reference()
            || group.proof["errorType"] != 8
            || group.proof["subType"] != 1
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid blocked spill proof",
            ));
        }
        let blocker = position(group.proof["blocker"].as_str().ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing blocked spill literal position",
            )
        })?)?;
        if blocker == group.master || !sequence.geometry.contains(blocker) {
            return Err(crate::coded(
                "spill-unsupported",
                "Blocked spill literal is outside its intended output",
            ));
        }
    } else if group.proof.get("status").is_some() && group.proof["status"] != "calculated"
        || sequence.geometry != group.geometry
    {
        return Err(crate::coded(
            "spill-bad-input",
            "Spill ownership geometry does not match constructor proof",
        ));
    }
    Ok(sequence)
}
fn validate_output_cells(db: &Connection, version: u64, group: &Group) -> Result<()> {
    let sequence = validate_constructor(group)?;
    let blocked = group.proof["status"] == "blocked";
    let page = crate::range::read_db(
        db,
        &json!({"version":version,"sheet":group.sheet,"range":group.geometry.reference(),"limit":MAX_OUTPUTS}),
    )?;
    let cells = page["cells"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing spill output cells"))?;
    if cells.len() as u64 != group.geometry.count() {
        return Err(crate::coded(
            "spill-refused",
            "Accepted spill group has missing output positions",
        ));
    }
    for (index, pair) in cells.iter().enumerate() {
        let at = position(pair[0].as_str().ok_or_else(|| {
            crate::coded("internal-missing-field", "Missing spill output address")
        })?)?;
        let cell = &pair[1];
        let typed = if blocked {
            cell["value"] == "#SPILL!"
                && cell["dataType"] == "error"
                && cell["rawValue"] == "#VALUE!"
                && cell["calculation"]["status"] == "error"
                && cell["calculation"]["error"] == "#SPILL!"
        } else {
            cell["value"] == sequence.values[index]
                && cell["dataType"] == "number"
                && cell["calculation"]["status"] == "calculated"
                && cell["calculation"]["value"] == sequence.values[index]
                && cell.get("error").is_none()
        };
        let authored = if at == group.master {
            cell["formula"] == group.formula
                && cell["formulaType"] == "array"
                && cell["formulaReference"] == group.geometry.reference()
        } else {
            cell.get("formula").is_none() && cell.get("formulaType").is_none()
        };
        if !typed || !authored || cell["spill"] != group.cell_metadata(at)? {
            return Err(crate::coded(
                "spill-stale",
                "Accepted spill output no longer agrees with its complete ownership proof",
            ));
        }
    }
    Ok(())
}
pub(crate) fn accepted_cell(
    db: &Connection,
    version: u64,
    sheet: &str,
    at: (u32, u32),
    cell: &Value,
) -> Result<bool> {
    let Some(metadata) = cell.get("spill") else {
        return Ok(false);
    };
    if !present(db)? {
        return Err(crate::coded(
            "spill-bad-input",
            "Spill cell has no native ownership table",
        ));
    }
    let master = position(
        metadata["master"]
            .as_str()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing spill cell owner"))?,
    )?;
    let owner = effective(db, version, sheet)?
        .into_iter()
        .find(|owner| owner.master == master)
        .ok_or_else(|| crate::coded("spill-bad-input", "Spill cell has no active native owner"))?;
    if metadata != &owner.cell_metadata(at)? {
        return Err(crate::coded(
            "internal-mismatch",
            "Spill cell metadata differs from accepted owner",
        ));
    }
    validate_output_cells(db, version, &owner)?;
    Ok(true)
}
pub(crate) fn export_groups(db: &Connection, version: u64, view: &Value) -> Result<Vec<Value>> {
    let mut groups = Vec::new();
    for sheet in view["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing spill export worksheets"))?
    {
        let path = crate::field(sheet, "path")?;
        for group in effective(db, version, path)? {
            if group.proof["qualification"] != "retained-sequence-v1"
                || !retained_arguments(&json!(validate_constructor(&group)?.arguments))
            {
                return Err(crate::coded(
                    "spill-unsupported",
                    "Spill export is outside the retained actual constructor slice",
                ));
            }
            validate_output_cells(db, version, &group)?;
            let mut item = json!({"schema":SCHEMA,"sheet":path,"master":address(group.master.0,group.master.1),"range":group.geometry.reference(),"status":if group.proof["status"]=="blocked"{"error"}else{"calculated"}});
            if group.proof["status"] == "blocked" {
                item["error"] = json!("#SPILL!");
                item["intendedRange"] = group.proof["intendedRange"].clone();
            }
            groups.push(item);
        }
    }
    Ok(groups)
}

/// Refuses a content change that touches an accepted spill: no edit path
/// releases one yet.
pub(crate) fn guard(db: &Connection, version: u64, sheet: &str, target: &Geometry) -> Result<()> {
    if effective(db, version, sheet)?
        .iter()
        .any(|group| group.geometry.intersects(target))
    {
        return Err(crate::coded(
            "spill-unsupported",
            "A spill follower cannot be edited independently of its master",
        ));
    }
    Ok(())
}

fn released_master(db: &Connection, version: u64, sheet: &str, master: (u32, u32)) -> Result<bool> {
    if !present(db)? {
        return Ok(false);
    }
    let row:Option<(bool,u32)>=db.prepare_cached("WITH RECURSIVE ancestry(id,depth) AS (SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN ancestry a ON v.id=a.id WHERE v.parent IS NOT NULL) SELECT g.active,g.schema FROM ancestry a JOIN spill_groups g ON g.version=a.id WHERE g.sheet=?2 AND g.mr=?3 AND g.mc=?4 ORDER BY a.depth LIMIT 1").map_err(crate::err)?.query_row(params![version,sheet,master.0,master.1],|row|Ok((row.get(0)?,row.get(1)?))).optional().map_err(err)?;
    match row {
        Some((_, schema)) if schema != SCHEMA => Err(crate::coded(
            "spill-unsupported",
            "Unsupported released spill ownership schema",
        )),
        Some((active, _)) => Ok(!active),
        None => Ok(false),
    }
}

/// The caller will edit only after this private version has released the whole
/// owner. No API can release a follower or retain its old cached output cells.
pub(crate) fn release_master(root: &Path, input: &Value) -> Result<Value> {
    let mut db = crate::database(root)?;
    let tx = db.transaction().map_err(err)?;
    let parent = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing spill release version"))?;
    let sheet = crate::field(input, "sheet")?;
    let master = position(crate::field(input, "cell")?)?;
    let mut owner = effective(&tx, parent, sheet)?
        .into_iter()
        .find(|owner| owner.geometry.contains(master))
        .ok_or_else(|| {
            crate::coded(
                "spill-bad-input",
                "Selected cell has no accepted spill owner",
            )
        })?;
    if owner.master != master {
        return Err(crate::coded(
            "spill-unsupported",
            "A spill follower cannot release its master's output area",
        ));
    }
    let page = crate::range::read_db(
        &tx,
        &json!({"version":parent,"sheet":sheet,"range":owner.geometry.reference(),"limit":MAX_OUTPUTS}),
    )?;
    if page["cells"]
        .as_array()
        .ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing complete spill release page",
            )
        })?
        .len() as u64
        != owner.geometry.count()
    {
        return Err(crate::coded(
            "spill-unsupported",
            "Spill release has incomplete output positions",
        ));
    }
    tx.execute("INSERT INTO versions(parent) VALUES(?1)", [parent])
        .map_err(err)?;
    let version = tx.last_insert_rowid() as u64;
    let geometry = owner.geometry.reference();
    for pair in page["cells"].as_array().unwrap() {
        crate::cancel::check()?;
        let at = position(pair[0].as_str().ok_or_else(|| {
            crate::coded("internal-missing-field", "Missing spill release address")
        })?)?;
        if pair[1]["spill"] != owner.cell_metadata(at)? {
            return Err(crate::coded(
                "spill-stale",
                "Spill release output no longer agrees with its accepted owner",
            ));
        }
        let mut edit = crate::range::inherited_edit(&tx, parent, sheet, at.0, at.1)?;
        crate::range::merge_edit(&mut edit, &json!({"value":null}))?;
        edit["$spillRelease"] = json!(true);
        crate::range::store_calculated_change(
            &tx,
            version as i64,
            sheet,
            at.0,
            at.1,
            &blank_tombstone(&pair[1]),
            &edit,
        )?;
    }
    owner.version = version;
    record(&tx, &owner, false)?;
    tx.commit().map_err(err)?;
    Ok(
        json!({"version":version,"schema":SCHEMA,"master":address(master.0,master.1),"releasedRange":geometry,"enabled":false}),
    )
}

/// Retain destination formatting while preventing inherited cached output from
/// reappearing after shrink or clear. The tombstone has no formula/owner/result.
pub(crate) fn blank_tombstone(cell: &Value) -> Value {
    let mut blank = json!({"value":null,"dataType":"blank"});
    for key in crate::range::STYLE_FIELDS
        .iter()
        .copied()
        .chain(["styleId", "format"])
    {
        if let Some(value) = cell.get(key) {
            blank[key] = value.clone();
        }
    }
    blank
}

/// Persist only inside the caller's transaction, after collision/argument proof.
/// Snapshot readers still address the immutable parent while outputs and old
/// geometry tombstones are written to a new, unexposed calculation version.
pub(crate) fn publish(
    db: &Connection,
    group: &Group,
    previous: Option<&Geometry>,
    revision: &str,
) -> Result<()> {
    if db.is_autocommit() {
        return Err(crate::coded(
            "spill-bad-input",
            "Spill publication requires an owned transaction",
        ));
    }
    if previous.is_some_and(|geometry| geometry.count() > u64::from(MAX_OUTPUTS)) {
        return Err(crate::coded(
            "spill-limit",
            "Previous spill geometry exceeds ownership bound",
        ));
    }
    let parent: u64 = db
        .prepare_cached("SELECT parent FROM versions WHERE id=?1")
        .map_err(crate::err)?
        .query_row([group.version], |row| row.get(0))
        .map_err(err)?;
    let sequence = validate_constructor(group)?;
    let blocked = group.proof["status"] == "blocked";
    schema(db)?;
    if let Some(previous) = previous {
        for at in previous
            .positions()
            .filter(|at| !group.geometry.contains(*at))
        {
            crate::cancel::check()?;
            let old = crate::calculation::cell(db, &group.sheet, &address(at.0, at.1))?;
            let blank = blank_tombstone(&old);
            let mut edit = crate::range::inherited_edit(db, parent, &group.sheet, at.0, at.1)?;
            crate::range::merge_edit(&mut edit, &json!({"value":null}))?;
            edit["$spillRelease"] = json!(true);
            crate::range::store_calculated_change(
                db,
                group.version as i64,
                &group.sheet,
                at.0,
                at.1,
                &blank,
                &edit,
            )?;
        }
    }
    for (index, at) in group.geometry.positions().enumerate() {
        crate::cancel::check()?;
        let source = crate::calculation::cell(db, &group.sheet, &address(at.0, at.1))?;
        if at != group.master
            && (source["formula"].is_string()
                || source.get("formulaType").is_some()
                || !previous.is_some_and(|geometry| geometry.contains(at))
                    && !source["value"].is_null())
        {
            return Err(crate::coded(
                "spill-conflict",
                "Spill publication has an unproved occupied destination",
            ));
        }
        let mut cell = if source.is_object() {
            source
        } else {
            blank_tombstone(&Value::Null)
        };
        cell.as_object_mut().unwrap().remove("$sourceFields");
        cell["spill"] = group.cell_metadata(at)?;
        if at == group.master {
            cell["formula"] = json!(group.formula);
            cell["formulaType"] = json!("array");
            cell["formulaReference"] = json!(group.geometry.reference());
            cell["value"] = if blocked {
                json!("#SPILL!")
            } else {
                sequence.values[index].clone()
            };
            cell["dataType"] = json!(if blocked { "error" } else { "number" });
            cell["rawValue"] = json!(if blocked {
                "#VALUE!".to_string()
            } else {
                sequence.values[index].to_string()
            });
            if blocked {
                cell["error"] = json!("#SPILL!");
            } else {
                cell.as_object_mut().unwrap().remove("error");
            }
        } else {
            for key in [
                "formula",
                "formulaType",
                "formulaReference",
                "sharedFormulaId",
                "error",
            ] {
                cell.as_object_mut().unwrap().remove(key);
            }
            cell["value"] = sequence.values[index].clone();
            cell["dataType"] = json!("number");
            cell["rawValue"] = json!(sequence.values[index].to_string());
            cell["cachedValue"] = Value::Null;
        }
        cell["calculation"] = if blocked {
            json!({"status":"error","error":"#SPILL!","revision":revision})
        } else {
            json!({"status":"calculated","value":sequence.values[index],"revision":revision})
        };
        let edit = crate::range::inherited_edit(db, parent, &group.sheet, at.0, at.1)?;
        crate::range::store_calculated_change(
            db,
            group.version as i64,
            &group.sheet,
            at.0,
            at.1,
            &cell,
            &edit,
        )?;
    }
    record(db, group, true)
}

/// Parse the actual OOXML metadata indirection. A `cm` attribute is a one-based
/// cellMetadata index; its rc/@t selects a one-based metadata type and rc/@v a
/// zero-based futureMetadata block. These three indexes are not interchangeable.
pub(crate) fn dynamic_cell_metadata(xml: &str) -> Result<std::collections::HashSet<u32>> {
    use quick_xml::{NsReader, events::Event, name::ResolveResult};
    const DYNAMIC: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray";
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut buffer = Vec::new();
    let mut stack: Vec<String> = Vec::new();
    let mut types: Vec<(String, bool)> = Vec::new();
    let mut future: Vec<Option<bool>> = Vec::new();
    let mut records: Vec<Vec<(usize, usize)>> = Vec::new();
    let mut future_name = None;
    let mut future_seen = false;
    let mut types_seen = false;
    let mut records_seen = false;
    let mut declared_types = None;
    let mut declared_future = None;
    let mut declared_records = None;
    loop {
        let (namespace, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
        let empty = matches!(&event, Event::Empty(_));
        match event {
            Event::DocType(_) => {
                return Err(crate::coded(
                    "spill-malformed",
                    "XML entities are unsupported",
                ));
            }
            Event::Start(element) | Event::Empty(element) => {
                let standard =
                    namespace == ResolveResult::Bound(quick_xml::name::Namespace(crate::xml::NS));
                let local = std::str::from_utf8(element.local_name().as_ref())
                    .map_err(err)?
                    .to_owned();
                let attributes = crate::xml::attrs(&element)?;
                let canonical = |value: &str| -> Result<usize> {
                    if value.is_empty()
                        || value.len() > 1 && value.starts_with('0')
                        || !value.bytes().all(|byte| byte.is_ascii_digit())
                    {
                        return Err(crate::coded(
                            "spill-malformed",
                            "Noncanonical dynamic metadata index",
                        ));
                    }
                    value.parse::<usize>().map_err(err)
                };
                let count = || {
                    attributes
                        .get("count")
                        .map(|value| canonical(value))
                        .transpose()
                };
                if future_name.as_deref() == Some("XLDAPR") && stack.len() >= 2 {
                    let valid = match stack.len() {
                        2 => standard && local == "bk",
                        3 => standard && local == "extLst",
                        4 => {
                            standard
                                && local == "ext"
                                && attributes.get("uri").map(String::as_str)
                                    == Some("{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}")
                        }
                        5 => {
                            namespace == ResolveResult::Bound(quick_xml::name::Namespace(DYNAMIC))
                                && local == "dynamicArrayProperties"
                        }
                        _ => false,
                    };
                    if !valid
                        || stack.len() < 5
                            && attributes.keys().any(|key| {
                                !(stack.len() == 4 && key == "uri"
                                    || key == "xmlns"
                                    || key.starts_with("xmlns:"))
                            })
                    {
                        return Err(crate::coded(
                            "spill-unsupported",
                            "Extended dynamic future metadata requires qualification",
                        ));
                    }
                }
                if stack.is_empty() {
                    if !standard || local != "metadata" {
                        return Err(crate::coded(
                            "spill-malformed",
                            "Invalid dynamic metadata root",
                        ));
                    }
                } else if standard {
                    match local.as_str() {
                        "metadataTypes" if stack.len() == 1 => {
                            if types_seen {
                                return Err(crate::coded(
                                    "spill-duplicate",
                                    "Duplicate metadata types",
                                ));
                            }
                            types_seen = true;
                            declared_types = count()?;
                            if declared_types.is_none() {
                                return Err(crate::coded(
                                    "internal-missing-field",
                                    "Missing metadata type count",
                                ));
                            }
                        }
                        "metadataType" if stack.last().is_some_and(|s| s == "metadataTypes") => {
                            let name = attributes
                                .get("name")
                                .ok_or_else(|| {
                                    crate::coded(
                                        "internal-missing-field",
                                        "Missing metadata type name",
                                    )
                                })?
                                .clone();
                            if types.iter().any(|kind| kind.0 == name)
                                || attributes
                                    .get("cellMeta")
                                    .is_some_and(|value| !["0", "1"].contains(&value.as_str()))
                            {
                                return Err(crate::coded(
                                    "spill-ambiguous",
                                    "Ambiguous dynamic metadata type",
                                ));
                            }
                            types
                                .push((name, attributes.get("cellMeta").is_some_and(|v| v == "1")));
                        }
                        "futureMetadata" if stack.len() == 1 => {
                            let name = attributes
                                .get("name")
                                .ok_or_else(|| {
                                    crate::coded(
                                        "internal-missing-field",
                                        "Missing future metadata name",
                                    )
                                })?
                                .clone();
                            if name == "XLDAPR" {
                                if future_seen {
                                    return Err(crate::coded(
                                        "spill-duplicate",
                                        "Duplicate dynamic future metadata",
                                    ));
                                }
                                future_seen = true;
                                declared_future = count()?;
                                if declared_future.is_none() {
                                    return Err(crate::coded(
                                        "internal-missing-field",
                                        "Missing dynamic future count",
                                    ));
                                }
                            }
                            future_name = Some(name);
                        }
                        "cellMetadata" if stack.len() == 1 => {
                            if records_seen {
                                return Err(crate::coded(
                                    "spill-duplicate",
                                    "Duplicate cell metadata",
                                ));
                            }
                            records_seen = true;
                            declared_records = count()?;
                            if declared_records.is_none() {
                                return Err(crate::coded(
                                    "internal-missing-field",
                                    "Missing cell metadata count",
                                ));
                            }
                        }
                        "bk" if stack.last().is_some_and(|s| s == "futureMetadata")
                            && future_name.as_deref() == Some("XLDAPR") =>
                        {
                            future.push(None)
                        }
                        "bk" if stack.last().is_some_and(|s| s == "cellMetadata") => {
                            records.push(Vec::new())
                        }
                        "rc" if stack.last().is_some_and(|s| s == "bk")
                            && stack
                                .get(stack.len().saturating_sub(2))
                                .is_some_and(|s| s == "cellMetadata") =>
                        {
                            let t = canonical(attributes.get("t").ok_or_else(|| {
                                crate::coded(
                                    "internal-missing-field",
                                    "Missing metadata type index",
                                )
                            })?)?;
                            let v = canonical(attributes.get("v").ok_or_else(|| {
                                crate::coded(
                                    "internal-missing-field",
                                    "Missing future metadata index",
                                )
                            })?)?;
                            records
                                .last_mut()
                                .ok_or_else(|| {
                                    crate::coded(
                                        "internal-missing-field",
                                        "Missing cell metadata record",
                                    )
                                })?
                                .push((t, v));
                        }
                        _ => (),
                    }
                } else if namespace == ResolveResult::Bound(quick_xml::name::Namespace(DYNAMIC))
                    && local == "dynamicArrayProperties"
                    && future_name.as_deref() == Some("XLDAPR")
                {
                    if stack.as_slice() != ["metadata", "futureMetadata", "bk", "extLst", "ext"] {
                        return Err(crate::coded(
                            "spill-unsupported",
                            "Extended dynamic metadata nesting requires qualification",
                        ));
                    }
                    let value = attributes
                        .get("fDynamic")
                        .map(String::as_str)
                        .unwrap_or("0");
                    if !["0", "1"].contains(&value) {
                        return Err(crate::coded(
                            "internal-invalid-field",
                            "Invalid dynamic-array flag",
                        ));
                    }
                    let block = future.last_mut().ok_or_else(|| {
                        crate::coded("internal-missing-field", "Missing dynamic metadata block")
                    })?;
                    if block.is_some() {
                        return Err(crate::coded(
                            "spill-duplicate",
                            "Duplicate dynamic-array properties",
                        ));
                    }
                    let collapsed = attributes
                        .get("fCollapsed")
                        .map(String::as_str)
                        .unwrap_or("0");
                    if !["0", "1"].contains(&collapsed)
                        || attributes.keys().any(|key| {
                            !["fDynamic", "fCollapsed"].contains(&key.as_str())
                                && key != "xmlns"
                                && !key.starts_with("xmlns:")
                        })
                    {
                        return Err(crate::coded(
                            "spill-unsupported",
                            "Extended dynamic-array properties require qualification",
                        ));
                    }
                    *block = Some(value == "1" && collapsed == "0");
                }
                if !empty {
                    stack.push(local);
                } else if local == "futureMetadata" {
                    future_name = None;
                }
            }
            Event::End(_) => {
                if stack.pop().as_deref() == Some("futureMetadata") {
                    future_name = None;
                }
            }
            Event::Eof => break,
            _ => (),
        }
        buffer.clear();
    }
    if !stack.is_empty()
        || declared_types.is_some_and(|count| count != types.len())
        || declared_future.is_some_and(|count| count != future.len())
        || declared_records.is_some_and(|count| count != records.len())
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid metadata counts or nesting",
        ));
    }
    let mut eligible = std::collections::HashSet::new();
    for (index, record) in records.iter().enumerate() {
        // Multiple records can contain other metadata types, but duplicate
        // XLDAPR entries are ambiguous ownership and therefore refused.
        let mut dynamic = None;
        for &(t, v) in record {
            let kind = types
                .get(t.checked_sub(1).ok_or_else(|| {
                    crate::coded("spill-malformed", "Invalid one-based metadata type index")
                })?)
                .ok_or_else(|| crate::coded("spill-malformed", "Invalid metadata type index"))?;
            if kind.0 == "XLDAPR" {
                if dynamic.is_some() || !kind.1 {
                    return Err(crate::coded(
                        "spill-ambiguous",
                        "Ambiguous dynamic cell metadata",
                    ));
                }
                dynamic = Some(future.get(v).and_then(|value| *value).ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing dynamic-array properties")
                })?);
            }
        }
        if dynamic.is_some() && record.len() != 1 {
            return Err(crate::coded(
                "spill-unsupported",
                "Composite dynamic cell metadata requires preservation qualification",
            ));
        }
        if dynamic == Some(true) {
            eligible.insert(u32::try_from(index + 1).map_err(err)?);
        }
    }
    Ok(eligible)
}

pub(crate) fn source_attributes(
    root: &Path,
    db: &Connection,
    sheet: &str,
    at: (u32, u32),
) -> Result<Option<std::collections::BTreeMap<String, String>>> {
    use rusqlite::OptionalExtension;
    let location: Option<(u64, u64)> = db
        .prepare_cached("SELECT start,open FROM cells WHERE sheet=?1 AND r=?2 AND c=?3")
        .map_err(crate::err)?
        .query_row(params![sheet, at.0, at.1], |row| {
            Ok((row.get(0)?, row.get(1)?))
        })
        .optional()
        .map_err(err)?;
    let Some((start, open)) = location else {
        return Ok(None);
    };
    if open < start || open - start > 65536 {
        return Err(crate::coded(
            "spill-limit",
            "Source cell opening exceeds spill proof bound",
        ));
    }
    let index: Value =
        serde_json::from_slice(&std::fs::read(root.join("index.json")).map_err(err)?)
            .map_err(err)?;
    let artifact = index["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing native sheet index"))?
        .iter()
        .find(|item| item["path"] == sheet)
        .and_then(|item| item["artifact"].as_str())
        .ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing indexed worksheet artifact",
            )
        })?;
    let mut source = std::fs::File::open(crate::local(root, artifact)?).map_err(err)?;
    source.seek(SeekFrom::Start(start)).map_err(err)?;
    let mut bytes = vec![0; usize::try_from(open - start).map_err(err)?];
    source.read_exact(&mut bytes).map_err(err)?;
    let mut reader = quick_xml::Reader::from_reader(bytes.as_slice());
    match reader.read_event().map_err(err)? {
        quick_xml::events::Event::Start(element) | quick_xml::events::Event::Empty(element)
            if element.local_name().as_ref() == b"c" =>
        {
            Ok(Some(crate::xml::attrs(&element)?.into_iter().collect()))
        }
        _ => Err(crate::coded(
            "spill-malformed",
            "Invalid indexed source cell opening",
        )),
    }
}
/// The indexed value alone omits unknown child XML. Adoption may rewrite the
/// recognized f/v/is fragments, so prove that the source cell has no other
/// content or attributes hidden between them. Rich inline text stays protected.
pub(crate) fn source_content_proof(
    root: &Path,
    db: &Connection,
    sheet: &str,
    at: (u32, u32),
    formula: bool,
) -> Result<()> {
    use quick_xml::{NsReader, events::Event, name::ResolveResult};
    let location: Option<(u64, u64)> = db
        .prepare_cached("SELECT start,end FROM cells WHERE sheet=?1 AND r=?2 AND c=?3")
        .map_err(crate::err)?
        .query_row(params![sheet, at.0, at.1], |row| {
            Ok((row.get(0)?, row.get(1)?))
        })
        .optional()
        .map_err(err)?;
    let Some((start, end)) = location else {
        return Ok(());
    };
    if end < start || end - start > 65536 {
        return Err(crate::coded(
            "spill-limit",
            "Source cell content exceeds spill proof bound",
        ));
    }
    let index: Value =
        serde_json::from_slice(&std::fs::read(root.join("index.json")).map_err(err)?)
            .map_err(err)?;
    let artifact = index["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing native sheet index"))?
        .iter()
        .find(|item| item["path"] == sheet)
        .and_then(|item| item["artifact"].as_str())
        .ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing indexed worksheet artifact",
            )
        })?;
    let mut file = std::fs::File::open(crate::local(root, artifact)?).map_err(err)?;
    file.seek(SeekFrom::Start(start)).map_err(err)?;
    let mut bytes = vec![0; usize::try_from(end - start).map_err(err)?];
    file.read_exact(&mut bytes).map_err(err)?;
    let text = std::str::from_utf8(&bytes).map_err(err)?;
    // Source worksheet prefixes are already refused by the indexed `extended`
    // guard; the synthetic wrapper supplies the inherited default namespace.
    let wrapped = format!(
        "<worksheet xmlns=\"{}\">{text}</worksheet>",
        String::from_utf8_lossy(crate::xml::NS)
    );
    let mut reader = NsReader::from_reader(wrapped.as_bytes());
    let mut buffer = Vec::new();
    let mut stack: Vec<String> = Vec::new();
    let mut seen = std::collections::HashSet::new();
    loop {
        let (ns, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
        let empty = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if ns != ResolveResult::Bound(quick_xml::name::Namespace(crate::xml::NS)) {
                    return Err(crate::coded(
                        "spill-unsupported",
                        "Opaque source cell namespace requires preservation qualification",
                    ));
                }
                let name = std::str::from_utf8(element.local_name().as_ref())
                    .map_err(err)?
                    .to_owned();
                let attributes = crate::xml::attrs(&element)?;
                let allowed = match stack.len() {
                    0 => name == "worksheet",
                    1 => name == "c",
                    2 => ["v", "is"].contains(&name.as_str()) || formula && name == "f",
                    3 => stack.last().is_some_and(|parent| parent == "is") && name == "t",
                    _ => false,
                };
                if !allowed || stack.len() == 2 && !seen.insert(name.clone()) {
                    return Err(crate::coded(
                        "spill-unsupported",
                        "Opaque or duplicate source cell children require preservation qualification",
                    ));
                }
                if stack.len() >= 2
                    && attributes.iter().any(|(key, value)| {
                        !(key == "xmlns" && value.as_bytes() == crate::xml::NS
                            || name == "f" && ["t", "ref", "aca", "ca"].contains(&key.as_str())
                            || name == "t"
                                && key == "xml:space"
                                && ["preserve", "default"].contains(&value.as_str()))
                    })
                {
                    return Err(crate::coded(
                        "spill-unsupported",
                        "Opaque source value attributes require preservation qualification",
                    ));
                }
                if !empty {
                    stack.push(name);
                }
            }
            Event::End(_) => {
                stack.pop().ok_or_else(|| {
                    crate::coded("spill-malformed", "Unbalanced source cell content")
                })?;
            }
            Event::Text(text) => {
                if !stack
                    .last()
                    .is_some_and(|name| ["f", "v", "t"].contains(&name.as_str()))
                    && !text.unescape().map_err(err)?.trim().is_empty()
                {
                    return Err(crate::coded(
                        "spill-unsupported",
                        "Opaque mixed source cell content requires preservation qualification",
                    ));
                }
            }
            Event::CData(_) | Event::Comment(_) | Event::PI(_) | Event::DocType(_) => {
                return Err(crate::coded(
                    "spill-unsupported",
                    "Opaque source cell content requires preservation qualification",
                ));
            }
            Event::Eof => break,
            _ => (),
        }
        buffer.clear();
    }
    if !stack.is_empty() {
        return Err(crate::coded(
            "spill-malformed",
            "Unbalanced source cell content",
        ));
    }
    Ok(())
}
fn metadata_proof(root: &Path) -> Result<std::collections::HashSet<u32>> {
    let mut zip =
        zip::ZipArchive::new(std::fs::File::open(root.join("metadata.xlsx")).map_err(err)?)
            .map_err(err)?;
    let mut xml = String::new();
    {
        let mut part = match zip.by_name("xl/metadata.xml") {
            Ok(part) => part,
            Err(zip::result::ZipError::FileNotFound) => return Ok(Default::default()),
            Err(error) => return Err(err(error)),
        };
        if part.size() > 8 * 1024 * 1024 {
            return Err(crate::coded(
                "spill-limit",
                "Dynamic metadata exceeds the qualification bound",
            ));
        }
        part.read_to_string(&mut xml).map_err(err)?;
    }
    fn entries(
        xml: &str,
        root_name: &str,
        namespace: &[u8],
        child: &str,
    ) -> Result<Vec<std::collections::BTreeMap<String, String>>> {
        use quick_xml::{NsReader, events::Event, name::ResolveResult};
        let mut reader = NsReader::from_reader(xml.as_bytes());
        let mut buffer = Vec::new();
        let mut depth = 0;
        let mut result = Vec::new();
        loop {
            let (ns, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
            let empty = matches!(&event, Event::Empty(_));
            match event {
                Event::DocType(_) => {
                    return Err(crate::coded(
                        "spill-malformed",
                        "XML entities are unsupported",
                    ));
                }
                Event::Start(element) | Event::Empty(element) => {
                    let standard =
                        ns == ResolveResult::Bound(quick_xml::name::Namespace(namespace));
                    let local = std::str::from_utf8(element.local_name().as_ref())
                        .map_err(err)?
                        .to_owned();
                    if depth == 0 && (!standard || local != root_name) {
                        return Err(crate::coded(
                            "internal-invalid-field",
                            "Invalid dynamic metadata package graph",
                        ));
                    }
                    if depth == 1 && standard && local == child {
                        result.push(crate::xml::attrs(&element)?.into_iter().collect());
                    }
                    if !empty {
                        depth += 1;
                    }
                }
                Event::End(_) => depth -= 1,
                Event::Eof => break,
                _ => (),
            }
            buffer.clear();
        }
        Ok(result)
    }
    let mut relationships = String::new();
    zip.by_name("xl/_rels/workbook.xml.rels")
        .map_err(err)?
        .read_to_string(&mut relationships)
        .map_err(err)?;
    let relationships = entries(
        &relationships,
        "Relationships",
        b"http://schemas.openxmlformats.org/package/2006/relationships",
        "Relationship",
    )?;
    let mut ids = std::collections::HashSet::new();
    if relationships.iter().any(|item| {
        item.get("Id")
            .is_none_or(|id| id.is_empty() || !ids.insert(id.clone()))
    }) {
        return Err(crate::coded(
            "spill-ambiguous",
            "Ambiguous dynamic metadata relationship identifiers",
        ));
    }
    let matching=relationships.into_iter().filter(|item|item.get("Type").is_some_and(|value|value=="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata")).collect::<Vec<_>>();
    if matching.len() != 1
        || matching[0].get("Target").map(String::as_str) != Some("metadata.xml")
        || matching[0]
            .get("TargetMode")
            .is_some_and(|value| value != "Internal")
    {
        return Err(crate::coded(
            "spill-unsupported",
            "Dynamic metadata has no unique qualified workbook relationship",
        ));
    }
    let mut types = String::new();
    zip.by_name("[Content_Types].xml")
        .map_err(err)?
        .read_to_string(&mut types)
        .map_err(err)?;
    let types = entries(
        &types,
        "Types",
        b"http://schemas.openxmlformats.org/package/2006/content-types",
        "Override",
    )?;
    let mut parts = std::collections::HashSet::new();
    if types.iter().any(|item| {
        item.get("PartName")
            .is_none_or(|part| part.is_empty() || !parts.insert(part.clone()))
    }) {
        return Err(crate::coded(
            "spill-ambiguous",
            "Ambiguous dynamic metadata content type identifiers",
        ));
    }
    let matching = types
        .into_iter()
        .filter(|item| {
            item.get("PartName")
                .is_some_and(|value| value == "/xl/metadata.xml")
        })
        .collect::<Vec<_>>();
    if matching.len() != 1
        || matching[0].get("ContentType").map(String::as_str)
            != Some("application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml")
    {
        return Err(crate::coded(
            "spill-unsupported",
            "Dynamic metadata has no unique qualified content type",
        ));
    }
    dynamic_cell_metadata(&xml)
}
/// Original package geometry is authoritative even if a caller omits it from
/// its current view. Current layout/table geometry is fenced separately below.
fn source_barriers(root: &Path, sheet: &str) -> Result<Vec<Geometry>> {
    use quick_xml::{NsReader, events::Event, name::ResolveResult};
    let mut zip =
        zip::ZipArchive::new(std::fs::File::open(root.join("metadata.xlsx")).map_err(err)?)
            .map_err(err)?;
    let mut xml = String::new();
    {
        let mut part = zip.by_name(sheet).map_err(err)?;
        if part.size() > 8 * 1024 * 1024 {
            return Err(crate::coded(
                "spill-limit",
                "Source worksheet metadata exceeds spill proof bound",
            ));
        }
        part.read_to_string(&mut xml).map_err(err)?;
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut buffer = Vec::new();
    let mut stack: Vec<String> = Vec::new();
    let mut rectangles = Vec::new();
    let mut table_ids = std::collections::BTreeSet::new();
    loop {
        let (ns, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
        let empty = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let name = std::str::from_utf8(element.local_name().as_ref())
                    .map_err(err)?
                    .to_owned();
                if ns == ResolveResult::Bound(quick_xml::name::Namespace(crate::xml::NS)) {
                    if name == "mergeCell" && stack.as_slice() == ["worksheet", "mergeCells"] {
                        rectangles.push(Geometry::from_reference(
                            crate::xml::attrs(&element)?.get("ref").ok_or_else(|| {
                                crate::coded("spill-bad-input", "Source merge has no rectangle")
                            })?,
                        )?);
                    }
                    if name == "tablePart" && stack.as_slice() == ["worksheet", "tableParts"] {
                        let mut id = None;
                        for attribute in element.attributes() {
                            let attribute = attribute.map_err(err)?;
                            let (binding, local) = reader.resolve_attribute(attribute.key);
                            if local.as_ref()==b"id"&&binding==ResolveResult::Bound(quick_xml::name::Namespace(b"http://schemas.openxmlformats.org/officeDocument/2006/relationships")){if id.is_some(){return Err(crate::coded("spill-ambiguous", "Ambiguous table relationship identity"));}id=Some(attribute.unescape_value().map_err(err)?.into_owned());}
                        }
                        let id = id.ok_or_else(|| {
                            crate::coded(
                                "spill-unsupported",
                                "Source table has no qualified relationship identity",
                            )
                        })?;
                        if id.is_empty() || !table_ids.insert(id) {
                            return Err(crate::coded(
                                "spill-ambiguous",
                                "Ambiguous source table ownership",
                            ));
                        }
                    }
                }
                if !empty {
                    stack.push(name);
                }
            }
            Event::End(_) => {
                stack.pop().ok_or_else(|| {
                    crate::coded("spill-malformed", "Unbalanced source worksheet metadata")
                })?;
            }
            Event::DocType(_) => {
                return Err(crate::coded(
                    "spill-malformed",
                    "XML entities are unsupported",
                ));
            }
            Event::Eof => break,
            _ => (),
        }
        buffer.clear();
    }
    if table_ids.is_empty() {
        return Ok(rectangles);
    }
    let (directory, file) = sheet
        .rsplit_once('/')
        .ok_or_else(|| crate::coded("spill-malformed", "Invalid source worksheet path"))?;
    let rels = format!("{directory}/_rels/{file}.rels");
    let mut xml = String::new();
    zip.by_name(&rels)
        .map_err(err)?
        .read_to_string(&mut xml)
        .map_err(err)?;
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut buffer = Vec::new();
    let mut relationships = std::collections::BTreeMap::new();
    loop {
        let (ns, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
        match event {
            Event::Start(element) | Event::Empty(element)
                if ns
                    == ResolveResult::Bound(quick_xml::name::Namespace(
                        b"http://schemas.openxmlformats.org/package/2006/relationships",
                    ))
                    && element.local_name().as_ref() == b"Relationship" =>
            {
                let attributes = crate::xml::attrs(&element)?;
                let id = attributes
                    .get("Id")
                    .ok_or_else(|| {
                        crate::coded(
                            "internal-missing-field",
                            "Missing source relationship identity",
                        )
                    })?
                    .clone();
                if id.is_empty() || relationships.insert(id, attributes).is_some() {
                    return Err(crate::coded(
                        "spill-ambiguous",
                        "Ambiguous source worksheet relationship identities",
                    ));
                }
            }
            Event::DocType(_) => {
                return Err(crate::coded(
                    "spill-malformed",
                    "XML entities are unsupported",
                ));
            }
            Event::Eof => break,
            _ => (),
        }
        buffer.clear();
    }
    for id in table_ids {
        let relationship = relationships.get(&id).ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Source table relationship is missing",
            )
        })?;
        if relationship.get("Type").map(String::as_str)
            != Some("http://schemas.openxmlformats.org/officeDocument/2006/relationships/table")
            || relationship
                .get("TargetMode")
                .is_some_and(|mode| mode != "Internal")
        {
            return Err(crate::coded(
                "spill-unsupported",
                "Source table relationship is unqualified",
            ));
        }
        let target = relationship.get("Target").ok_or_else(|| {
            crate::coded("internal-missing-field", "Source table target is missing")
        })?;
        let mut components = if target.starts_with('/') {
            Vec::new()
        } else {
            directory.split('/').collect::<Vec<_>>()
        };
        for component in target.trim_start_matches('/').split('/') {
            match component {
                ".." => {
                    components.pop().ok_or_else(|| {
                        crate::coded("spill-unsupported", "Source table target escapes package")
                    })?;
                }
                "." | "" => {
                    return Err(crate::coded(
                        "spill-unsupported",
                        "Source table target is noncanonical",
                    ));
                }
                _ => components.push(component),
            }
        }
        let part = components.join("/");
        if !part.starts_with("xl/tables/")
            || !part.ends_with(".xml")
            || part.matches('/').count() != 2
        {
            return Err(crate::coded(
                "spill-unsupported",
                "Source table target is outside the qualified private part layout",
            ));
        }
        let mut xml = String::new();
        {
            let mut file = zip.by_name(&part).map_err(err)?;
            if file.size() > 8 * 1024 * 1024 {
                return Err(crate::coded(
                    "spill-limit",
                    "Source table exceeds spill proof bound",
                ));
            }
            file.read_to_string(&mut xml).map_err(err)?;
        }
        let mut reader = NsReader::from_reader(xml.as_bytes());
        let mut buffer = Vec::new();
        loop {
            let (ns, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
            match event {
                Event::Start(element) | Event::Empty(element) => {
                    if ns != ResolveResult::Bound(quick_xml::name::Namespace(crate::xml::NS))
                        || element.local_name().as_ref() != b"table"
                    {
                        return Err(crate::coded(
                            "spill-unsupported",
                            "Source table root is unqualified",
                        ));
                    }
                    rectangles.push(Geometry::from_reference(
                        crate::xml::attrs(&element)?.get("ref").ok_or_else(|| {
                            crate::coded("spill-bad-input", "Source table has no rectangle")
                        })?,
                    )?);
                    break;
                }
                Event::DocType(_) => {
                    return Err(crate::coded(
                        "spill-malformed",
                        "XML entities are unsupported",
                    ));
                }
                Event::Eof => return Err(crate::coded("spill-bad-input", "Empty source table")),
                _ => (),
            }
            buffer.clear();
        }
    }
    Ok(rectangles)
}
// Accepted geometry overlays are native version data. A caller cannot omit
// new merges from its view or keep stale source barriers after unmerge.
fn version_barriers(
    root: &Path,
    db: &Connection,
    version: u64,
    sheet: &str,
) -> Result<Vec<Geometry>> {
    if let Some(ranges) = crate::merges::selected_override(db, version, sheet)? {
        ranges
            .iter()
            .map(|range| Geometry::from_reference(range))
            .collect()
    } else {
        source_barriers(root, sheet)
    }
}
pub(crate) fn source_group(
    root: &Path,
    db: &Connection,
    sheet: &str,
    reference: &str,
    cell: &Value,
    eligible: &std::collections::HashSet<u32>,
) -> Result<Option<Geometry>> {
    if cell.get("formulaType").is_none() {
        return Ok(None);
    }
    if cell["formulaType"] != "array" {
        return Err(crate::coded(
            "spill-unsupported",
            "Legacy/shared/data-table source is outside dynamic spill adoption",
        ));
    }
    let master = position(reference)?;
    let raw_attributes: String = db
        .prepare_cached("SELECT formula_attrs FROM cells WHERE sheet=?1 AND r=?2 AND c=?3")
        .map_err(crate::err)?
        .query_row(params![sheet, master.0, master.1], |row| row.get(0))
        .map_err(err)?;
    let formula_attributes: Value = serde_json::from_str(&raw_attributes).map_err(err)?;
    if formula_attributes["t"] != "array"
        || formula_attributes.as_object().is_none_or(|attributes| {
            attributes
                .keys()
                .any(|key| !["t", "ref", "aca", "ca", "xmlns"].contains(&key.as_str()))
        })
        || formula_attributes.get("xmlns").is_some_and(|value| {
            value.as_str() != Some(std::str::from_utf8(crate::xml::NS).unwrap())
        })
        || ["aca", "ca"].iter().any(|key| {
            formula_attributes
                .get(*key)
                .is_some_and(|value| !["0", "1"].contains(&value.as_str().unwrap_or("")))
        })
    {
        return Err(crate::coded(
            "spill-unsupported",
            "Extended source array attributes require preservation qualification",
        ));
    }
    let shape = Geometry::from_reference(formula_attributes["ref"].as_str().ok_or_else(|| {
        crate::coded("internal-missing-field", "Missing source array rectangle")
    })?)?;
    if shape.first != master || shape.count() > u64::from(MAX_OUTPUTS) {
        return Err(crate::coded(
            "spill-bad-input",
            "Dynamic source master is not at its bounded output origin",
        ));
    }
    let attributes = source_attributes(root, db, sheet, master)?.ok_or_else(|| {
        crate::coded(
            "spill-bad-input",
            "Dynamic source master has no source attributes",
        )
    })?;
    if attributes
        .keys()
        .any(|key| !["r", "s", "t", "cm", "vm", "xmlns"].contains(&key.as_str()))
        || attributes
            .get("xmlns")
            .is_some_and(|value| value.as_bytes() != crate::xml::NS)
    {
        return Err(crate::coded(
            "spill-unsupported",
            "Opaque source master attributes require preservation qualification",
        ));
    }
    let cm = attributes.get("cm").ok_or_else(|| {
        crate::coded(
            "spill-bad-input",
            "Dynamic source master has no cell metadata index",
        )
    })?;
    if cm.is_empty() || cm.starts_with('0') || !cm.bytes().all(|byte| byte.is_ascii_digit()) {
        return Err(crate::coded(
            "spill-malformed",
            "Invalid canonical source cell metadata index",
        ));
    }
    let cm = cm.parse::<u32>().map_err(err)?;
    if !eligible.contains(&cm) {
        return Err(crate::coded(
            "spill-refused",
            "Source array lacks cm-to-XLDAPR dynamic ownership proof",
        ));
    }
    // The current formula must still identify the actual imported master. A
    // foreign array record cannot borrow another source cell's metadata proof.
    let raw: String = db
        .prepare_cached("SELECT formula FROM cells WHERE sheet=?1 AND r=?2 AND c=?3")
        .map_err(crate::err)?
        .query_row(params![sheet, master.0, master.1], |row| row.get(0))
        .map_err(err)?;
    if cell["formula"].as_str() != Some(raw.as_str()) {
        return Err(crate::coded(
            "spill-stale",
            "Dynamic source adoption has changed formula identity",
        ));
    }
    source_content_proof(root, db, sheet, master, true)?;
    Ok(Some(shape))
}

fn scan(
    root: &Path,
    db: &Connection,
    input: &Value,
    limit: Option<usize>,
    mut accept: impl FnMut(&str, &str, Value) -> Result<()>,
) -> Result<usize> {
    let version = input["version"].as_u64().ok_or_else(|| {
        crate::coded("internal-missing-field", "Missing spill inspection version")
    })?;
    crate::calculation::snapshot(db, version)?;
    crate::scenarios::prepare(db, &input["view"])?;
    db.execute_batch("CREATE TEMP TABLE spill_protected(sheet TEXT,mr INTEGER,mc INTEGER,r0 INTEGER,c0 INTEGER,r1 INTEGER,c1 INTEGER,PRIMARY KEY(sheet,mr,mc)) WITHOUT ROWID;").map_err(err)?;
    for sheet in input["view"]["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheets"))?
    {
        let path = crate::field(sheet, "path")?;
        let mut after = None;
        loop {
            let page = crate::calculation::page(&db, path, (0, 0), (1048575, 16383), after, true)?;
            for (reference, cell) in &page {
                if cell.get("formulaType").is_none() {
                    continue;
                }
                let shape = Geometry::from_reference(
                    cell["formulaReference"].as_str().ok_or_else(|| {
                        crate::coded(
                            "spill-protected",
                            "Special formula has no protected rectangle",
                        )
                    })?,
                )?;
                let at = position(reference)?;
                db.execute(
                    "INSERT INTO spill_protected VALUES(?1,?2,?3,?4,?5,?6,?7)",
                    params![
                        path,
                        at.0,
                        at.1,
                        shape.first.0,
                        shape.first.1,
                        shape.last.0,
                        shape.last.1
                    ],
                )
                .map_err(err)?;
            }
            if page.len() < 512 {
                break;
            }
            after = Some(position(&page.last().unwrap().0)?);
        }
    }
    let eligible = metadata_proof(root)?;
    let source_metadata = metadata::source(root)?;
    let view = &input["view"];
    let mut engine = crate::formula::Eval::from_sqlite(view, db, false);
    engine.fence_spill_inputs();
    let index: Value =
        serde_json::from_slice(&std::fs::read(root.join("index.json")).map_err(err)?)
            .map_err(err)?;
    let mut count = 0;
    for sheet in view["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheets"))?
    {
        let (name, path) = (crate::field(sheet, "name")?, crate::field(sheet, "path")?);
        let mut after = None;
        let owners = effective(&db, version, path)?;
        let source_barriers = version_barriers(root, db, version, path)?;
        loop {
            let page = crate::calculation::page(&db, path, (0, 0), (1048575, 16383), after, true)?;
            for (reference, cell) in &page {
                let Some(formula) = cell["formula"].as_str() else {
                    continue;
                };
                let descriptor =
                    crate::formula::sequence_dependencies(view, name, reference, formula);
                if matches!(&descriptor, Ok(None)) {
                    continue;
                }
                if limit.is_some_and(|limit| count >= limit) {
                    return Err(crate::coded(
                        "spill-limit",
                        "Spill inspection response exceeds its constructor bound",
                    ));
                }
                let result = (|| -> Result<Value> {
                    if index["sheets"]
                        .as_array()
                        .ok_or_else(|| {
                            crate::coded("internal-missing-field", "Missing native sheet metadata")
                        })?
                        .iter()
                        .find(|item| item["path"] == path)
                        .is_none_or(|item| item["extended"] == true)
                    {
                        return Err(crate::coded(
                            "spill-unsupported",
                            "Opaque worksheet extensions require spill preservation qualification",
                        ));
                    }
                    let descriptor = descriptor?.ok_or_else(|| {
                        crate::coded("internal-missing-field", "Missing SEQUENCE descriptor")
                    })?;
                    let at = position(reference)?;
                    let owner = owners.iter().find(|owner| owner.master == at);
                    if owner.is_some_and(|owner| owner.formula != formula) {
                        return Err(crate::coded(
                            "spill-stale",
                            "Spill formula identity changed without ownership release",
                        ));
                    }
                    let source = if let Some(owner) = owner {
                        validate_output_cells(db, version, owner)?;
                        None
                    } else {
                        source_group(root, &db, path, reference, cell, &eligible)?
                    };
                    if source.is_none() && owner.is_none() {
                        let released = released_master(db, version, path, at)?;
                        let attributes = source_attributes(root, &db, path, at)?;
                        if attributes.as_ref().is_some_and(|attributes| {
                            attributes.keys().any(|key| {
                                !if released {
                                    ["r", "s", "t", "cm", "vm", "xmlns"].contains(&key.as_str())
                                } else {
                                    ["r", "s", "t", "xmlns"].contains(&key.as_str())
                                }
                            }) || attributes
                                .get("xmlns")
                                .is_some_and(|value| value.as_bytes() != crate::xml::NS)
                        }) {
                            return Err(crate::coded(
                                "spill-unsupported",
                                "Opaque authored master attributes require preservation qualification",
                            ));
                        }
                        if released
                            && let Some(cm) = attributes
                                .as_ref()
                                .and_then(|attributes| attributes.get("cm"))
                        {
                            let index = cm.parse::<u32>().map_err(err)?;
                            if cm.starts_with('0') || !eligible.contains(&index) {
                                return Err(crate::coded(
                                    "spill-unsupported",
                                    "Released source master has unqualified dynamic metadata",
                                ));
                            }
                        }
                        source_content_proof(root, &db, path, at, true)?;
                    }
                    let old = owner.map(|owner| owner.geometry.clone()).or(source.clone());
                    let args = engine
                        .sequence_arguments(name, reference, formula)?
                        .ok_or_else(|| {
                            crate::coded("internal-missing-field", "Missing SEQUENCE arguments")
                        })?;
                    let sequence = Sequence::new(position(reference)?, &args)?;
                    if source_barriers
                        .iter()
                        .any(|barrier| barrier.intersects(&sequence.geometry))
                    {
                        return Err(crate::coded(
                            "spill-conflict",
                            "Source merge/table intersection is outside the qualified spill slice",
                        ));
                    }
                    if let Some(source_shape) = source.as_ref() {
                        let attributes =
                            source_attributes(root, db, path, at)?.ok_or_else(|| {
                                crate::coded(
                                    "internal-missing-field",
                                    "Missing source master attributes",
                                )
                            })?;
                        let cm = attributes
                            .get("cm")
                            .ok_or_else(|| {
                                crate::coded("internal-missing-field", "Missing source cm index")
                            })?
                            .parse::<u64>()
                            .map_err(err)?;
                        let mut master = json!({"cm":cm,"range":source_shape.reference()});
                        if let Some(vm) = attributes.get("vm") {
                            if source_shape.count() != 1
                                || vm.is_empty()
                                || vm.starts_with('0')
                                || !vm.bytes().all(|byte| byte.is_ascii_digit())
                            {
                                return Err(crate::coded(
                                    "spill-unsupported",
                                    "Extended/noncanonical source value metadata requires preservation qualification",
                                ));
                            }
                            master["vm"] = json!(vm.parse::<u64>().map_err(err)?);
                            master["intendedRange"] = json!(sequence.geometry.reference());
                            master["aca"] = json!("1");
                            master["ca"] = json!("1");
                        }
                        let mut mapping = serde_json::Map::new();
                        mapping.insert(format!("{path}\0{reference}"), master);
                        metadata::validate(
                            &source_metadata,
                            &Default::default(),
                            &Value::Object(mapping),
                        )?;
                    }
                    let protected:bool=db.prepare_cached("SELECT EXISTS(SELECT 1 FROM spill_protected WHERE sheet=?1 AND r0<=?2 AND r1>=?3 AND c0<=?4 AND c1>=?5 AND NOT(mr=?6 AND mc=?7))").map_err(crate::err)?.query_row(params![path,sequence.geometry.last.0,sequence.geometry.first.0,sequence.geometry.last.1,sequence.geometry.first.1,at.0,at.1],|row|row.get(0)).map_err(err)?;
                    if protected
                        || owners.iter().any(|owner| {
                            owner.master != at && owner.geometry.intersects(&sequence.geometry)
                        })
                    {
                        return Err(crate::coded(
                            "spill-conflict",
                            "Another array/special/spill owner intersects the intended output",
                        ));
                    }
                    for range in descriptor["ranges"].as_array().ok_or_else(|| {
                        crate::coded(
                            "internal-missing-field",
                            "Missing SEQUENCE input dependencies",
                        )
                    })? {
                        if range["sheet"] == name
                            && sequence
                                .geometry
                                .intersects(&Geometry::from_reference(crate::field(range, "ref")?)?)
                        {
                            return Err(crate::coded(
                                "spill-refused",
                                "SEQUENCE constructor reads its own intended output",
                            ));
                        }
                    }
                    if sheet["merges"].as_array().is_some_and(|merges| {
                        merges.iter().any(|merge| {
                            merge
                                .as_str()
                                .and_then(|r| Geometry::from_reference(r).ok())
                                .is_some_and(|merge| merge.intersects(&sequence.geometry))
                        })
                    }) {
                        return Err(crate::coded(
                            "spill-conflict",
                            "Merged intersection is outside the qualified spill slice",
                        ));
                    }
                    if view["tables"].as_array().is_some_and(|tables| {
                        tables.iter().any(|table| {
                            table["sheet"] == name
                                && table["ref"]
                                    .as_str()
                                    .and_then(|r| Geometry::from_reference(r).ok())
                                    .is_some_and(|table| table.intersects(&sequence.geometry))
                        })
                    }) {
                        return Err(crate::coded(
                            "spill-conflict",
                            "Table intersection is outside the qualified spill slice",
                        ));
                    }
                    let mut blocked = None;
                    for at in sequence.geometry.positions() {
                        if at == sequence.geometry.first {
                            continue;
                        }
                        let destination =
                            crate::calculation::cell(&db, path, &address(at.0, at.1))?;
                        if destination["formula"].is_string()
                            || destination.get("formulaType").is_some()
                        {
                            return Err(crate::coded(
                                "spill-conflict",
                                "Special/formula intersection is outside the qualified spill slice",
                            ));
                        }
                        let attrs = source_attributes(root, &db, path, at)?;
                        if attrs.as_ref().is_some_and(|attrs| {
                            attrs
                                .keys()
                                .any(|key| !["r", "s", "t", "xmlns"].contains(&key.as_str()))
                                || attrs
                                    .get("xmlns")
                                    .is_some_and(|value| value.as_bytes() != crate::xml::NS)
                        }) {
                            return Err(crate::coded(
                                "spill-unsupported",
                                "Opaque destination attributes require spill preservation qualification",
                            ));
                        }
                        source_content_proof(root, &db, path, at, false)?;
                        if old.as_ref().is_some_and(|geometry| geometry.contains(at)) {
                            continue;
                        }
                        if !destination.is_null() && !destination["value"].is_null() {
                            blocked = Some(address(at.0, at.1));
                            break;
                        }
                    }
                    Ok(
                        json!({"ref":reference,"formula":formula,"range":sequence.geometry.reference(),"sourceRange":old.map(|geometry|geometry.reference()),"arguments":sequence.arguments,"values":sequence.values,"dependencies":descriptor["ranges"],"status":if blocked.is_some(){"blocked"}else{"qualified-constructor"},"blocker":blocked}),
                    )
                })();
                accept(
                    path,
                    name,
                    match result {
                        Ok(item) => item,
                        Err(reason) => {
                            json!({"ref":reference,"formula":formula,"status":"refused","reason":reason})
                        }
                    },
                )?;
                count += 1;
            }
            if page.len() < 512 {
                break;
            }
            after = Some(position(&page.last().unwrap().0)?);
        }
    }
    Ok(count)
}

pub fn inspect(root: &Path, input: &Value) -> Result<Value> {
    let db = crate::database(root)?;
    let mut items = Vec::new();
    scan(root, &db, input, Some(128), |_, _, item| {
        items.push(item);
        Ok(())
    })?;
    Ok(json!({"enabled":false,"version":input["version"],"schema":SCHEMA,"items":items}))
}

/// Recompute constructor scalars against the export version instead of trusting
/// a prior version's result fields. Any missing/stale owner refuses the export.
pub(crate) fn validate_current(root: &Path, db: &Connection, input: &Value) -> Result<()> {
    let version = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing spill proof version"))?;
    let mut owners = std::collections::BTreeMap::new();
    for sheet in input["view"]["sheets"].as_array().ok_or_else(|| {
        crate::coded(
            "spill-bad-input",
            "Spill proof requires current worksheet metadata",
        )
    })? {
        let path = crate::field(sheet, "path")?;
        for group in effective(db, version, path)? {
            validate_output_cells(db, version, &group)?;
            owners.insert(
                (path.to_owned(), address(group.master.0, group.master.1)),
                group,
            );
        }
    }
    let mut matched = std::collections::BTreeSet::new();
    let mut stored_validated = std::collections::BTreeSet::new();
    scan(root, db, input, None, |sheet, _, item| {
        let key = (sheet.to_owned(), crate::field(&item, "ref")?.to_owned());
        let Some(group) = owners.get(&key) else {
            return Ok(());
        };
        let sequence = validate_constructor(group)?;
        if crate::calculation_policy::pending(db, version, sheet, group.master)?
            && crate::calculation_transfer::publication_spill_certificate(root, db, version, group)?
        {
            matched.insert(key);
            return Ok(());
        }
        if group.proof["cacheAuthority"] == "source"
            && crate::calculation_policy::pending(db, version, sheet, group.master)?
        {
            validate_source_cache(root, db, group, &input["view"])?;
            matched.insert(key);
            return Ok(());
        }
        if group.version < version
            && crate::calculation_policy::pending(db, version, sheet, group.master)?
        {
            // Manual retains this complete accepted cache at the version that
            // produced it. Revalidate that actual ancestor, never current
            // constructor inputs or a caller-supplied pending flag alone.
            if stored_validated.insert(group.version) {
                let ancestor = crate::database(root)?;
                validate_current(
                    root,
                    &ancestor,
                    &json!({"version":group.version,"view":input["view"]}),
                )?;
            }
            matched.insert(key);
            return Ok(());
        }
        let args = item["arguments"]
            .as_array()
            .ok_or_else(|| {
                crate::coded(
                    "spill-stale",
                    "Accepted spill constructor is no longer qualified",
                )
            })?
            .iter()
            .map(|value| {
                value.as_f64().ok_or_else(|| {
                    crate::coded(
                        "internal-invalid-field",
                        "Invalid current spill scalar proof",
                    )
                })
            })
            .collect::<Result<Vec<_>>>()?;
        let current = Sequence::new(group.master, &args)?;
        let blocked = group.proof["status"] == "blocked";
        if current.arguments != sequence.arguments
            || current.geometry != sequence.geometry
            || item["status"]
                != if blocked {
                    "blocked"
                } else {
                    "qualified-constructor"
                }
            || blocked && item["blocker"] != group.proof["blocker"]
        {
            return Err(crate::coded(
                "spill-stale",
                "Accepted spill constructor/output proof is stale at this version",
            ));
        }
        matched.insert(key);
        Ok(())
    })?;
    if matched.len() != owners.len() {
        return Err(crate::coded(
            "spill-refused",
            "Accepted spill ownership has a missing current constructor",
        ));
    }
    Ok(())
}

/// The unqualified staging primitive, for the tests only: command dispatch
/// uses `stage_qualified`. The diagnostic's reply cap is not a production
/// constructor cap: candidates stream into transaction-local SQL.
#[cfg(test)]
fn stage(root: &Path, input: &Value) -> Result<Value> {
    stage_mode(root, input, false)
}
fn retained_arguments(arguments: &Value) -> bool {
    let Some(arguments) = arguments.as_array() else {
        return false;
    };
    let args = arguments
        .iter()
        .map(Value::as_f64)
        .collect::<Option<Vec<_>>>();
    args.is_some_and(|args| {
        [
            [3.0, 2.0, 1.0, 1.0],
            [2.0, 2.0, 1.0, 1.0],
            [2.0, 2.0, 10.0, 2.0],
        ]
        .iter()
        .any(|qualified| args.as_slice() == qualified)
    })
}
pub(crate) fn stage_qualified(root: &Path, input: &Value) -> Result<Value> {
    stage_mode(root, input, true)
}
fn exact_source_values(
    root: &Path,
    db: &Connection,
    sheet: &str,
    sequence: &Sequence,
) -> Result<bool> {
    let page = crate::range::read_db(
        db,
        &json!({"version":0,"sheet":sheet,"range":sequence.geometry.reference(),"limit":MAX_OUTPUTS}),
    )?;
    let cells = page["cells"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing source cache cells"))?;
    if cells.len() as u64 != sequence.geometry.count() {
        return Ok(false);
    }
    for (pair, expected) in cells.iter().zip(&sequence.values) {
        let at =
            position(pair[0].as_str().ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing source address")
            })?)?;
        let cell = &pair[1];
        if at != sequence.geometry.first
            && (cell.get("formula").is_some() || cell.get("formulaType").is_some())
        {
            return Ok(false);
        }
        let value = if at == sequence.geometry.first {
            &cell["cachedValue"]
        } else {
            &cell["value"]
        };
        let number = value.as_f64();
        if cell["dataType"] != "number"
            || !number
                .zip(expected.as_f64())
                .is_some_and(|(a, b)| a.to_bits() == b.to_bits())
            || cell["rawValue"].as_str().is_some_and(|raw| {
                number.is_none_or(|value| {
                    crate::bulk::meaning(raw) != crate::bulk::meaning(&value.to_string())
                })
            })
        {
            return Ok(false);
        }
        if at != sequence.geometry.first {
            let attrs = source_attributes(root, db, sheet, at)?;
            if attrs.as_ref().is_some_and(|attrs| {
                attrs
                    .keys()
                    .any(|key| !["r", "s", "t", "xmlns"].contains(&key.as_str()))
            }) {
                return Ok(false);
            }
            source_content_proof(root, db, sheet, at, false)?;
        }
    }
    Ok(true)
}
fn validate_source_cache(root: &Path, db: &Connection, group: &Group, view: &Value) -> Result<()> {
    let original = crate::range::read_db(
        db,
        &json!({"version":0,"sheet":group.sheet,"range":address(group.master.0,group.master.1),"limit":1}),
    )?;
    let source = original["cells"]
        .as_array()
        .and_then(|cells| cells.first())
        .map(|pair| &pair[1])
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing source cache master"))?;
    if source["formula"] != group.formula
        || source_group(
            root,
            db,
            &group.sheet,
            &address(group.master.0, group.master.1),
            source,
            &metadata_proof(root)?,
        )? != Some(group.geometry.clone())
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Stored ownership differs from immutable source metadata",
        ));
    }
    let sequence = validate_constructor(group)?;
    if !exact_source_values(root, db, &group.sheet, &sequence)?
        || source_barriers(root, &group.sheet)?
            .iter()
            .any(|barrier| barrier.intersects(&group.geometry))
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Stored ownership differs from complete immutable source values/geometry",
        ));
    }
    let name = view["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing source sheets"))?
        .iter()
        .find(|sheet| sheet["path"] == group.sheet)
        .and_then(|sheet| sheet["name"].as_str())
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing source cache worksheet"))?;
    let mut engine = crate::formula::Eval::from_sqlite(view, db, false);
    if !engine.sequence_cache_profile(
        name,
        &address(group.master.0, group.master.1),
        &group.formula,
        &sequence.arguments,
    )? {
        return Err(crate::coded(
            "internal-mismatch",
            "Stored cache changes a literal constructor argument",
        ));
    }
    Ok(())
}
fn adopt_source_cache(root: &Path, input: &Value) -> Result<Value> {
    let mut db = crate::database(root)?;
    let parent = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing stored version"))?;
    if parent != 0
        || crate::calculation_policy::options(&db, parent)?.is_none_or(|(mode, _)| mode != "manual")
    {
        return Err(crate::coded(
            "spill-bad-input",
            "Stored spill adoption requires the initialized Manual base",
        ));
    }
    let tx = db.transaction().map_err(err)?;
    crate::calculation::snapshot(&tx, parent)?;
    let eligible = metadata_proof(root)?;
    let mut groups = Vec::new();
    let mut refused = 0;
    let mut engine = crate::formula::Eval::from_sqlite(&input["view"], &tx, false);
    for sheet in input["view"]["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing stored sheets"))?
    {
        let path = crate::field(sheet, "path")?;
        let name = crate::field(sheet, "name")?;
        let mut after = None;
        loop {
            let page = crate::calculation::page(&tx, path, (0, 0), (1048575, 16383), after, true)?;
            for (reference, cell) in &page {
                let Some(formula) = cell["formula"].as_str() else {
                    continue;
                };
                let descriptor = match crate::formula::sequence_dependencies(
                    &input["view"],
                    name,
                    reference,
                    formula,
                ) {
                    Ok(Some(descriptor)) => descriptor,
                    _ => continue,
                };
                let master = position(reference)?;
                if !crate::calculation_policy::pending(&tx, parent, path, master)? {
                    continue;
                }
                let geometry = match source_group(root, &tx, path, reference, cell, &eligible) {
                    Ok(Some(geometry)) => geometry,
                    Ok(None) => continue,
                    Err(_) => {
                        crate::cancel::check()?;
                        refused += 1;
                        continue;
                    }
                };
                let mut accepted = None;
                for arguments in [
                    [3.0, 2.0, 1.0, 1.0],
                    [2.0, 2.0, 1.0, 1.0],
                    [2.0, 2.0, 10.0, 2.0],
                ] {
                    let sequence = Sequence::new(master, &arguments)?;
                    if sequence.geometry != geometry
                        || !engine
                            .sequence_cache_profile(name, reference, formula, &arguments)
                            .unwrap_or(false)
                        || !exact_source_values(root, &tx, path, &sequence)?
                        || version_barriers(root, &tx, parent, path)?
                            .iter()
                            .any(|barrier| barrier.intersects(&geometry))
                    {
                        continue;
                    }
                    if descriptor["ranges"]
                        .as_array()
                        .ok_or_else(|| {
                            crate::coded("internal-missing-field", "Missing stored dependencies")
                        })?
                        .iter()
                        .any(|range| {
                            range["sheet"] == name
                                && Geometry::from_reference(range["ref"].as_str().unwrap_or(""))
                                    .is_ok_and(|shape| shape.intersects(&geometry))
                        })
                    {
                        continue;
                    }
                    accepted = Some(Group {
                        version: 0,
                        sheet: path.into(),
                        master,
                        geometry: geometry.clone(),
                        formula: formula.into(),
                        proof: json!({"constructor":"sequence-integer-v1","arguments":arguments,"dependencies":descriptor["ranges"],"status":"calculated","qualification":"retained-sequence-v1","cacheAuthority":"source"}),
                    });
                    break;
                }
                if let Some(group) = accepted {
                    groups.push(group);
                } else {
                    refused += 1;
                }
            }
            if page.len() < 512 {
                break;
            }
            after = Some(position(&page.last().unwrap().0)?);
        }
    }
    if groups.len() > 128
        || groups.iter().enumerate().any(|(index, a)| {
            groups
                .iter()
                .skip(index + 1)
                .any(|b| a.sheet == b.sheet && a.geometry.intersects(&b.geometry))
        })
    {
        return Ok(
            json!({"version":parent,"groups":0,"refusedConstructors":groups.len()+refused,"storedOnly":true}),
        );
    }
    if groups.is_empty() {
        return Ok(
            json!({"version":parent,"groups":0,"refusedConstructors":refused,"storedOnly":true}),
        );
    }
    drop(engine);
    schema(&tx)?;
    tx.execute("INSERT INTO versions(parent)VALUES(?1)", [parent])
        .map_err(err)?;
    let version = tx.last_insert_rowid() as u64;
    let mut positions = 0;
    for group in &mut groups {
        group.version = version;
        publish(
            &tx,
            group,
            Some(&group.geometry),
            crate::field(input, "revision")?,
        )?;
        positions += group.geometry.count();
    }
    tx.commit().map_err(err)?;
    Ok(
        json!({"version":version,"groups":groups.len(),"outputPositions":positions,"storedOnly":true}),
    )
}
fn stage_mode(root: &Path, input: &Value, qualified: bool) -> Result<Value> {
    if input["storedOnly"] == true {
        return adopt_source_cache(root, input);
    }
    let mut db = crate::database(root)?;
    let tx = db.transaction().map_err(err)?;
    let parent = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing spill staging version"))?;
    let revision = crate::field(input, "revision")?;
    known_version(&tx, parent)?;
    // Ordinary scalar workbooks do not pay for a second formula projection.
    // False positives only enter the closed parser; this hint never qualifies
    // syntax, current ownership, or values on its own.
    let candidate:bool=tx.prepare_cached("SELECT EXISTS(SELECT 1 FROM cells WHERE formula IS NOT NULL AND instr(upper(formula),'SEQUENCE')>0) OR EXISTS(SELECT 1 FROM changes WHERE instr(upper(data),'SEQUENCE')>0)").map_err(crate::err)?.query_row([],|row|row.get(0)).map_err(err)?;
    if !candidate && !any_active(&tx, parent)? {
        return Ok(json!({"version":parent,"schema":SCHEMA,"groups":0,"enabled":false}));
    }
    tx.execute_batch("CREATE TEMP TABLE spill_staged(sheet TEXT,mr INTEGER,mc INTEGER,r0 INTEGER,c0 INTEGER,r1 INTEGER,c1 INTEGER,old_range TEXT,formula TEXT,arguments TEXT,dependencies TEXT,blocked INTEGER,blocker TEXT,PRIMARY KEY(sheet,mr,mc)) WITHOUT ROWID;").map_err(err)?;
    let mut refused = 0;
    scan(root, &tx, input, None, |sheet, _, item| {
        if !["qualified-constructor", "blocked"].contains(&item["status"].as_str().unwrap_or(""))
            || qualified && !retained_arguments(&item["arguments"])
        {
            refused += 1;
            return Ok(());
        }
        let at = position(crate::field(&item, "ref")?)?;
        let geometry = Geometry::from_reference(crate::field(&item, "range")?)?;
        if input["storedOnly"] == true {
            if item["status"] != "qualified-constructor" || item["sourceRange"] != item["range"] {
                refused += 1;
                return Ok(());
            }
            let expected = item["values"].as_array().ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing retained source values")
            })?;
            let mut exact = true;
            for (position, value) in geometry.positions().zip(expected) {
                let cell = crate::calculation::cell(&tx, sheet, &address(position.0, position.1))?;
                let cache = if cell.get("spill").is_some() {
                    &cell["calculation"]["value"]
                } else if position == at {
                    &cell["cachedValue"]
                } else {
                    &cell["value"]
                };
                let number = cache.as_f64();
                exact &= cell["dataType"] == "number"
                    && number
                        .zip(value.as_f64())
                        .is_some_and(|(a, b)| a.to_bits() == b.to_bits())
                    && cell["rawValue"].as_str().is_none_or(|raw| {
                        number.is_some_and(|value| {
                            crate::bulk::meaning(raw) == crate::bulk::meaning(&value.to_string())
                        })
                    });
            }
            if !exact {
                refused += 1;
                return Ok(());
            }
        }
        tx.execute(
            "INSERT INTO spill_staged VALUES(?1,?2,?3,?4,?5,?6,?7,?8,?9,?10,?11,?12,?13)",
            params![
                sheet,
                at.0,
                at.1,
                geometry.first.0,
                geometry.first.1,
                geometry.last.0,
                geometry.last.1,
                item["sourceRange"].as_str(),
                crate::field(&item, "formula")?,
                item["arguments"].to_string(),
                item["dependencies"].to_string(),
                u32::from(item["status"] == "blocked"),
                item["blocker"].as_str()
            ],
        )
        .map_err(err)?;
        Ok(())
    })?;
    // A previously accepted owner may never silently keep stale followers when
    // its constructor becomes unsupported, changes identity, or disappears.
    for sheet in input["view"]["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheets"))?
    {
        let path = crate::field(sheet, "path")?;
        for owner in effective(&tx, parent, path)? {
            if !tx
                .prepare_cached(
                    "SELECT EXISTS(SELECT 1 FROM spill_staged WHERE sheet=?1 AND mr=?2 AND mc=?3)",
                )
                .map_err(crate::err)?
                .query_row(params![path, owner.master.0, owner.master.1], |row| {
                    row.get::<_, bool>(0)
                })
                .map_err(err)?
            {
                return Err(crate::coded(
                    "spill-unsupported",
                    "Accepted spill owner became unqualified without an ownership release",
                ));
            }
        }
    }
    let count: u64 = tx
        .prepare_cached("SELECT count(*) FROM spill_staged")
        .map_err(crate::err)?
        .query_row([], |row| row.get(0))
        .map_err(err)?;
    if count == 0 {
        return Ok(
            json!({"version":parent,"schema":SCHEMA,"groups":0,"refusedConstructors":refused,"enabled":false}),
        );
    }
    let overlap:bool=tx.prepare_cached("SELECT EXISTS(SELECT 1 FROM spill_staged a JOIN spill_staged b ON a.sheet=b.sheet AND (a.mr,a.mc)<(b.mr,b.mc) AND a.r0<=b.r1 AND b.r0<=a.r1 AND a.c0<=b.c1 AND b.c0<=a.c1)").map_err(crate::err)?.query_row([],|row|row.get(0)).map_err(err)?;
    let refusal = if qualified && count > 128 {
        Some("Qualified SEQUENCE ownership exceeds the 128-master workbook bound")
    } else if overlap {
        Some("Candidate spill constructors have intersecting intended output areas")
    } else {
        None
    };
    if let Some(reason) = refusal {
        // Opening an unowned workbook must retain unsupported arrays instead
        // of making spill authoring limits into file-format admission limits.
        // The ordinary authored-array guard keeps these formulas unresolved.
        // Existing owners cannot take this path: their complete followers must
        // stay fenced by atomic refusal, never silently become stale inputs.
        if qualified && !any_active(&tx, parent)? {
            return Ok(
                json!({"version":parent,"schema":SCHEMA,"groups":0,"refusedConstructors":refused+count,"reason":reason,"enabled":false}),
            );
        }
        return Err(crate::coded("spill-refused", reason));
    }
    tx.execute("INSERT INTO versions(parent) VALUES(?1)", [parent])
        .map_err(err)?;
    let version = tx.last_insert_rowid() as u64;
    let mut statement=tx.prepare_cached("SELECT sheet,mr,mc,r0,c0,r1,c1,old_range,formula,arguments,dependencies,blocked,blocker FROM spill_staged ORDER BY sheet,mr,mc").map_err(err)?;
    let mut rows = statement.query([]).map_err(err)?;
    let mut positions = 0_u64;
    let mut blocked = 0_u64;
    while let Some(row) = rows.next().map_err(err)? {
        crate::cancel::check()?;
        let master = (
            row.get::<_, u32>(1).map_err(err)?,
            row.get::<_, u32>(2).map_err(err)?,
        );
        let intended = Geometry {
            first: (row.get(3).map_err(err)?, row.get(4).map_err(err)?),
            last: (row.get(5).map_err(err)?, row.get(6).map_err(err)?),
        };
        let is_blocked = row.get::<_, bool>(11).map_err(err)?;
        let geometry = if is_blocked {
            blocked += 1;
            Geometry {
                first: master,
                last: master,
            }
        } else {
            intended.clone()
        };
        let arguments: Value =
            serde_json::from_str(&row.get::<_, String>(9).map_err(err)?).map_err(err)?;
        let dependencies: Value =
            serde_json::from_str(&row.get::<_, String>(10).map_err(err)?).map_err(err)?;
        let mut proof = json!({"constructor":"sequence-integer-v1","arguments":arguments,"dependencies":dependencies,"status":if is_blocked{"blocked"}else{"calculated"}});
        if qualified {
            proof["qualification"] = json!("retained-sequence-v1");
        }
        if is_blocked {
            proof["intendedRange"] = json!(intended.reference());
            proof["errorType"] = json!(8);
            proof["subType"] = json!(1);
            proof["blocker"] = json!(row.get::<_, String>(12).map_err(err)?);
        }
        let previous = row
            .get::<_, Option<String>>(7)
            .map_err(err)?
            .map(|reference| Geometry::from_reference(&reference))
            .transpose()?;
        let group = Group {
            version,
            sheet: row.get(0).map_err(err)?,
            master,
            geometry,
            formula: row.get(8).map_err(err)?,
            proof,
        };
        publish(&tx, &group, previous.as_ref(), revision)?;
        positions += group.geometry.count();
    }
    drop(rows);
    drop(statement);
    tx.commit().map_err(err)?;
    let mut result = json!({"version":version,"schema":SCHEMA,"groups":count,"outputPositions":positions,"blocked":blocked,"refusedConstructors":refused,"enabled":false});
    if qualified {
        result["owners"] = json!(export_groups(&db, version, &input["view"])?);
    }
    Ok(result)
}

#[cfg(test)]
mod tests {
    use super::*;
    fn mutated_fixture(
        part_name: &str,
        mutate: impl FnOnce(String) -> String,
    ) -> tempfile::TempDir {
        use std::io::Write;
        let directory = tempfile::tempdir().unwrap();
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../fixtures/excel-dynamic-reference.xlsx");
        let mut before = zip::ZipArchive::new(std::fs::File::open(path).unwrap()).unwrap();
        let mut after = zip::ZipWriter::new(
            std::fs::File::create(directory.path().join("source.xlsx")).unwrap(),
        );
        let mut mutate = Some(mutate);
        for index in 0..before.len() {
            let mut part = before.by_index(index).unwrap();
            let name = part.name().to_string();
            if name == part_name {
                let mut text = String::new();
                part.read_to_string(&mut text).unwrap();
                let changed = mutate.take().unwrap()(text);
                after
                    .start_file(&name, zip::write::SimpleFileOptions::default())
                    .unwrap();
                after.write_all(changed.as_bytes()).unwrap();
            } else {
                after.raw_copy_file(part).unwrap();
            }
        }
        assert!(mutate.is_none());
        after.finish().unwrap();
        crate::import::index(directory.path()).unwrap();
        directory
    }
    fn inspect_input() -> Value {
        json!({"version":0,"view":{"sheets":[{"name":"Cases","path":"xl/worksheets/sheet1.xml","merges":[],"cells":{}}],"tables":[],"namedRanges":[]}})
    }
    fn db() -> Connection {
        let db = Connection::open_in_memory().unwrap();
        db.execute_batch("CREATE TABLE versions(id INTEGER PRIMARY KEY,parent INTEGER); INSERT INTO versions VALUES(0,NULL),(1,0),(2,1),(3,1); CREATE TABLE changes(version INTEGER,sheet TEXT,r INTEGER,c INTEGER,data TEXT,edit TEXT,PRIMARY KEY(version,sheet,r,c));").unwrap();
        schema(&db).unwrap();
        db
    }
    fn group(version: u64, reference: &str) -> Group {
        let geometry = Geometry::from_reference(reference).unwrap();
        let arguments = vec![
            (geometry.last.0 - geometry.first.0 + 1) as f64,
            (geometry.last.1 - geometry.first.1 + 1) as f64,
        ];
        Group {
            version,
            sheet: "s".into(),
            master: geometry.first,
            geometry,
            formula: format!("_xlfn.SEQUENCE({},{})", arguments[0], arguments[1]),
            proof: json!({"constructor":"sequence-integer-v1","arguments":arguments}),
        }
    }
    fn seed(db: &Connection, group: &Group) {
        let transaction = db.is_autocommit();
        if transaction {
            db.execute_batch("BEGIN").unwrap();
        }
        let args = group.proof["arguments"]
            .as_array()
            .unwrap()
            .iter()
            .map(|arg| arg.as_f64().unwrap())
            .collect::<Vec<_>>();
        let sequence = Sequence::new(group.master, &args).unwrap();
        for (at, value) in group.geometry.positions().zip(sequence.values) {
            let mut cell = json!({"spill":group.cell_metadata(at).unwrap(),"calculation":{"status":"calculated","value":value},"value":value,"dataType":"number"});
            if at == group.master {
                cell["formula"] = json!(group.formula);
                cell["formulaType"] = json!("array");
                cell["formulaReference"] = json!(group.geometry.reference());
            }
            db.execute(
                "INSERT INTO changes VALUES(?1,?2,?3,?4,?5,'{}')",
                params![group.version, group.sheet, at.0, at.1, cell.to_string()],
            )
            .unwrap();
        }
        record(db, group, true).unwrap();
        if transaction {
            db.execute_batch("COMMIT").unwrap();
        }
    }
    fn seed_release(db: &Connection, group: &Group) {
        let transaction = db.is_autocommit();
        if transaction {
            db.execute_batch("BEGIN").unwrap();
        }
        for at in group.geometry.positions() {
            db.execute(
                "INSERT INTO changes VALUES(?1,?2,?3,?4,?5,?6)",
                params![
                    group.version,
                    group.sheet,
                    at.0,
                    at.1,
                    json!({"value":null,"dataType":"blank"}).to_string(),
                    json!({"value":null,"$spillRelease":true}).to_string()
                ],
            )
            .unwrap();
        }
        record(db, group, false).unwrap();
        if transaction {
            db.execute_batch("COMMIT").unwrap();
        }
    }
    #[test]
    fn exact_row_major_sequence_and_bounds() {
        let sequence = Sequence::new((0, 0), &[3.0, 2.0]).unwrap();
        assert_eq!(sequence.geometry.reference(), "A1:B3");
        assert_eq!(
            sequence.values,
            vec![
                json!(1.0),
                json!(2.0),
                json!(3.0),
                json!(4.0),
                json!(5.0),
                json!(6.0)
            ]
        );
        assert_eq!(
            Sequence::new((5, 0), &[2.0, 2.0, 10.0, 2.0])
                .unwrap()
                .values,
            vec![json!(10.0), json!(12.0), json!(14.0), json!(16.0)]
        );
        assert_eq!(
            Sequence::new((0, 0), &[2.0, 2.0, 3.0, -1.0])
                .unwrap()
                .values,
            vec![json!(3.0), json!(2.0), json!(1.0), json!(0.0)]
        );
        assert_eq!(
            Sequence::new((0, 0), &[1.0, 2.0, 0.0, 0.0]).unwrap().values,
            vec![json!(0.0), json!(0.0)]
        );
        assert_eq!(Sequence::new((0, 0), &[1000.0]).unwrap().values.len(), 1000);
        for args in [
            &[][..],
            &[0.0],
            &[-1.0],
            &[1.5],
            &[1001.0],
            &[32.0, 32.0],
            &[1.0, 2.0, 0.0, 0.5],
            &[1.0, 1.0, f64::INFINITY],
            &[1.0, 2.0, MAX_EXACT_INTEGER, 1.0],
            &[1.0, 3.0, -MAX_EXACT_INTEGER, MAX_EXACT_INTEGER],
            &[-0.0],
        ] {
            assert!(Sequence::new((0, 0), args).is_err(), "{args:?}");
        }
        assert!(Sequence::new((1_048_575, 0), &[2.0]).is_err());
        assert!(Sequence::new((0, 16_383), &[1.0, 2.0]).is_err());
    }
    #[test]
    fn immutable_owner_shrink_release_and_branch() {
        let db = db();
        seed(&db, &group(1, "A1:B3"));
        seed(&db, &group(2, "A1"));
        seed_release(&db, &group(3, "A1:B3"));
        assert_eq!(
            effective(&db, 1, "s").unwrap()[0].geometry.reference(),
            "A1:B3"
        );
        assert_eq!(
            effective(&db, 2, "s").unwrap()[0].geometry.reference(),
            "A1"
        );
        assert!(effective(&db, 3, "s").unwrap().is_empty());
        assert!(effective(&db, 0, "s").unwrap().is_empty());
        assert!(effective(&db, 99, "s").is_err());
        let owner = effective(&db, 1, "s").unwrap().remove(0);
        assert_eq!(
            owner.cell_metadata((2, 1)).unwrap(),
            json!({"schema":1,"master":"A1","range":"A1:B3","role":"follower"})
        );
        assert!(owner.cell_metadata((3, 0)).is_err());
    }
    #[test]
    fn followers_and_partial_groups_are_protected() {
        let db = db();
        seed(&db, &group(1, "A1:B3"));
        let guarded = |target: &str| guard(&db, 1, "s", &Geometry::from_reference(target).unwrap());
        for target in ["A1", "B1", "A2", "A1:B2", "A1:B3"] {
            assert!(guarded(target).is_err(), "{target}");
        }
        assert!(guarded("D1").is_ok());
        assert_eq!(
            blank_tombstone(
                &json!({"value":8,"formula":"bad","spill":{"role":"follower"},"calculation":{"value":8},"styleId":1,"bold":true,"rawValue":"8"})
            ),
            json!({"value":null,"dataType":"blank","styleId":1,"bold":true})
        );
    }
    #[test]
    fn transaction_rollback_and_invalid_owner_refusal() {
        let mut db = db();
        assert!(record(&db, &group(1, "A1:B3"), true).is_err());
        {
            let tx = db.transaction().unwrap();
            seed(&tx, &group(1, "A1:B3"));
        }
        assert!(effective(&db, 1, "s").unwrap().is_empty());
        assert_eq!(
            db.prepare_cached("SELECT count(*) FROM changes")
                .unwrap()
                .query_row([], |row| row.get::<_, u64>(0))
                .unwrap(),
            0
        );
        seed(&db, &group(1, "A1:B3"));
        db.execute(
            "INSERT INTO spill_groups VALUES(1,'s',0,1,1,0,1,1,1,'formula','{}',1)",
            [],
        )
        .unwrap();
        assert!(effective(&db, 1, "s").is_err());
        let bad = group(2, "D1");
        db.execute(
            "INSERT INTO spill_groups VALUES(2,'s',0,3,1,0,3,0,3,'formula','{}',2)",
            [],
        )
        .unwrap();
        assert!(effective(&db, 2, "s").is_err());
        assert!(record(&db, &bad, true).is_err());
    }
    #[test]
    fn cell_metadata_indirection_requires_dynamic_proof() {
        let xml = format!(
            "<metadata xmlns=\"{}\" xmlns:xda=\"http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray\"><metadataTypes count=\"2\"><metadataType name=\"opaque\"/><metadataType name=\"XLDAPR\" cellMeta=\"1\"/></metadataTypes><futureMetadata name=\"XLDAPR\" count=\"2\"><bk><extLst><ext uri=\"{{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}}\"><xda:dynamicArrayProperties fDynamic=\"0\"/></ext></extLst></bk><bk><extLst><ext uri=\"{{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}}\"><xda:dynamicArrayProperties fDynamic=\"1\"/></ext></extLst></bk></futureMetadata><cellMetadata count=\"3\"><bk><rc t=\"1\" v=\"0\"/></bk><bk><rc t=\"2\" v=\"1\"/></bk><bk><rc t=\"2\" v=\"0\"/></bk></cellMetadata></metadata>",
            String::from_utf8_lossy(crate::xml::NS)
        );
        assert_eq!(
            dynamic_cell_metadata(&xml).unwrap(),
            [2].into_iter().collect()
        );
        for bad in [
            xml.replace("t=\"2\"", "t=\"0\""),
            xml.replace("t=\"2\"", "t=\"02\""),
            xml.replace("v=\"1\"", "v=\"9\""),
            xml.replace("v=\"1\"", "v=\"01\""),
            xml.replace("cellMeta=\"1\"", "cellMeta=\"0\""),
            xml.replace("cellMeta=\"1\"", "cellMeta=\"true\""),
            xml.replace("fDynamic=\"1\"", "fDynamic=\"true\""),
            xml.replace("count=\"3\"", "count=\"4\""),
            xml.replace("count=\"3\"", "count=\"03\""),
            xml.replace(
                "<rc t=\"2\" v=\"1\"/>",
                "<rc t=\"2\" v=\"1\"/><rc t=\"2\" v=\"1\"/>",
            ),
            xml.replace("</metadata>", "<metadataTypes/><metadataTypes/></metadata>"),
            xml.replace("</metadata>", "<cellMetadata/><cellMetadata/></metadata>"),
            xml.replace("name=\"opaque\"", "name=\"XLDAPR\""),
            xml.replace(
                "<xda:dynamicArrayProperties fDynamic=\"1\"/>",
                "<unknown><xda:dynamicArrayProperties fDynamic=\"1\"/></unknown>",
            ),
        ] {
            assert!(dynamic_cell_metadata(&bad).is_err(), "{bad}");
        }
        assert!(dynamic_cell_metadata(&xml.replace("2017/dynamicarray", "2017/unknown")).is_err());
        assert!(dynamic_cell_metadata("<!DOCTYPE metadata [<!ENTITY a 'x'>]><metadata/>").is_err());
    }
    #[test]
    fn retained_actual_excel_source_and_reference_have_dynamic_cell_proof() {
        use std::io::Read;
        for name in ["excel-dynamic-source.xlsx", "excel-dynamic-reference.xlsx"] {
            let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../fixtures")
                .join(name);
            let mut zip = zip::ZipArchive::new(std::fs::File::open(path).unwrap()).unwrap();
            let mut xml = String::new();
            zip.by_name("xl/metadata.xml")
                .unwrap()
                .read_to_string(&mut xml)
                .unwrap();
            assert_eq!(
                dynamic_cell_metadata(&xml).unwrap(),
                [1].into_iter().collect()
            );
        }
    }
    #[test]
    fn retained_excel_constructor_inspection_is_exact_and_does_not_publish() {
        let oracle: Value = serde_json::from_slice(
            &std::fs::read(
                std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                    .join("../fixtures/excel-dynamic-reference.json"),
            )
            .unwrap(),
        )
        .unwrap();
        for file in ["excel-dynamic-source.xlsx", "excel-dynamic-reference.xlsx"] {
            let directory = tempfile::tempdir().unwrap();
            let root = directory.path();
            std::fs::copy(
                std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                    .join("../fixtures")
                    .join(file),
                root.join("source.xlsx"),
            )
            .unwrap();
            let original = crate::hash(&root.join("source.xlsx")).unwrap();
            crate::import::index(root).unwrap();
            let view = json!({"sheets":[{"name":"Cases","path":"xl/worksheets/sheet1.xml","merges":["B65:C65"],"cells":{}}],"tables":[],"namedRanges":[]});
            let result = inspect(root, &json!({"version":0,"view":view})).unwrap();
            assert_eq!(result["enabled"], false);
            assert_eq!(result["items"].as_array().unwrap().len(), 10);
            let item = |reference: &str| {
                result["items"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .find(|item| item["ref"] == reference)
                    .unwrap()
            };
            for reference in ["A1", "A6", "A56"] {
                assert_eq!(
                    item(reference)["status"],
                    "qualified-constructor",
                    "{file}: {reference}"
                );
                let shape =
                    Geometry::from_reference(item(reference)["range"].as_str().unwrap()).unwrap();
                for ((r, c), value) in shape
                    .positions()
                    .zip(item(reference)["values"].as_array().unwrap())
                {
                    assert_eq!(
                        value,
                        &oracle["cells"][address(r, c)]["expected"],
                        "{file}: {reference}"
                    );
                }
            }
            for (reference, blocker) in [("A41", "B42"), ("A60", "B61"), ("A69", "B70")] {
                assert_eq!(item(reference)["status"], "blocked");
                assert_eq!(item(reference)["blocker"], blocker);
            }
            for reference in ["A10", "A38", "A44", "A64"] {
                assert_eq!(item(reference)["status"], "refused", "{file}: {reference}");
            }
            assert_eq!(crate::hash(&root.join("source.xlsx")).unwrap(), original);
            let db = crate::database(root).unwrap();
            assert_eq!(
                db.prepare_cached("SELECT count(*) FROM versions")
                    .unwrap()
                    .query_row([], |row| row.get::<_, u64>(0))
                    .unwrap(),
                1
            );
            assert_eq!(
                db.prepare_cached("SELECT count(*) FROM changes")
                    .unwrap()
                    .query_row([], |row| row.get::<_, u64>(0))
                    .unwrap(),
                0
            );
            assert!(
                !present(&db).unwrap(),
                "Inspection does not allocate persistent ownership tables"
            );
        }
    }
    #[test]
    fn omitted_source_merges_and_opaque_cell_xml_cannot_bypass_adoption() {
        for (from, to) in [
            ("<c r=\"A1\" cm=\"1\">", "<c r=\"A1\" cm=\"01\">"),
            (
                "<c r=\"A1\" cm=\"1\">",
                "<c r=\"A1\" cm=\"1\" unknown=\"opaque\">",
            ),
            (
                "_xlfn.SEQUENCE(3,2)</f>",
                "_xlfn.SEQUENCE(3,2)</f><unknown>retained</unknown>",
            ),
            (
                "<c r=\"B1\"><v>2</v></c>",
                "<c r=\"B1\"><v unknown=\"opaque\">2</v></c>",
            ),
            (
                "<c r=\"B1\"><v>2</v></c>",
                "<c r=\"B1\"><v>2</v><unknown>retained</unknown></c>",
            ),
        ] {
            let directory = mutated_fixture("xl/worksheets/sheet1.xml", |text| {
                assert!(text.contains(from), "{from}");
                text.replace(from, to)
            });
            let original = crate::hash(&directory.path().join("source.xlsx")).unwrap();
            let result = inspect(directory.path(), &inspect_input()).unwrap();
            let item = result["items"]
                .as_array()
                .unwrap()
                .iter()
                .find(|item| item["ref"] == "A1")
                .unwrap();
            assert_eq!(item["status"], "refused", "{from}: {item}");
            let merged = result["items"]
                .as_array()
                .unwrap()
                .iter()
                .find(|item| item["ref"] == "A64")
                .unwrap();
            assert_eq!(merged["status"], "refused");
            assert!(merged["reason"].as_str().unwrap().contains("merge/table"));
            assert_eq!(
                crate::hash(&directory.path().join("source.xlsx")).unwrap(),
                original
            );
            let db = crate::database(directory.path()).unwrap();
            assert!(!present(&db).unwrap());
            assert_eq!(
                db.prepare_cached("SELECT count(*) FROM versions")
                    .unwrap()
                    .query_row([], |row| row.get::<_, u64>(0))
                    .unwrap(),
                1
            );
        }
    }
    #[test]
    fn metadata_package_ambiguity_refuses_without_allocating_or_changing_old_state() {
        for (part, from, to) in [
            (
                "xl/_rels/workbook.xml.rels",
                "relationships/sheetMetadata",
                "relationships/worksheet",
            ),
            (
                "xl/_rels/workbook.xml.rels",
                "Target=\"metadata.xml\"",
                "Target=\"metadata.xml\" TargetMode=\"External\"",
            ),
            ("xl/_rels/workbook.xml.rels", "Id=\"rId8\"", "Id=\"rId3\""),
            (
                "[Content_Types].xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml",
                "application/x-unknown",
            ),
            (
                "[Content_Types].xml",
                "</Types>",
                "<Override PartName=\"/xl/metadata.xml\" ContentType=\"application/x-unknown\"/></Types>",
            ),
        ] {
            let directory = mutated_fixture(part, |text| {
                assert!(text.contains(from));
                text.replace(from, to)
            });
            assert!(
                inspect(directory.path(), &inspect_input()).is_err(),
                "{part}: {from}"
            );
            let db = crate::database(directory.path()).unwrap();
            assert!(!present(&db).unwrap());
            assert_eq!(
                db.prepare_cached("SELECT count(*) FROM changes")
                    .unwrap()
                    .query_row([], |row| row.get::<_, u64>(0))
                    .unwrap(),
                0
            );
            assert_eq!(
                db.prepare_cached("SELECT count(*) FROM versions")
                    .unwrap()
                    .query_row([], |row| row.get::<_, u64>(0))
                    .unwrap(),
                1
            );
        }
    }
    #[test]
    fn publication_tombstones_shrink_and_clear_preserve_styles_old_versions_and_sources() {
        let directory = tempfile::tempdir().unwrap();
        let root = directory.path();
        std::fs::copy(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../fixtures/excel-dynamic-reference.xlsx"),
            root.join("source.xlsx"),
        )
        .unwrap();
        let original = crate::hash(&root.join("source.xlsx")).unwrap();
        crate::import::index(root).unwrap();
        let mut db = crate::database(root).unwrap();
        let path = "xl/worksheets/sheet1.xml";
        let shape = Geometry::from_reference("A56:B57").unwrap();
        assert!(
            publish(
                &db,
                &Group {
                    version: 1,
                    sheet: path.into(),
                    master: shape.first,
                    geometry: shape.clone(),
                    formula: "_xlfn.SEQUENCE(2,2)".into(),
                    proof: json!({"constructor":"sequence-integer-v1","arguments":[2.0,2.0]})
                },
                Some(&shape),
                "no-transaction"
            )
            .is_err()
        );
        let first = {
            let tx = db.transaction().unwrap();
            crate::calculation::snapshot(&tx, 0).unwrap();
            tx.execute("INSERT INTO versions(parent) VALUES(0)", [])
                .unwrap();
            let id = tx.last_insert_rowid() as u64;
            let group = Group {
                version: id,
                sheet: path.into(),
                master: shape.first,
                geometry: shape.clone(),
                formula: "_xlfn.SEQUENCE(2,2)".into(),
                proof: json!({"constructor":"sequence-integer-v1","arguments":[2.0,2.0]}),
            };
            publish(&tx, &group, Some(&shape), "first").unwrap();
            tx.commit().unwrap();
            id
        };
        db = crate::database(root).unwrap();
        let second = {
            let tx = db.transaction().unwrap();
            crate::calculation::snapshot(&tx, first).unwrap();
            tx.execute("INSERT INTO versions(parent) VALUES(?1)", [first])
                .unwrap();
            let id = tx.last_insert_rowid() as u64;
            let group = Group {
                version: id,
                sheet: path.into(),
                master: shape.first,
                geometry: Geometry::from_reference("A56").unwrap(),
                formula: "_xlfn.SEQUENCE(1,1)".into(),
                proof: json!({"constructor":"sequence-integer-v1","arguments":[1.0,1.0]}),
            };
            publish(&tx, &group, Some(&shape), "second").unwrap();
            tx.commit().unwrap();
            id
        };
        db = crate::database(root).unwrap();
        let old = crate::range::read_db(
            &db,
            &json!({"version":first,"sheet":path,"range":"A56:B57"}),
        )
        .unwrap();
        let new = crate::range::read_db(
            &db,
            &json!({"version":second,"sheet":path,"range":"A56:B57"}),
        )
        .unwrap();
        let cell = |page: &Value, reference: &str| {
            page["cells"]
                .as_array()
                .unwrap()
                .iter()
                .find(|item| item[0] == reference)
                .unwrap()[1]
                .clone()
        };
        assert_eq!(cell(&old, "B56")["value"], 2.0);
        assert!(cell(&new, "B56")["value"].is_null());
        assert!(cell(&new, "B56").get("spill").is_none());
        assert_eq!(cell(&old, "B56")["styleId"], cell(&new, "B56")["styleId"]);
        assert_eq!(cell(&new, "B56")["styleId"], 1);
        for reference in ["A57", "B57"] {
            assert!(cell(&new, reference)["value"].is_null());
            assert!(cell(&new, reference).get("calculation").is_none());
        }
        assert_eq!(
            effective(&db, first, path).unwrap()[0].geometry.reference(),
            "A56:B57"
        );
        assert_eq!(
            effective(&db, second, path).unwrap()[0]
                .geometry
                .reference(),
            "A56"
        );
        {
            let tx = db.transaction().unwrap();
            crate::calculation::snapshot(&tx, second).unwrap();
            tx.execute("INSERT INTO versions(parent) VALUES(?1)", [second])
                .unwrap();
            let id = tx.last_insert_rowid() as u64;
            let mut released = effective(&tx, second, path).unwrap().remove(0);
            released.version = id;
            for at in released.geometry.positions() {
                let old = crate::calculation::cell(&tx, path, &address(at.0, at.1)).unwrap();
                crate::range::store_calculated_change(
                    &tx,
                    id as i64,
                    path,
                    at.0,
                    at.1,
                    &blank_tombstone(&old),
                    &json!({"value":null,"$spillRelease":true}),
                )
                .unwrap();
            }
            record(&tx, &released, false).unwrap();
            tx.commit().unwrap();
            assert!(effective(&db, id, path).unwrap().is_empty());
            let cleared =
                crate::range::read_db(&db, &json!({"version":id,"sheet":path,"range":"A56:B57"}))
                    .unwrap();
            assert!(cell(&cleared, "A56")["value"].is_null());
            assert!(cell(&cleared, "A56").get("formula").is_none());
        }
        assert_eq!(crate::hash(&root.join("source.xlsx")).unwrap(), original);
    }
    #[test]
    fn blocked_publication_retains_literal_and_rollback_does_not_allocate_schema_or_results() {
        let directory = tempfile::tempdir().unwrap();
        let root = directory.path();
        std::fs::copy(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../fixtures/excel-dynamic-source.xlsx"),
            root.join("source.xlsx"),
        )
        .unwrap();
        crate::import::index(root).unwrap();
        let mut db = crate::database(root).unwrap();
        let path = "xl/worksheets/sheet1.xml";
        for commit in [false, true] {
            let tx = db.transaction().unwrap();
            crate::calculation::snapshot(&tx, 0).unwrap();
            tx.execute("INSERT INTO versions(parent) VALUES(0)", [])
                .unwrap();
            let id = tx.last_insert_rowid() as u64;
            let geometry = Geometry::from_reference("A41").unwrap();
            let group = Group {
                version: id,
                sheet: path.into(),
                master: geometry.first,
                geometry,
                formula: "_xlfn.SEQUENCE(2,2)".into(),
                proof: json!({"constructor":"sequence-integer-v1","arguments":[2.0,2.0],"status":"blocked","intendedRange":"A41:B42","errorType":8,"subType":1,"blocker":"B42"}),
            };
            publish(&tx, &group, None, "blocked").unwrap();
            if commit {
                tx.commit().unwrap();
                let page = crate::range::read_db(
                    &db,
                    &json!({"version":id,"sheet":path,"range":"A41:B42"}),
                )
                .unwrap();
                let master = &page["cells"][0][1];
                assert_eq!(master["calculation"]["error"], "#SPILL!");
                assert_eq!(master["spill"]["intendedRange"], "A41:B42");
                let literal = page["cells"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .find(|item| item[0] == "B42")
                    .unwrap();
                assert_eq!(literal[1]["value"], "occupied");
                assert!(literal[1].get("spill").is_none());
            } else {
                drop(tx);
                assert!(!present(&db).unwrap());
                assert_eq!(
                    db.prepare_cached("SELECT count(*) FROM versions")
                        .unwrap()
                        .query_row([], |row| row.get::<_, u64>(0))
                        .unwrap(),
                    1
                );
                assert_eq!(
                    db.prepare_cached("SELECT count(*) FROM changes")
                        .unwrap()
                        .query_row([], |row| row.get::<_, u64>(0))
                        .unwrap(),
                    0
                );
            }
        }
    }
    #[test]
    fn compaction_keeps_inherited_owner_and_released_branch_complete() {
        let directory = tempfile::tempdir().unwrap();
        let root = directory.path();
        std::fs::copy(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../fixtures/excel-dynamic-reference.xlsx"),
            root.join("source.xlsx"),
        )
        .unwrap();
        crate::import::index(root).unwrap();
        let mut db = crate::database(root).unwrap();
        let path = "xl/worksheets/sheet1.xml";
        let shape = Geometry::from_reference("A56:B57").unwrap();
        let owner = {
            let tx = db.transaction().unwrap();
            crate::calculation::snapshot(&tx, 0).unwrap();
            tx.execute("INSERT INTO versions(parent) VALUES(0)", [])
                .unwrap();
            let id = tx.last_insert_rowid() as u64;
            let group = Group {
                version: id,
                sheet: path.into(),
                master: shape.first,
                geometry: shape.clone(),
                formula: "_xlfn.SEQUENCE(2,2)".into(),
                proof: json!({"constructor":"sequence-integer-v1","arguments":[2.0,2.0]}),
            };
            publish(&tx, &group, Some(&shape), "first").unwrap();
            tx.commit().unwrap();
            id
        };
        let inherited = {
            db.execute("INSERT INTO versions(parent) VALUES(?1)", [owner])
                .unwrap();
            db.last_insert_rowid() as u64
        };
        db = crate::database(root).unwrap();
        let released = {
            let tx = db.transaction().unwrap();
            crate::calculation::snapshot(&tx, owner).unwrap();
            tx.execute("INSERT INTO versions(parent) VALUES(?1)", [owner])
                .unwrap();
            let id = tx.last_insert_rowid() as u64;
            let mut group = effective(&tx, owner, path).unwrap().remove(0);
            group.version = id;
            for at in shape.positions() {
                let old = crate::calculation::cell(&tx, path, &address(at.0, at.1)).unwrap();
                crate::range::store_calculated_change(
                    &tx,
                    id as i64,
                    path,
                    at.0,
                    at.1,
                    &blank_tombstone(&old),
                    &json!({"value":null,"$spillRelease":true}),
                )
                .unwrap();
            }
            record(&tx, &group, false).unwrap();
            tx.commit().unwrap();
            id
        };
        let read = |version| {
            crate::range::read_db(
                &db,
                &json!({"version":version,"sheet":path,"range":"A56:B57"}),
            )
            .unwrap()["cells"]
                .clone()
        };
        let before_inherited = read(inherited);
        let before_released = read(released);
        drop(db);
        let result =
            crate::compact::compact(root, &json!({"keep":[inherited,released],"vacuum":false}))
                .unwrap();
        assert_eq!(result["afterVersions"], 3);
        let db = crate::database(root).unwrap();
        assert_eq!(effective(&db, inherited, path).unwrap()[0].geometry, shape);
        assert!(effective(&db, released, path).unwrap().is_empty());
        assert_eq!(
            crate::range::read_db(
                &db,
                &json!({"version":inherited,"sheet":path,"range":"A56:B57"})
            )
            .unwrap()["cells"],
            before_inherited
        );
        assert_eq!(
            crate::range::read_db(
                &db,
                &json!({"version":released,"sheet":path,"range":"A56:B57"})
            )
            .unwrap()["cells"],
            before_released
        );
        assert!(effective(&db, owner, path).is_err());
        assert_eq!(
            db.prepare_cached("SELECT count(*) FROM spill_groups WHERE version=?1")
                .unwrap()
                .query_row([owner], |row| row.get::<_, u64>(0))
                .unwrap(),
            0
        );
    }
    #[test]
    fn constructor_cannot_borrow_cached_follower_even_outside_new_geometry_or_through_counts() {
        use std::io::Write;
        for argument in ["B3", "J1"] {
            for aggregate in ["SUM", "COUNT", "COUNTA", "COUNTBLANK"] {
                if argument == "B3" && aggregate != "SUM" {
                    continue;
                }
                let directory = tempfile::tempdir().unwrap();
                let root = directory.path();
                let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                    .join("../fixtures/excel-dynamic-reference.xlsx");
                let mut before = zip::ZipArchive::new(std::fs::File::open(path).unwrap()).unwrap();
                let mut after =
                    zip::ZipWriter::new(std::fs::File::create(root.join("source.xlsx")).unwrap());
                for index in 0..before.len() {
                    let mut part = before.by_index(index).unwrap();
                    let name = part.name().to_string();
                    if name == "xl/worksheets/sheet1.xml" {
                        let mut text = String::new();
                        part.read_to_string(&mut text).unwrap();
                        text = text.replace(
                            "_xlfn.SEQUENCE(3,2)",
                            &format!("_xlfn.SEQUENCE({argument})"),
                        );
                        if argument == "J1" {
                            let end = text.find("</row>").unwrap();
                            text.insert_str(
                                end,
                                &format!("<c r=\"J1\"><f>{aggregate}(B2:B3)</f><v>777</v></c>"),
                            );
                        }
                        after
                            .start_file(&name, zip::write::SimpleFileOptions::default())
                            .unwrap();
                        after.write_all(text.as_bytes()).unwrap();
                    } else {
                        after.raw_copy_file(part).unwrap();
                    }
                }
                after.finish().unwrap();
                crate::import::index(root).unwrap();
                let result=inspect(root,&json!({"version":0,"view":{"sheets":[{"name":"Cases","path":"xl/worksheets/sheet1.xml","merges":["B65:C65"],"cells":{}}],"tables":[],"namedRanges":[]}})).unwrap();
                let item = result["items"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .find(|item| item["ref"] == "A1")
                    .unwrap();
                assert_eq!(item["status"], "refused");
                assert!(
                    item["reason"]
                        .as_str()
                        .unwrap()
                        .contains("array/spill output"),
                    "{argument}: {aggregate}: {item}"
                );
            }
        }
    }
    #[test]
    fn private_stage_streams_complete_actual_groups_and_keeps_unsupported_arrays_protected() {
        let directory = tempfile::tempdir().unwrap();
        let root = directory.path();
        std::fs::copy(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../fixtures/excel-dynamic-reference.xlsx"),
            root.join("source.xlsx"),
        )
        .unwrap();
        let original = crate::hash(&root.join("source.xlsx")).unwrap();
        crate::import::index(root).unwrap();
        let mut input = inspect_input();
        input["revision"] = json!("private-stage");
        let result = stage(root, &input).unwrap();
        assert_eq!(result["enabled"], false);
        assert_eq!(result["groups"], 6);
        assert_eq!(result["blocked"], 3);
        assert_eq!(result["outputPositions"], 17);
        assert_eq!(result["refusedConstructors"], 4);
        let version = result["version"].as_u64().unwrap();
        let db = crate::database(root).unwrap();
        let path = "xl/worksheets/sheet1.xml";
        let owners = effective(&db, version, path).unwrap();
        let oracle: Value = serde_json::from_slice(
            &std::fs::read(
                std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                    .join("../fixtures/excel-dynamic-reference.json"),
            )
            .unwrap(),
        )
        .unwrap();
        for owner in owners {
            let page = crate::range::read_db(
                &db,
                &json!({"version":version,"sheet":path,"range":owner.geometry.reference()}),
            )
            .unwrap();
            for pair in page["cells"].as_array().unwrap() {
                let cell = &pair[1];
                let reference = pair[0].as_str().unwrap();
                let expected = &oracle["cells"][reference]["expected"];
                if owner.proof["status"] == "blocked" {
                    assert_eq!(cell["calculation"]["error"], *expected);
                    assert_eq!(cell["dataType"], "error");
                    assert_eq!(cell["rawValue"], "#VALUE!");
                } else {
                    assert_eq!(cell["calculation"]["value"], *expected);
                    assert_eq!(cell["value"], *expected);
                    assert_eq!(cell["dataType"], "number");
                }
                assert!(cell["spill"].is_object());
            }
        }
        for reference in ["A10", "A38", "A44", "A64"] {
            let page = crate::range::read_db(
                &db,
                &json!({"version":version,"sheet":path,"range":reference}),
            )
            .unwrap();
            assert!(page["cells"][0][1].get("spill").is_none());
        }
        assert_eq!(crate::hash(&root.join("source.xlsx")).unwrap(), original);
    }
    #[test]
    fn private_stage_resizes_constructor_inputs_and_refuses_invalid_existing_owner_atomically() {
        let directory = tempfile::tempdir().unwrap();
        let root = directory.path();
        std::fs::copy(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../fixtures/excel-sequence-initial-source.xlsx"),
            root.join("source.xlsx"),
        )
        .unwrap();
        crate::import::index(root).unwrap();
        let mut input = inspect_input();
        input["revision"] = json!("implementation-check-not-external-oracle");
        let initial = stage(root, &input).unwrap()["version"].as_u64().unwrap();
        let path = "xl/worksheets/sheet1.xml";
        input["version"] = json!(initial);
        validate_current(root, &crate::database(root).unwrap(), &input).unwrap();
        let changed=crate::range::patch(root,&json!({"version":initial,"changes":[{"sheet":path,"cell":"J1","value":{"value":4.0,"dataType":"number"},"edits":[{"value":4.0}]}]})).unwrap()["version"].as_u64().unwrap();
        input["version"] = json!(changed);
        assert!(
            validate_current(root, &crate::database(root).unwrap(), &input).is_err(),
            "Old complete outputs cannot certify changed constructor inputs"
        );
        let grown = stage(root, &input).unwrap()["version"].as_u64().unwrap();
        let db = crate::database(root).unwrap();
        assert_eq!(
            effective(&db, initial, path)
                .unwrap()
                .iter()
                .find(|owner| owner.master == position("A50").unwrap())
                .unwrap()
                .geometry
                .reference(),
            "A50:B51"
        );
        assert_eq!(
            effective(&db, grown, path)
                .unwrap()
                .iter()
                .find(|owner| owner.master == position("A50").unwrap())
                .unwrap()
                .geometry
                .reference(),
            "A50:B53"
        );
        drop(db);
        let changed=crate::range::patch(root,&json!({"version":grown,"changes":[{"sheet":path,"cell":"J1","value":{"value":1.0,"dataType":"number"},"edits":[{"value":1.0}]},{"sheet":path,"cell":"J2","value":{"value":1.0,"dataType":"number"},"edits":[{"value":1.0}]}]})).unwrap()["version"].as_u64().unwrap();
        input["version"] = json!(changed);
        let shrunk = stage(root, &input).unwrap()["version"].as_u64().unwrap();
        let db = crate::database(root).unwrap();
        assert_eq!(
            effective(&db, shrunk, path)
                .unwrap()
                .iter()
                .find(|owner| owner.master == position("A50").unwrap())
                .unwrap()
                .geometry
                .reference(),
            "A50"
        );
        let page = crate::range::read_db(
            &db,
            &json!({"version":shrunk,"sheet":path,"range":"A50:B53"}),
        )
        .unwrap();
        for pair in page["cells"].as_array().unwrap() {
            if pair[0] != "A50" {
                assert!(pair[1]["value"].is_null());
                assert!(pair[1].get("spill").is_none());
                assert!(pair[1].get("formula").is_none());
            }
        }
        drop(db);
        let invalid=crate::range::patch(root,&json!({"version":shrunk,"changes":[{"sheet":path,"cell":"J1","value":{"value":0.0,"dataType":"number"},"edits":[{"value":0.0}]}]})).unwrap()["version"].as_u64().unwrap();
        input["version"] = json!(invalid);
        let before = crate::database(root)
            .unwrap()
            .query_row("SELECT count(*) FROM versions", [], |row| {
                row.get::<_, u64>(0)
            })
            .unwrap();
        assert!(stage(root, &input).is_err());
        let db = crate::database(root).unwrap();
        assert_eq!(
            db.prepare_cached("SELECT count(*) FROM versions")
                .unwrap()
                .query_row([], |row| row.get::<_, u64>(0))
                .unwrap(),
            before
        );
        assert_eq!(
            effective(&db, shrunk, path)
                .unwrap()
                .iter()
                .find(|owner| owner.master == position("A50").unwrap())
                .unwrap()
                .geometry
                .reference(),
            "A50"
        );
    }
    #[test]
    fn staging_constructor_count_is_not_the_bounded_diagnostic_reply_limit() {
        let directory = mutated_fixture("xl/worksheets/sheet1.xml", |_| {
            let rows = (1..=256)
                .map(|row| {
                    format!("<row r=\"{row}\"><c r=\"A{row}\"><f>_xlfn.SEQUENCE(1)</f></c></row>")
                })
                .collect::<String>();
            format!(
                "<worksheet xmlns=\"{}\"><sheetData>{rows}</sheetData></worksheet>",
                String::from_utf8_lossy(crate::xml::NS)
            )
        });
        let root = directory.path();
        let mut input = inspect_input();
        input["revision"] = json!("implementation-check-not-external-oracle");
        assert!(inspect(root, &input).is_err());
        let before = crate::database(root).unwrap();
        assert!(!present(&before).unwrap());
        drop(before);
        let result = stage(root, &input).unwrap();
        assert_eq!(result["groups"], 256);
        assert_eq!(result["outputPositions"], 256);
        let db = crate::database(root).unwrap();
        assert_eq!(
            effective(
                &db,
                result["version"].as_u64().unwrap(),
                "xl/worksheets/sheet1.xml"
            )
            .unwrap()
            .len(),
            256
        );
    }
    #[test]
    fn ordinary_mutation_paths_cannot_write_or_copy_owned_outputs_and_master_release_is_complete() {
        let directory = tempfile::tempdir().unwrap();
        let root = directory.path();
        std::fs::copy(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../fixtures/excel-dynamic-reference.xlsx"),
            root.join("source.xlsx"),
        )
        .unwrap();
        crate::import::index(root).unwrap();
        let mut input = inspect_input();
        input["revision"] = json!("private-guard-check");
        let version = stage(root, &input).unwrap()["version"].as_u64().unwrap();
        let path = "xl/worksheets/sheet1.xml";
        let count = || {
            crate::database(root)
                .unwrap()
                .query_row("SELECT count(*) FROM versions", [], |row| {
                    row.get::<_, u64>(0)
                })
                .unwrap()
        };
        let before = count();
        for change in [
            json!({"sheet":path,"cell":"B56","value":{"value":99.0},"edits":[{"value":99.0}]}),
            json!({"sheet":path,"cell":"A56","value":{"value":null,"formula":"1+1"},"edits":[{"formula":"1+1"}]}),
            json!({"sheet":path,"cell":"Z1","value":{"value":1.0,"spill":{"schema":1,"master":"A56","range":"A56:B57","role":"follower"}},"edits":[{"value":1.0}]}),
        ] {
            assert!(
                crate::range::patch(root, &json!({"version":version,"changes":[change]})).is_err()
            );
            assert_eq!(count(), before);
        }
        for command in [
            json!({"type":"clearRange","sheet":path,"range":"A56:B57"}),
            json!({"type":"fillRange","sheet":path,"range":"Z1","source":"B56"}),
            json!({"type":"pasteFormulas","sheet":path,"range":"Z1:AA2","source":"A56:B57"}),
        ] {
            assert!(
                crate::bulk::apply(root, &json!({"version":version,"commands":[command]})).is_err()
            );
            assert_eq!(count(), before);
        }
        assert!(crate::transform::apply(root,&json!({"version":version,"kind":"sortRows","sheet":path,"range":"A56:B57","column":"A","direction":"asc","view":input["view"]})).is_err());
        assert!(
            release_master(root, &json!({"version":version,"sheet":path,"cell":"B56"})).is_err()
        );
        assert_eq!(count(), before);
        // Source-only Values uses a proved current scalar, without copying
        // ownership or formatting and without releasing/changing its source.
        let source_hash = crate::hash(&root.join("source.xlsx")).unwrap();
        let before_db = crate::database(root).unwrap();
        let source_before = crate::range::read_db(
            &before_db,
            &json!({"version":version,"sheet":path,"range":"A56:B57"}),
        )
        .unwrap();
        let target_before = crate::range::read_db(
            &before_db,
            &json!({"version":version,"sheet":path,"range":"Z1"}),
        )
        .unwrap();
        drop(before_db);
        let copied=crate::bulk::apply(root,&json!({"version":version,"commands":[{"type":"pasteValues","sheet":path,"range":"Z1","source":"B56"}]})).unwrap()["version"].as_u64().unwrap();
        let copied_db = crate::database(root).unwrap();
        let target = crate::range::read_db(
            &copied_db,
            &json!({"version":copied,"sheet":path,"range":"Z1"}),
        )
        .unwrap();
        assert_eq!(target["cells"][0][1]["value"], 2.0);
        assert_eq!(target["cells"][0][1]["dataType"], "number");
        assert!(target["cells"][0][1].get("spill").is_none());
        assert!(target["cells"][0][1].get("formula").is_none());
        assert_eq!(
            target["cells"][0][1]["styleId"],
            target_before["cells"]
                .as_array()
                .and_then(|cells| cells.first())
                .map(|cell| cell[1]["styleId"].clone())
                .unwrap_or(json!(0))
        );
        assert_eq!(
            crate::range::read_db(
                &copied_db,
                &json!({"version":version,"sheet":path,"range":"A56:B57"})
            )
            .unwrap(),
            source_before
        );
        validate_current(
            root,
            &crate::database(root).unwrap(),
            &json!({"version":version,"view":input["view"]}),
        )
        .unwrap();
        assert_eq!(crate::hash(&root.join("source.xlsx")).unwrap(), source_hash);
        drop(copied_db);
        let released = release_master(root, &json!({"version":version,"sheet":path,"cell":"A56"}))
            .unwrap()["version"]
            .as_u64()
            .unwrap();
        let db = crate::database(root).unwrap();
        let page = crate::range::read_db(
            &db,
            &json!({"version":released,"sheet":path,"range":"A56:B57"}),
        )
        .unwrap();
        for pair in page["cells"].as_array().unwrap() {
            assert!(pair[1]["value"].is_null());
            assert!(pair[1].get("spill").is_none());
            assert!(pair[1].get("formulaType").is_none());
        }
        let styled = page["cells"]
            .as_array()
            .unwrap()
            .iter()
            .find(|pair| pair[0] == "B56")
            .unwrap();
        assert_eq!(styled[1]["styleId"], 1);
        let old =
            crate::range::read_db(&db, &json!({"version":version,"sheet":path,"range":"B56"}))
                .unwrap();
        assert_eq!(old["cells"][0][1]["value"], 2.0);
        assert_eq!(old["cells"][0][1]["spill"]["role"], "follower");
        drop(db);
        let replaced=crate::range::patch(root,&json!({"version":released,"changes":[{"sheet":path,"cell":"A56","value":{"value":null,"formula":"_xlfn.SEQUENCE(1)"},"edits":[{"formula":"_xlfn.SEQUENCE(1)"}]}]})).unwrap()["version"].as_u64().unwrap();
        input["version"] = json!(replaced);
        let restored = stage(root, &input).unwrap()["version"].as_u64().unwrap();
        input["version"] = json!(restored);
        validate_current(root, &crate::database(root).unwrap(), &input).unwrap();
        let repeated = stage(root, &input).unwrap()["version"].as_u64().unwrap();
        let db = crate::database(root).unwrap();
        let owner = effective(&db, repeated, path)
            .unwrap()
            .into_iter()
            .find(|owner| owner.master == position("A56").unwrap())
            .unwrap();
        assert_eq!(owner.geometry.reference(), "A56");
        assert_eq!(owner.formula, "_xlfn.SEQUENCE(1)");
        assert!(any_history(&db, released).unwrap());
        assert!(
            any_active(&db, released).unwrap(),
            "Other independent owners remain active"
        );
        drop(db);
        std::fs::write(root.join("export-metadata.json"), b"{\"styleIds\":[]}").unwrap();
        let index: Value =
            serde_json::from_slice(&std::fs::read(root.join("index.json")).unwrap()).unwrap();
        let error = crate::export::write(
            root,
            &json!({"version":restored,"revision":index["revision"],"calculationMetadata":true}),
        )
        .unwrap_err();
        assert!(error.contains("Spill ownership export remains gated"));
        assert!(!root.join("export.xlsx").exists());
    }
}
