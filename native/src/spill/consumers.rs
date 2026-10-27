//! Readonly scalar proofs from complete current numeric spill ownership.
//! Portable syntax acceptance and caller-supplied cache fields are not proofs.
use super::*;
use crate::formula::ConsumerKind;
use std::collections::{BTreeMap, HashMap, HashSet};

pub(crate) struct Context<'a> {
    root: &'a Path,
    db: &'a Connection,
    version: u64,
    view: Value,
    owners: HashMap<(String, (u32, u32)), Group>,
    eligible: HashSet<u32>,
    source: BTreeMap<String, Vec<u8>>,
    barriers: HashMap<String, Vec<Geometry>>,
    qualified_sheets: HashSet<String>,
}
impl<'a> Context<'a> {
    pub(crate) fn new(
        root: &'a Path,
        db: &'a Connection,
        version: u64,
        view: &Value,
    ) -> Result<Self> {
        known_version(db, version)?;
        let index: Value =
            serde_json::from_slice(&std::fs::read(root.join("index.json")).map_err(err)?)
                .map_err(err)?;
        let mut owners = HashMap::new();
        let mut barriers = HashMap::new();
        let mut qualified_sheets = HashSet::new();
        for sheet in view["sheets"]
            .as_array()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing consumer worksheets"))?
        {
            let path = crate::field(sheet, "path")?;
            if index["sheets"]
                .as_array()
                .ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing consumer sheet index")
                })?
                .iter()
                .any(|item| item["path"] == path && item["extended"] != true)
            {
                qualified_sheets.insert(path.to_owned());
            }
            let mut blocked = source_barriers(root, path)?;
            for table in view["tables"].as_array().ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing consumer table metadata")
            })? {
                if table["sheet"] == sheet["name"] {
                    blocked.push(Geometry::from_reference(crate::field(table, "ref")?)?);
                }
            }
            barriers.insert(path.to_owned(), blocked);
            for group in effective(db, version, path)? {
                if group.proof["qualification"] != "retained-sequence-v1"
                    || !retained_arguments(&json!(validate_constructor(&group)?.arguments))
                {
                    return Err(crate::coded(
                        "spill-unsupported",
                        "Scalar consumer has no retained constructor qualification",
                    ));
                }
                validate_output_cells(db, version, &group)?;
                owners.insert((path.to_owned(), group.master), group);
            }
        }
        Ok(Self {
            root,
            db,
            version,
            view: view.clone(),
            owners,
            eligible: metadata_proof(root)?,
            source: metadata::source(root)?,
            barriers,
            qualified_sheets,
        })
    }
    fn source_admission(&self, sheet: &str, reference: &str, cell: &Value) -> Result<()> {
        let at = position(reference)?;
        if !self.qualified_sheets.contains(sheet) {
            return Err(crate::coded(
                "spill-unsupported",
                "Opaque consumer worksheet requires qualification",
            ));
        }
        if self
            .barriers
            .get(sheet)
            .is_none_or(|areas| areas.iter().any(|area| area.contains(at)))
        {
            return Err(crate::coded(
                "spill-unsupported",
                "Merge/table scalar consumers require qualification",
            ));
        }
        if self
            .owners
            .values()
            .any(|owner| owner.sheet == sheet && owner.geometry.contains(at))
        {
            return Err(crate::coded(
                "spill-unsupported",
                "A scalar consumer cannot replace a spill output",
            ));
        }
        if cell.get("formulaType").is_some() {
            if cell["formulaType"] != "array"
                || cell["formulaReference"] != reference
                || cell.get("scenarioResult").is_some()
            {
                return Err(crate::coded(
                    "spill-unsupported",
                    "Only imported singleton dynamic-array consumers are qualified",
                ));
            }
            let shape = source_group(self.root, self.db, sheet, reference, cell, &self.eligible)?
                .ok_or_else(|| {
                crate::coded(
                    "internal-missing-field",
                    "Missing singleton dynamic consumer source proof",
                )
            })?;
            if shape.first != at || shape.last != at {
                return Err(crate::coded(
                    "spill-bad-input",
                    "Scalar consumer source rectangle is not singleton",
                ));
            }
            let attrs = source_attributes(self.root, self.db, sheet, at)?.ok_or_else(|| {
                crate::coded(
                    "internal-missing-field",
                    "Missing singleton consumer source attributes",
                )
            })?;
            if attrs.get("vm").is_some() {
                return Err(crate::coded(
                    "spill-unsupported",
                    "Rich-error scalar consumer adoption is not qualified",
                ));
            }
            let cm = attrs
                .get("cm")
                .ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing scalar consumer cm proof")
                })?
                .parse::<u64>()
                .map_err(err)?;
            let mut masters = serde_json::Map::new();
            masters.insert(
                format!("{sheet}\0{reference}"),
                json!({"cm":cm,"range":reference}),
            );
            metadata::validate(&self.source, &Default::default(), &Value::Object(masters))?;
        } else {
            if source_attributes(self.root, self.db, sheet, at)?.is_some_and(|attrs| {
                attrs
                    .keys()
                    .any(|key| !["r", "s", "t", "xmlns"].contains(&key.as_str()))
                    || attrs
                        .get("xmlns")
                        .is_some_and(|value| value.as_bytes() != crate::xml::NS)
            }) {
                return Err(crate::coded(
                    "spill-unsupported",
                    "Opaque scalar consumer attributes require qualification",
                ));
            }
            source_content_proof(self.root, self.db, sheet, at, true)?;
        }
        Ok(())
    }
    fn calculate(&self, sheet: &str, reference: &str, cell: &Value) -> Result<(f64, Value)> {
        let formula = crate::field(cell, "formula")?;
        let consumer = crate::formula::spill_consumer(formula)?.ok_or_else(|| {
            crate::coded(
                "spill-unsupported",
                "Formula is outside the scalar spill consumer grammar",
            )
        })?;
        self.source_admission(sheet, reference, cell)?;
        let owner = self
            .owners
            .get(&(sheet.to_owned(), consumer.anchor))
            .ok_or_else(|| {
                crate::coded(
                    "spill-bad-input",
                    "Spill anchor is not a complete accepted same-sheet master",
                )
            })?;
        if owner.proof["status"] != "calculated" {
            return Err(crate::coded(
                "spill-protected",
                "Blocked spill scalar consumers remain protected",
            ));
        }
        let sequence = validate_constructor(owner)?;
        let (kind, axes, value) = match consumer.kind {
            ConsumerKind::Sum => (
                "SUM",
                Vec::<u32>::new(),
                sequence.values.iter().try_fold(0.0, |sum, value| {
                    Ok::<_, String>(
                        sum + value.as_f64().ok_or_else(|| {
                            crate::coded("spill-unsupported", "Non-numeric accepted spill output")
                        })?,
                    )
                })?,
            ),
            ConsumerKind::Index(row, column) => {
                let rows = owner.geometry.last.0 - owner.master.0 + 1;
                let columns = owner.geometry.last.1 - owner.master.1 + 1;
                if row > rows || column > columns {
                    return Err(crate::coded(
                        "spill-unsupported",
                        "Spill INDEX axes are outside current ownership geometry",
                    ));
                }
                (
                    "INDEX",
                    vec![row, column],
                    sequence.values[((row - 1) * columns + column - 1) as usize]
                        .as_f64()
                        .ok_or_else(|| {
                            crate::coded("spill-unsupported", "Non-numeric accepted spill output")
                        })?,
                )
            }
        };
        // Retained vectors are exact positive integers. Canonical integer wire
        // fields survive JSON round trips without int/f64 representation drift.
        let producer_arguments = sequence.arguments.map(|value| value as u32);
        let proof = json!({"schema":1,"kind":kind,"formula":formula,"canonicalFormula":consumer.canonical,"master":address(owner.master.0,owner.master.1),"range":owner.geometry.reference(),"axes":axes,"producerFormula":owner.formula,"producerArguments":producer_arguments});
        Ok((value, proof))
    }
    /// Recompute at the export version. Proofs contain stable ownership content,
    /// not private version IDs which compaction may legally rebase.
    pub(crate) fn validate(
        &self,
        sheet: &str,
        reference: &str,
        cell: &Value,
    ) -> Result<Option<String>> {
        if cell["calculation"]["spillConsumer"].is_object()
            && crate::calculation_transfer::publication_certificate(
                self.root,
                self.db,
                self.version,
                sheet,
                reference,
                cell,
            )?
        {
            return Ok(Some(
                crate::field(&cell["calculation"]["spillConsumer"], "canonicalFormula")?.into(),
            ));
        }
        let marker = cell["calculation"].get("spillConsumer");
        let syntax = cell["formula"].as_str().map(crate::formula::spill_consumer);
        if marker.is_none() {
            if matches!(syntax, Some(Ok(Some(_))) | Some(Err(_)))
                && cell["calculation"]["status"] == "calculated"
            {
                return Err(crate::coded(
                    "spill-bad-input",
                    "Scalar spill cache has no independent native proof",
                ));
            }
            return Ok(None);
        }
        let (value, expected) = if let Some(version) = crate::calculation_policy::cache_ancestor(
            self.db,
            self.version,
            sheet,
            reference,
            cell,
            "spillConsumer",
        )? {
            Context::new(self.root, self.db, version, &self.view)?
                .calculate(sheet, reference, cell)?
        } else {
            self.calculate(sheet, reference, cell)?
        };
        if marker != Some(&expected)
            || cell["calculation"]["status"] != "calculated"
            || cell["calculation"]["value"].as_f64() != Some(value)
        {
            return Err(crate::coded(
                "spill-stale",
                "Scalar spill consumer cache/ownership proof is stale or forged",
            ));
        }
        Ok(Some(
            crate::field(&expected, "canonicalFormula")?.to_owned(),
        ))
    }
}
pub(crate) fn command(root: &Path, input: &Value) -> Result<Value> {
    let db = crate::database(root)?;
    let version = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing scalar consumer version"))?;
    let items = input["items"].as_array().ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing scalar consumer coordinates",
        )
    })?;
    if items.len() > 64 {
        return Err(crate::coded(
            "spill-limit",
            "Scalar consumer proof pages use at most64 coordinates",
        ));
    }
    validate_current(
        root,
        &db,
        &input["view"]
            .as_object()
            .map(|_| json!({"version":version,"view":input["view"]}))
            .ok_or_else(|| {
                crate::coded(
                    "internal-missing-field",
                    "Missing consumer worksheet metadata",
                )
            })?,
    )?;
    // validate_current already built the complete version's calculation scratch.
    let context = Context::new(root, &db, version, &input["view"])?;
    let mut results = Vec::new();
    let mut refused = Vec::new();
    let mut seen = HashSet::new();
    for item in items {
        crate::cancel::check()?;
        let sheet = crate::field(item, "sheet")?;
        let reference = crate::field(item, "ref")?;
        if item.as_object().is_none_or(|fields| fields.len() != 2)
            || !seen.insert((sheet, reference))
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid or duplicate scalar consumer coordinate",
            ));
        }
        let cell = crate::calculation::cell(&db, sheet, reference)?;
        match context.calculate(sheet,reference,&cell){Ok((value,proof))=>results.push(json!({"sheet":sheet,"ref":reference,"formula":cell["formula"],"calculation":{"status":"calculated","value":value,"revision":input["revision"],"spillConsumer":proof}})),Err(reason)=>refused.push(json!({"sheet":sheet,"ref":reference,"reason":reason}))}
    }
    Ok(json!({"version":context.version,"results":results,"refused":refused}))
}

/// Final stream-export admission also checks prepared cache serialization. The
/// accepted marker set comes from version ancestry, not mutable export scratch.
pub(crate) fn validate_export(root: &Path, db: &Connection, version: u64) -> Result<()> {
    if !any_history(db, version)? {
        return Ok(());
    }
    let raw: String = db
        .prepare_cached("SELECT spill_view FROM export_state WHERE version=?1")
        .map_err(crate::err)?
        .query_row([version], |row| row.get(0))
        .map_err(err)?;
    let view: Value = serde_json::from_str(&raw).map_err(err)?;
    let context = Context::new(root, db, version, &view)?;
    let mut statement=db.prepare_cached("WITH RECURSIVE ancestry(id,depth) AS (SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN ancestry a ON v.id=a.id WHERE v.parent IS NOT NULL), ranked AS (SELECT c.*,row_number() OVER (PARTITION BY sheet,r,c ORDER BY depth) priority FROM changes c JOIN ancestry a ON c.version=a.id) SELECT sheet,r,c FROM ranked WHERE priority=1 AND json_extract(data,'$.calculation.spillConsumer') IS NOT NULL").map_err(err)?;
    let mut rows = statement.query([version]).map_err(err)?;
    let mut accepted = HashSet::new();
    while let Some(row) = rows.next().map_err(err)? {
        crate::cancel::check()?;
        let sheet: String = row.get(0).map_err(err)?;
        let at = (
            row.get::<_, u32>(1).map_err(err)?,
            row.get::<_, u32>(2).map_err(err)?,
        );
        let reference = address(at.0, at.1);
        let read = crate::range::read_db(
            db,
            &json!({"version":version,"sheet":sheet,"range":reference,"limit":1}),
        )?;
        let cell = &read["cells"][0][1];
        let canonical = context.validate(&sheet, &reference, cell)?.ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing accepted scalar spill proof",
            )
        })?;
        let prepared: Option<String> = db
            .prepare_cached("SELECT data FROM export_changes WHERE sheet=?1 AND r=?2 AND c=?3")
            .map_err(crate::err)?
            .query_row(params![sheet, at.0, at.1], |row| row.get(0))
            .optional()
            .map_err(err)?;
        let prepared: Value = serde_json::from_str(&prepared.ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing prepared scalar spill cache",
            )
        })?)
        .map_err(err)?;
        if cell["formulaType"] == "array" {
            if prepared.get("formula").is_some()
                || prepared.get("value").is_some()
                || prepared
                    .get("cache")
                    .is_some_and(|cache| cache != &cell["calculation"])
                || prepared.get("cache").is_none()
                    && cell["calculation"]["value"].as_f64() != cell["cachedValue"].as_f64()
            {
                return Err(crate::coded(
                    "spill-stale",
                    "Singleton scalar spill cache changed source XML or omitted its fresh result",
                ));
            }
        } else if prepared["formula"].as_str() != Some(canonical.as_str())
            || prepared["calculation"] != cell["calculation"]
        {
            return Err(crate::coded(
                "internal-mismatch",
                "Prepared scalar spill formula/cache differs from accepted proof",
            ));
        }
        accepted.insert((sheet, at.0, at.1));
    }
    let mut rows=db.prepare_cached("SELECT sheet,r,c FROM export_changes WHERE json_extract(data,'$.calculation.spillConsumer') IS NOT NULL OR json_extract(data,'$.cache.spillConsumer') IS NOT NULL").map_err(err)?;
    let mut rows = rows.query([]).map_err(err)?;
    while let Some(row) = rows.next().map_err(err)? {
        if !accepted.contains(&(
            row.get::<_, String>(0).map_err(err)?,
            row.get::<_, u32>(1).map_err(err)?,
            row.get::<_, u32>(2).map_err(err)?,
        )) {
            return Err(crate::coded(
                "spill-bad-input",
                "Prepared scalar spill cache has no accepted owner proof",
            ));
        }
    }
    Ok(())
}
