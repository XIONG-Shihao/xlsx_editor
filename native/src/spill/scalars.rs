//! Version-fenced scalar-array evaluation and independent export revalidation.
use super::*;
use sha2::{Digest, Sha256};
use std::collections::{HashMap, HashSet};
pub(crate) struct Context<'a> {
    root: &'a Path,
    db: &'a Connection,
    version: u64,
    sheets: HashSet<String>,
    barriers: HashMap<String, Vec<Geometry>>,
    literals: HashMap<(String, String), (Value, Value)>,
    literal_bytes: usize,
}
impl<'a> Context<'a> {
    pub(crate) fn new(root: &'a Path, db: &'a Connection, version: u64) -> Result<Self> {
        known_version(db, version)?;
        let index: Value =
            serde_json::from_slice(&std::fs::read(root.join("index.json")).map_err(err)?)
                .map_err(err)?;
        let sheets = index["sheets"]
            .as_array()
            .ok_or_else(|| {
                crate::coded(
                    "internal-missing-field",
                    "Missing scalar source sheet index",
                )
            })?
            .iter()
            .filter(|s| s["extended"] != true)
            .map(|s| crate::field(s, "path").map(str::to_owned))
            .collect::<Result<HashSet<_>>>()?;
        Ok(Self {
            root,
            db,
            version,
            sheets,
            barriers: HashMap::new(),
            literals: HashMap::new(),
            literal_bytes: 0,
        })
    }
    fn barriers(&mut self, sheet: &str) -> Result<&Vec<Geometry>> {
        if !self.sheets.contains(sheet) {
            return Err(crate::coded(
                "spill-unsupported",
                "Opaque scalar-array worksheet requires qualification",
            ));
        }
        if !self.barriers.contains_key(sheet) {
            let mut areas = source_barriers(self.root, sheet)?;
            let mut statement = self
                .db
                .prepare_cached(
                    "SELECT formula_attrs FROM cells WHERE sheet=?1 AND formula_attrs IS NOT NULL",
                )
                .map_err(err)?;
            let rows = statement
                .query_map([sheet], |r| r.get::<_, String>(0))
                .map_err(err)?;
            for row in rows {
                let attrs: Value = serde_json::from_str(&row.map_err(err)?).map_err(err)?;
                if attrs.get("t").is_some() {
                    if let Some(reference) = attrs["ref"].as_str() {
                        areas.push(Geometry::from_reference(reference)?);
                    } else {
                        return Err(crate::coded(
                            "spill-protected",
                            "Special formula without output rectangle requires protection",
                        ));
                    }
                }
            }
            for owner in effective(self.db, self.version, sheet)? {
                areas.push(owner.geometry);
                if let Some(range) = owner.proof["intendedRange"].as_str() {
                    areas.push(Geometry::from_reference(range)?);
                }
            }
            if areas.len() > 10000 {
                return Err(crate::coded(
                    "spill-limit",
                    "Scalar source barrier count exceeds its bound",
                ));
            }
            self.barriers.insert(sheet.into(), areas);
        }
        Ok(&self.barriers[sheet])
    }
    fn cell(&self, sheet: &str, reference: &str) -> Result<Value> {
        let page = crate::range::read_db(
            self.db,
            &json!({"version":self.version,"sheet":sheet,"range":reference,"limit":2}),
        )?;
        Ok(page["cells"]
            .as_array()
            .and_then(|cells| cells.first())
            .map(|r| r[1].clone())
            .unwrap_or(Value::Null))
    }
    fn source_admission(&mut self, sheet: &str, reference: &str, cell: &Value) -> Result<()> {
        let at = position(reference)?;
        let special = cell.get("formulaType").is_some();
        // Imported singleton output may coincide with its own source barrier only.
        let own_shape = if special {
            let eligible = metadata_proof(self.root)?;
            let shape = source_group(self.root, self.db, sheet, reference, cell, &eligible)?
                .ok_or_else(|| {
                    crate::coded(
                        "internal-missing-field",
                        "Missing scalar array metadata ownership",
                    )
                })?;
            if shape.first != at || shape.last != at {
                return Err(crate::coded(
                    "spill-unsupported",
                    "Only singleton dynamic scalar outputs are qualified",
                ));
            }
            let attrs = source_attributes(self.root, self.db, sheet, at)?.ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing scalar master attributes")
            })?;
            if attrs.contains_key("vm") {
                return Err(crate::coded(
                    "spill-unsupported",
                    "Rich-error scalar source requires qualification",
                ));
            }
            let cm = attrs["cm"].parse::<u64>().map_err(err)?;
            let mut mappings = serde_json::Map::new();
            mappings.insert(
                format!("{sheet}\0{reference}"),
                json!({"cm":cm,"range":reference}),
            );
            metadata::validate(
                &metadata::source(self.root)?,
                &Default::default(),
                &Value::Object(mappings),
            )?;
            Some(shape)
        } else {
            None
        };
        let mut own = 0;
        for area in self.barriers(sheet)? {
            if area.contains(at) {
                if own_shape.as_ref() == Some(area) && own == 0 {
                    own += 1;
                } else {
                    return Err(crate::coded(
                        "spill-conflict",
                        "Scalar array intersects protected merge/table/output geometry",
                    ));
                }
            }
        }
        if !special {
            if source_attributes(self.root, self.db, sheet, at)?.is_some_and(|attrs| {
                attrs
                    .keys()
                    .any(|key| !["r", "s", "t", "xmlns"].contains(&key.as_str()))
                    || attrs
                        .get("xmlns")
                        .is_some_and(|ns| ns.as_bytes() != crate::xml::NS)
            }) {
                return Err(crate::coded(
                    "spill-unsupported",
                    "Opaque scalar-array master attributes require qualification",
                ));
            }
            source_content_proof(self.root, self.db, sheet, at, true)?;
        }
        Ok(())
    }
    /// A context owns one immutable version; memoized literal proofs never cross
    /// an operation/version boundary. Both entry count and serialized payload bound
    /// retained memory while repeated lookups avoid repeated source/XML reads.
    fn literal(&mut self, sheet: &str, reference: &str, at: (u32, u32)) -> Result<(Value, Value)> {
        let key = (sheet.to_owned(), reference.to_owned());
        if let Some(value) = self.literals.get(&key) {
            return Ok(value.clone());
        }
        let cell = self.cell(sheet, reference)?;
        if cell.is_null()
            || cell.get("formula").is_some()
            || cell.get("formulaType").is_some()
            || cell.get("spill").is_some()
            || !["number", "text", "boolean"].contains(&cell["dataType"].as_str().unwrap_or(""))
        {
            return Err(crate::coded(
                "spill-bad-input",
                "Scalar array input requires a stored typed literal without coercion",
            ));
        }
        if source_attributes(self.root, self.db, sheet, at)?.is_some_and(|attrs| {
            attrs
                .keys()
                .any(|key| !["r", "s", "t", "xmlns"].contains(&key.as_str()))
                || attrs
                    .get("xmlns")
                    .is_some_and(|ns| ns.as_bytes() != crate::xml::NS)
        }) {
            return Err(crate::coded(
                "spill-unsupported",
                "Opaque scalar array input attributes require qualification",
            ));
        }
        source_content_proof(self.root, self.db, sheet, at, false)?;
        let value = crate::bulk::literal(&cell, false)?;
        if value.as_str().is_some_and(|text| text.len() > 4096) {
            return Err(crate::coded(
                "spill-limit",
                "Scalar array text input exceeds4 KiB",
            ));
        }
        let normalized = if let Some(number) = value.as_f64() {
            if !number.is_finite()
                || number.abs() > MAX_EXACT_INTEGER
                || number == 0.0 && number.is_sign_negative()
            {
                return Err(crate::coded(
                    "spill-limit",
                    "Scalar input is outside finite precision bounds",
                ));
            }
            if number.fract() == 0.0 {
                json!(number as i64)
            } else {
                json!(number)
            }
        } else {
            value.clone()
        };
        let item = json!({"ref":reference,"type":cell["dataType"],"value":normalized});
        let bytes = serde_json::to_vec(&item).map_err(err)?.len()
            + serde_json::to_vec(&value).map_err(err)?.len()
            + sheet.len()
            + reference.len()
            + 128;
        if self.literals.len() < 8192 && self.literal_bytes + bytes <= 2 * 1024 * 1024 {
            self.literal_bytes += bytes;
            self.literals.insert(key, (value.clone(), item.clone()));
        }
        Ok((value, item))
    }
    fn calculate(&mut self, sheet: &str, reference: &str, cell: &Value) -> Result<Option<Value>> {
        let Some(formula) = cell["formula"].as_str() else {
            return Ok(None);
        };
        if !formula.to_ascii_uppercase().contains("XLOOKUP")
            && !formula.to_ascii_uppercase().contains("SEQUENCE")
            && !formula.to_ascii_uppercase().contains("UNIQUE")
        {
            return Ok(None);
        }
        let mut inputs = Vec::new();
        let mut input_bytes = 0usize;
        let mut cache: HashMap<String, Value> = HashMap::new();
        let areas = self.barriers(sheet)?.clone();
        let result = crate::formula::scalar_array_result(formula, |reference| {
            if let Some(value) = cache.get(reference) {
                return Ok(value.clone());
            }
            let at = position(reference)?;
            if areas.iter().any(|area| area.contains(at)) {
                return Err(crate::coded(
                    "spill-protected",
                    "Scalar array input reads protected array/merge/table geometry",
                ));
            }
            let (value, item) = self.literal(sheet, reference, at)?;
            if inputs.len() >= 2004 {
                return Err(crate::coded(
                    "spill-limit",
                    "Scalar array input proof exceeds its bound",
                ));
            }
            input_bytes += serde_json::to_vec(&item).map_err(err)?.len();
            if input_bytes > 512 * 1024 {
                return Err(crate::coded(
                    "spill-limit",
                    "Scalar array input proof exceeds512 KiB",
                ));
            }
            inputs.push(item);
            cache.insert(reference.to_owned(), value.clone());
            Ok(value)
        })?;
        let Some((value, kind, canonical)) = result else {
            return Ok(None);
        };
        self.source_admission(sheet, reference, cell)?;
        let proof = json!({"schema":1,"formula":formula,"canonicalFormula":canonical,"valueType":kind,"inputCells":inputs.len(),"inputSha256":format!("{:x}",Sha256::digest(serde_json::to_vec(&inputs).map_err(err)?))});
        Ok(Some(if kind == "error" {
            json!({"status":"error","error":value,"scalarArray":proof})
        } else {
            json!({"status":"calculated","value":value,"scalarArray":proof})
        }))
    }
    pub(crate) fn validate(
        &mut self,
        sheet: &str,
        reference: &str,
        cell: &Value,
    ) -> Result<Option<String>> {
        if cell["calculation"]["scalarArray"].is_object()
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
                crate::field(&cell["calculation"]["scalarArray"], "canonicalFormula")?.into(),
            ));
        }
        let marker = cell["calculation"].get("scalarArray");
        if marker.is_none() {
            if cell["formula"]
                .as_str()
                .is_some_and(crate::formula::scalar_array_root)
                && matches!(
                    cell["calculation"]["status"].as_str(),
                    Some("calculated" | "error")
                )
            {
                return Err(crate::coded(
                    "spill-bad-input",
                    "Scalar array result has no independent native proof",
                ));
            }
            return Ok(None);
        }
        let expected = if let Some(version) = crate::calculation_policy::cache_ancestor(
            self.db,
            self.version,
            sheet,
            reference,
            cell,
            "scalarArray",
        )? {
            let mut prior = Context::new(self.root, self.db, version)?;
            prior.calculate(sheet, reference, cell)?
        } else {
            self.calculate(sheet, reference, cell)?
        }
        .ok_or_else(|| {
            crate::coded(
                "spill-unsupported",
                "Scalar array cache is outside qualified grammar",
            )
        })?;
        let value_matches = if expected["status"] == "error" {
            cell["calculation"]["error"] == expected["error"]
        } else if let Some(value) = expected["value"].as_f64() {
            cell["calculation"]["value"].as_f64() == Some(value)
        } else {
            cell["calculation"]["value"] == expected["value"]
        };
        if marker != expected.get("scalarArray")
            || cell["calculation"]["status"] != expected["status"]
            || !value_matches
        {
            return Err(crate::coded(
                "spill-stale",
                "Scalar array cache/input proof is stale or forged",
            ));
        }
        Ok(Some(
            crate::field(&expected["scalarArray"], "canonicalFormula")?.into(),
        ))
    }
}
pub(crate) fn command(root: &Path, input: &Value) -> Result<Value> {
    let db = crate::database(root)?;
    let version = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing scalar array version"))?;
    let items = input["items"].as_array().ok_or_else(|| {
        crate::coded("internal-missing-field", "Missing scalar array coordinates")
    })?;
    if items.len() > 64 {
        return Err(crate::coded(
            "spill-limit",
            "Scalar proof page exceeds64 coordinates",
        ));
    }
    let mut context = Context::new(root, &db, version)?;
    let mut seen = HashSet::new();
    let mut results = Vec::new();
    let mut refused = Vec::new();
    for item in items {
        crate::cancel::check()?;
        let sheet = crate::field(item, "sheet")?;
        let reference = crate::field(item, "ref")?;
        if item.as_object().is_none_or(|fields| fields.len() != 2)
            || !seen.insert((sheet, reference))
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid scalar proof coordinates",
            ));
        }
        let cell = context.cell(sheet, reference)?;
        match context.calculate(sheet, reference, &cell) {
            Ok(Some(mut calculation)) => {
                calculation["revision"] = input["revision"].clone();
                results.push(json!({"sheet":sheet,"ref":reference,"formula":cell["formula"],"calculation":calculation}));
            }
            Ok(None) => (),
            Err(reason) => refused.push(json!({"sheet":sheet,"ref":reference,"reason":reason})),
        }
    }
    Ok(json!({"version":version,"results":results,"refused":refused}))
}
/// Check prepared caches against accepted nearest-ancestor records, never scratch alone.
pub(crate) fn validate_export(root: &Path, db: &Connection, version: u64) -> Result<()> {
    let mut context = Context::new(root, db, version)?;
    let mut accepted = HashSet::new();
    let mut after = (String::new(), 0u32, 0u32);
    loop {
        // Range reads prepare request-local ancestry/shared tables. Close the window
        // cursor before those TEMP writes; otherwise SQLite can abort its sorter
        // with SQLITE_ABORT_ROLLBACK. Immutable accepted versions permit keyset pages.
        let page = {
            let mut statement=db.prepare_cached("WITH RECURSIVE ancestry(id,depth) AS (SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN ancestry a ON v.id=a.id WHERE v.parent IS NOT NULL), ranked AS (SELECT c.*,row_number() OVER(PARTITION BY sheet,r,c ORDER BY depth) priority FROM changes c JOIN ancestry a ON c.version=a.id) SELECT sheet,r,c FROM ranked WHERE priority=1 AND json_extract(data,'$.calculation.scalarArray') IS NOT NULL AND (sheet,r,c)>(?2,?3,?4) ORDER BY sheet,r,c LIMIT 64").map_err(err)?;
            statement
                .query_map(params![version, after.0, after.1, after.2], |row| {
                    Ok((
                        row.get::<_, String>(0)?,
                        row.get::<_, u32>(1)?,
                        row.get::<_, u32>(2)?,
                    ))
                })
                .map_err(err)?
                .collect::<std::result::Result<Vec<_>, _>>()
                .map_err(err)?
        };
        if page.is_empty() {
            break;
        }
        after = page.last().unwrap().clone();
        for (sheet, r, c) in &page {
            crate::cancel::check()?;
            let reference = address(*r, *c);
            let cell = context.cell(sheet, &reference)?;
            let canonical = context.validate(sheet, &reference, &cell)?.ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing scalar array proof")
            })?;
            let raw: Option<String> = db
                .prepare_cached("SELECT data FROM export_changes WHERE sheet=?1 AND r=?2 AND c=?3")
                .map_err(crate::err)?
                .query_row(params![sheet, r, c], |row| row.get(0))
                .optional()
                .map_err(err)?;
            let prepared: Value = serde_json::from_str(&raw.ok_or_else(|| {
                crate::coded(
                    "internal-missing-field",
                    "Missing prepared scalar array cache",
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
                        && cell["calculation"]["value"] != cell["cachedValue"]
                {
                    return Err(crate::coded(
                        "spill-stale",
                        "Singleton scalar-array prepared cache changed source or omitted fresh result",
                    ));
                }
            } else if prepared["formula"].as_str() != Some(canonical.as_str())
                || prepared["calculation"] != cell["calculation"]
            {
                return Err(crate::coded(
                    "internal-mismatch",
                    "Prepared scalar array formula/cache differs from accepted proof",
                ));
            }
            accepted.insert((sheet.clone(), *r, *c));
        }
        if page.len() < 64 {
            break;
        }
    }
    let mut statement=db.prepare_cached("SELECT sheet,r,c FROM export_changes WHERE json_extract(data,'$.calculation.scalarArray') IS NOT NULL OR json_extract(data,'$.cache.scalarArray') IS NOT NULL").map_err(err)?;
    let mut rows = statement.query([]).map_err(err)?;
    while let Some(row) = rows.next().map_err(err)? {
        if !accepted.contains(&(
            row.get::<_, String>(0).map_err(err)?,
            row.get::<_, u32>(1).map_err(err)?,
            row.get::<_, u32>(2).map_err(err)?,
        )) {
            return Err(crate::coded(
                "spill-bad-input",
                "Prepared scalar array cache has no accepted proof",
            ));
        }
    }
    Ok(())
}
