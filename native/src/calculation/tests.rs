use super::*;
use crate::spill::{Geometry, Group, Sequence};

struct Fixture {
    root: tempfile::TempDir,
    view: Value,
    group: Group,
}
impl Fixture {
    fn new(qualified: bool, blocked: bool) -> Self {
        let root = tempfile::tempdir().unwrap();
        let db = crate::database(root.path()).unwrap();
        db.execute_batch("CREATE TABLE versions(id INTEGER PRIMARY KEY,parent INTEGER); INSERT INTO versions VALUES(0,NULL),(1,0); CREATE TABLE changes(version INTEGER,sheet TEXT,r INTEGER,c INTEGER,data TEXT,edit TEXT,PRIMARY KEY(version,sheet,r,c)); CREATE INDEX changes_coordinates ON changes(sheet,r,c,version); CREATE TABLE strings(id INTEGER PRIMARY KEY,value TEXT); CREATE TABLE cells(sheet TEXT,r INTEGER,c INTEGER,kind TEXT,raw TEXT,inline TEXT,formula TEXT,formula_attrs TEXT,style INTEGER,PRIMARY KEY(sheet,r,c)); CREATE INDEX cells_formulas ON cells(sheet,r,c) WHERE formula IS NOT NULL;").unwrap();
        crate::used_range::schema(&db).unwrap();
        crate::spill::schema(&db).unwrap();
        let geometry = Geometry::from_reference(if blocked { "A1" } else { "A1:B2" }).unwrap();
        let mut proof = json!({"constructor":"sequence-integer-v1","arguments":[2.0,2.0,1.0,1.0],"status":if blocked{"blocked"}else{"calculated"}});
        if qualified {
            proof["qualification"] = json!("retained-sequence-v1");
        }
        if blocked {
            proof["intendedRange"] = json!("A1:B2");
            proof["errorType"] = json!(8);
            proof["subType"] = json!(1);
            proof["blocker"] = json!("B2");
            db.execute(
                "INSERT INTO changes VALUES(0,'sheet',1,1,?1,'{}')",
                [json!({"value":99.0,"dataType":"number"}).to_string()],
            )
            .unwrap();
        }
        let group = Group {
            version: 1,
            sheet: "sheet".into(),
            master: (0, 0),
            geometry,
            formula: "_xlfn.SEQUENCE(SUM(G1:G2),2)".into(),
            proof,
        };
        let sequence = Sequence::new(group.master, &[2.0, 2.0]).unwrap();
        db.execute_batch("BEGIN").unwrap();
        for (index, at) in group.geometry.positions().enumerate() {
            let value = if blocked {
                json!("#SPILL!")
            } else {
                sequence.values[index].clone()
            };
            let mut cell = json!({"value":value,"dataType":if blocked{"error"}else{"number"},"rawValue":if blocked{"#VALUE!".to_owned()}else{value.to_string()},"styleId":7,"spill":group.cell_metadata(at).unwrap(),"calculation":if blocked{json!({"status":"error","error":"#SPILL!","revision":"accepted"})}else{json!({"status":"calculated","value":value,"revision":"accepted"})}});
            if at == group.master {
                cell["formula"] = json!(group.formula);
                cell["formulaType"] = json!("array");
                cell["formulaReference"] = json!(group.geometry.reference());
                if blocked {
                    cell["error"] = json!("#SPILL!");
                }
            }
            db.execute(
                "INSERT INTO changes VALUES(1,'sheet',?1,?2,?3,?4)",
                params![
                    at.0,
                    at.1,
                    cell.to_string(),
                    json!({"bold":true}).to_string()
                ],
            )
            .unwrap();
        }
        crate::spill::record(&db, &group, true).unwrap();
        db.execute_batch("COMMIT").unwrap();
        for reference in ["G1", "G2"] {
            let (r, c) = position(reference).unwrap();
            db.execute(
                "INSERT INTO changes VALUES(0,'sheet',?1,?2,?3,'{}')",
                params![
                    r,
                    c,
                    json!({"value":1.0,"dataType":"number","rawValue":"1"}).to_string()
                ],
            )
            .unwrap();
        }
        Self {
            root,
            view: json!({"revision":"accepted","sheets":[{"name":"Sheet","path":"sheet","rows":10,"columns":10,"cells":{}}],"namedRanges":[],"tables":[]}),
            group,
        }
    }
    fn db(&self) -> Connection {
        crate::database(self.root.path()).unwrap()
    }
    fn formula(&self, reference: &str, formula: &str) {
        let (r, c) = position(reference).unwrap();
        self.db().execute("INSERT INTO changes VALUES(1,'sheet',?1,?2,?3,'{}')", params![r,c,json!({"value":null,"dataType":"number","formula":formula,"rawValue":null,"styleId":0}).to_string()]).unwrap();
    }
    fn input(&self, version: u64) -> Value {
        json!({"version":version,"revision":"accepted","metadata":"native-owned-spill-tests","view":self.view})
    }
    fn read(&self, version: u64, range: &str) -> Value {
        crate::range::read_db(
            &self.db(),
            &json!({"version":version,"sheet":"sheet","range":range}),
        )
        .unwrap()
    }
    fn assert_preserved(&self, accepted: u64) {
        assert_eq!(
            self.read(1, &self.group.geometry.reference()),
            self.read(accepted, &self.group.geometry.reference())
        );
        let db = self.db();
        let changed:u64=db.query_row("SELECT count(*) FROM changes WHERE version=?1 AND sheet='sheet' AND r<=?2 AND c<=?3",params![accepted,self.group.geometry.last.0,self.group.geometry.last.1],|row|row.get(0)).unwrap();
        assert_eq!(
            changed, 0,
            "Scalar publication must inherit spill records unchanged"
        );
        let owned = crate::spill::effective(&db, accepted, "sheet").unwrap();
        assert_eq!(owned.len(), 1);
        assert_eq!(owned[0].version, 1);
    }
}
fn value(page: &Value, reference: &str) -> Value {
    page["cells"]
        .as_array()
        .unwrap()
        .iter()
        .find(|pair| pair[0] == reference)
        .unwrap()[1]["calculation"]["value"]
        .clone()
}

#[test]
fn owned_spills_are_literal_predecessors_without_scalar_publication() {
    let fixture = Fixture::new(true, false);
    fixture.formula("D1", "A1+B2");
    fixture.formula("D2", "SUM(A1:B2)");
    fixture.formula("D3", "COUNT(A1:B2)");
    let reply = calculate(&mut fixture.db(), &fixture.input(1)).unwrap();
    assert_eq!(reply["evaluatedFormulaCells"], 3);
    let accepted = reply["version"].as_u64().unwrap();
    let page = fixture.read(accepted, "D1:D3");
    assert_eq!(value(&page, "D1"), json!(5.0));
    assert_eq!(value(&page, "D2"), json!(10.0));
    assert_eq!(value(&page, "D3"), json!(4.0));
    fixture.assert_preserved(accepted);
    let edges:u64=fixture.db().query_row("SELECT count(*) FROM dependency_ranges WHERE owner_ref='A1' AND sheet='Sheet' AND r0=0 AND r1=1 AND c0=6 AND c1=6",[],|row|row.get(0)).unwrap();
    assert_eq!(
        edges, 1,
        "Constructor range dependencies must survive scalar planning"
    );
}

#[test]
fn partial_and_selected_plans_leave_owned_spills_with_their_owner() {
    let fixture = Fixture::new(true, false);
    fixture.formula("D1", "SUM(A1:B2)");
    let mut partial = fixture.input(1);
    partial["partial"] = json!(true);
    let reply = calculate(&mut fixture.db(), &partial).unwrap();
    assert_eq!(reply["evaluatedFormulaCells"], 1);
    assert_eq!(reply["unsupportedFormulaCells"], 0);
    fixture.assert_preserved(reply["version"].as_u64().unwrap());
    let reply = calculate(&mut fixture.db(), &fixture.input(1)).unwrap();
    let version = reply["version"].as_u64().unwrap();
    let db = fixture.db();
    let plan: u64 = db
        .query_row(
            "SELECT plan FROM dependency_bindings WHERE version=?1",
            [version],
            |row| row.get(0),
        )
        .unwrap();
    db.execute(
        "INSERT INTO dependency_selected VALUES(?1,'Sheet','A1'),(?1,'Sheet','D1')",
        [plan],
    )
    .unwrap();
    let mut input = fixture.input(version);
    input["plan"] = json!(plan);
    let selected = calculate(&mut fixture.db(), &input).unwrap();
    assert_eq!(selected["evaluatedFormulaCells"], 1);
    fixture.assert_preserved(selected["version"].as_u64().unwrap());
}

#[test]
fn unqualified_arrays_and_caller_groups_do_not_become_literal_proofs() {
    let fixture = Fixture::new(false, false);
    fixture.formula("D1", "A1+B2");
    let mut input = fixture.input(1);
    input["view"]["spillGroups"] =
        json!([{"sheet":"sheet","master":"A1","range":"A1:B2","status":"calculated"}]);
    let reason = calculate(&mut fixture.db(), &input).unwrap_err();
    assert!(reason.contains("Special formula"), "{reason}");
    assert_eq!(
        fixture
            .db()
            .query_row("SELECT count(*) FROM versions", [], |row| row
                .get::<_, u64>(0))
            .unwrap(),
        2
    );
}

#[test]
fn corrupt_owned_outputs_refuse_calculation_before_publication() {
    let fixture = Fixture::new(true, false);
    fixture.formula("D1", "A1+B2");
    fixture
        .db()
        .execute(
            "UPDATE changes SET data=json_set(data,'$.value',99) WHERE version=1 AND r=1 AND c=1",
            [],
        )
        .unwrap();
    let reason = calculate(&mut fixture.db(), &fixture.input(1)).unwrap_err();
    assert!(reason.contains("ownership proof"), "{reason}");
    assert_eq!(
        fixture
            .db()
            .query_row("SELECT count(*) FROM versions", [], |row| row
                .get::<_, u64>(0))
            .unwrap(),
        2
    );
}

#[test]
fn owned_spills_do_not_hide_remaining_scalar_compatibility_misses() {
    let fixture = Fixture::new(true, false);
    fixture.formula("D1", "A1+B2");
    fixture.formula("D2", "UNQUALIFIED_FUNCTION(1)");
    let reason = calculate(&mut fixture.db(), &fixture.input(1)).unwrap_err();
    assert!(
        crate::sentence(&reason).starts_with("Unsupported native calculation:"),
        "{reason}"
    );
    let mut input = fixture.input(1);
    input["partial"] = json!(true);
    let reply = calculate(&mut fixture.db(), &input).unwrap();
    assert_eq!(reply["evaluatedFormulaCells"], 1);
    assert_eq!(reply["unsupportedFormulaCells"], 1);
    let accepted = reply["version"].as_u64().unwrap();
    assert_eq!(value(&fixture.read(accepted, "D1"), "D1"), json!(5.0));
    assert!(
        fixture.read(accepted, "D2")["cells"][0][1]
            .get("calculation")
            .is_none()
    );
    fixture.assert_preserved(accepted);
}

#[test]
fn blocked_owned_spills_keep_their_error_and_refuse_numeric_consumers() {
    let fixture = Fixture::new(true, true);
    fixture.formula("D1", "1+1");
    let reply = calculate(&mut fixture.db(), &fixture.input(1)).unwrap();
    fixture.assert_preserved(reply["version"].as_u64().unwrap());
    fixture.formula("D2", "A1+1");
    let reason = calculate(&mut fixture.db(), &fixture.input(1)).unwrap_err();
    assert!(
        crate::sentence(&reason).starts_with("Unsupported native calculation:"),
        "{reason}"
    );
    assert_eq!(
        fixture.read(1, "A1")["cells"][0][1]["calculation"]["error"],
        "#SPILL!"
    );
}
