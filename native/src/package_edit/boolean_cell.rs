//! Admission for the editor's literal Boolean-cell profile, not Excel control metadata.
use super::{print_settings::Tree, *};
pub(super) fn check(root: &Path, input: &Value) -> Result<Value> {
    let map = input
        .as_object()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid Boolean-cell request"))?;
    let required = [
        "action",
        "operation",
        "root",
        "version",
        "sourceRevision",
        "command",
    ];
    if required.iter().any(|k| !map.contains_key(*k))
        || map
            .keys()
            .any(|k| !required.contains(&k.as_str()) && k != "diagnostics")
        || input["action"] != "workbook"
        || input["operation"] != "boolean-cell-check"
        || input.get("diagnostics").is_some_and(|v| !v.is_boolean())
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid Boolean-cell request fields",
        ));
    }
    let command = &input["command"];
    let fields = command
        .as_object()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid Boolean-cell command"))?;
    let keys = ["type", "action", "sheet", "cell", "expectedValue"];
    if fields.len() != keys.len()
        || keys.iter().any(|k| !fields.contains_key(*k))
        || command["type"] != "booleanCell"
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid Boolean-cell command fields",
        ));
    }
    let action = field(command, "action")?;
    let expected = &command["expectedValue"];
    if !["insert", "toggle", "remove"].contains(&action)
        || !(expected.is_boolean() || expected.is_null())
        || action == "insert" && !expected.is_null()
        || action == "toggle" && !expected.is_boolean()
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid Boolean-cell action or expected value",
        ));
    }
    let reference = field(command, "cell")?;
    let at = crate::xml::position(reference)?;
    if crate::xml::address(at.0, at.1) != reference {
        return Err(crate::coded(
            "package-bad-input",
            "Boolean-cell requires one canonical cell",
        ));
    }
    let version = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid Boolean-cell version"))?;
    let db = crate::database(root)?;
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
            "Unknown Boolean-cell version",
        ));
    }
    let index: Value =
        serde_json::from_slice(&fs::read(local(root, "index.json")?).map_err(err)?).map_err(err)?;
    if index["revision"] != input["sourceRevision"]
        || crate::hash(&local(root, "source.xlsx")?)? != field(input, "sourceRevision")?
    {
        return Err(crate::coded(
            "package-stale",
            "Boolean-cell source revision changed",
        ));
    }
    let sheet = field(command, "sheet")?;
    let inventory = crate::used_range_transfer::worksheet_inventory(&local(root, "source.xlsx")?)?;
    let selected = inventory.iter().find(|s| s.name == sheet).ok_or_else(|| {
        crate::coded(
            "internal-unknown-reference",
            "Unknown Boolean-cell worksheet",
        )
    })?;
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "metadata.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    for name in zip.file_names() {
        let name = name.to_ascii_lowercase();
        if [
            "xl/activex/",
            "xl/ctrlprops/",
            "xl/featurepropertybag/",
            "xl/richdata/",
        ]
        .iter()
        .any(|prefix| name.starts_with(prefix))
            || name == "xl/metadata.xml"
        {
            return Err(crate::coded(
                "package-unsupported",
                "Foreign controls or cell metadata are outside the editor Boolean-cell profile; source is preserved",
            ));
        }
    }
    let tree = Tree::parse(&part_bytes(
        &mut zip,
        &selected.path,
        &serde_json::Map::new(),
    )?)?;
    tree.root(NS, "worksheet")?;
    for node in &tree.nodes {
        if [
            "sheetProtection",
            "controls",
            "control",
            "oleObjects",
            "oleObject",
            "legacyDrawing",
            "legacyDrawingHF",
            "cellControls",
            "cellControl",
            "AlternateContent",
            "extLst",
        ]
        .contains(&node.name.as_str())
        {
            return Err(crate::coded(
                "package-protected",
                "Foreign/protected worksheet control graph is outside the Boolean-cell profile",
            ));
        }
        if node.name == "mergeCell" {
            let (a, b) =
                crate::bulk::bounds(node.attrs.get("ref").ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing merged range")
                })?)?;
            if at.0 >= a.0 && at.0 <= b.0 && at.1 >= a.1 && at.1 <= b.1 {
                return Err(crate::coded(
                    "package-bad-input",
                    "Boolean-cell checkbox requires an unmerged cell",
                ));
            }
        }
    }
    let page = crate::range::read_public_db(
        &db,
        &json!({"version":version,"sheet":selected.path,"range":reference,"limit":2}),
    )?;
    let cells = page["cells"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid Boolean-cell range"))?;
    if cells.len() > 1 {
        return Err(crate::coded(
            "package-ambiguous",
            "Ambiguous Boolean-cell target",
        ));
    }
    let cell = cells.first().map(|p| p[1].clone()).unwrap_or(Value::Null);
    if [
        "formula",
        "formulaType",
        "spill",
        "formulaOutput",
        "calculation",
    ]
    .iter()
    .any(|key| !cell[*key].is_null())
    {
        return Err(crate::coded(
            "package-unsupported",
            "Formula or derived cells cannot become Boolean-cell checkboxes",
        ));
    }
    let value = &cell["value"];
    if !(value.is_null() || value.is_boolean())
        || !cell["dataType"].is_null()
            && cell["dataType"] != "blank"
            && cell["dataType"] != "boolean"
    {
        return Err(crate::coded(
            "package-bad-input",
            "Boolean-cell checkbox requires a blank or literal Boolean cell",
        ));
    }
    if value != expected {
        return Err(crate::coded(
            "package-stale",
            "Boolean-cell target changed since capture",
        ));
    }
    let result = match action {
        "insert" => json!(false),
        "toggle" => json!(
            !expected
                .as_bool()
                .ok_or_else(|| crate::coded("package-bad-input", "Toggle requires a Boolean"))?
        ),
        _ => Value::Null,
    };
    Ok(
        json!({"schema":workbook_format!(SCHEMA_BOOLEAN_CELL_ADMISSION_V1),"sheet":sheet,"cell":reference,"value":result,"changed":result!=*value}),
    )
}
