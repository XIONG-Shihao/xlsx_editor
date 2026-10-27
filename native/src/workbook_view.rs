//! Answers questions about a workbook package: its sheets, names, tables,
//! styles, filters, rules and what keeps it from being edited. The session layer
//! asks here instead of opening the package itself; it decides what to allow
//! from the answers. Reads the cell-free copy the index keeps (`metadata.xlsx`).
mod dom;
mod filter;
mod package;
mod rules;
mod styles;

use crate::{Result, coded, field, local};
use dom::{Dom, MAIN};
use package::Package;
use serde_json::{Map, Value, json};
use sha2::{Digest, Sha256};
use std::{collections::HashMap, path::Path};
use styles::{Issues, js_number, number_value};

const OFFICE_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
/// Elements whose presence on a sheet decides whether rows and columns can move.
const COUNTED: [&str; 14] = [
    "f",
    "mergeCell",
    "tableParts",
    "drawing",
    "conditionalFormatting",
    "dataValidations",
    "hyperlink",
    "legacyDrawingHF",
    "oleObjects",
    "controls",
    "autoFilter",
    "rowBreaks",
    "colBreaks",
    "legacyDrawing",
];

struct Relation {
    kind: String,
    path: Option<String>,
}

fn sha(bytes: &[u8]) -> String {
    format!("{:x}", Sha256::digest(bytes))
}

/// The workbook's relationships by id, strictly: nothing foreign, nested or repeated.
fn relationships(rels: &Dom) -> Result<HashMap<String, Relation>> {
    if !rels.root_is(PACKAGE_REL, "Relationships") {
        return Err(coded(
            "package-malformed",
            "Invalid workbook relationship namespace/root",
        ));
    }
    let mut total = 0;
    rels.walk(0, &mut |i| {
        if rels.nodes[i].ns == PACKAGE_REL && rels.nodes[i].local == "Relationship" {
            total += 1;
        }
    });
    let kids = rels.child_elements(0);
    let foreign = kids
        .iter()
        .any(|i| rels.nodes[*i].ns != PACKAGE_REL || rels.nodes[*i].local != "Relationship");
    if total != kids.len() || foreign || rels.has_loose_text(0) {
        return Err(coded(
            "package-unsupported",
            "Foreign or nested workbook relationship graph",
        ));
    }
    let mut targets = HashMap::new();
    for node in kids {
        let (id, kind) = (rels.attr(node, "Id"), rels.attr(node, "Type"));
        let (target, mode) = (rels.attr(node, "Target"), rels.attr(node, "TargetMode"));
        if id.is_empty()
            || kind.is_empty()
            || target.is_empty()
            || targets.contains_key(id)
            || !rels.child_elements(node).is_empty()
            || rels.has_loose_text(node)
        {
            return Err(coded(
                "internal-missing-field",
                "Missing, malformed or duplicate workbook relationship",
            ));
        }
        let path = if target.contains("..")
            || (!mode.is_empty() && !["Internal", "External"].contains(&mode))
            || mode == "External"
        {
            None
        } else if let Some(rooted) = target.strip_prefix('/') {
            Some(rooted.to_string())
        } else {
            Some(format!("xl/{target}"))
        };
        targets.insert(
            id.to_string(),
            Relation {
                kind: kind.to_string(),
                path,
            },
        );
    }
    Ok(targets)
}

fn inactive_protection(book: &Dom, node: usize) -> bool {
    book.child_elements(node).is_empty()
        && !book.has_loose_text(node)
        && book.nodes[node].attrs.iter().all(|a| {
            a.ns.is_empty()
                && ["lockStructure", "lockWindows", "lockRevision"].contains(&a.local.as_str())
                && ["0", "false"].contains(&a.value.as_str())
        })
}

/// The first of the sheet's views, when that is unambiguous: zoom and what is shown.
fn source_view(sheet: &Dom) -> Value {
    let containers = sheet.children(0, "sheetViews");
    if containers.len() != 1 {
        return json!({});
    }
    let views: Vec<usize> = sheet
        .children(containers[0], "sheetView")
        .into_iter()
        .filter(|v| sheet.attr(*v, "workbookViewId") == "0")
        .collect();
    if views.len() != 1 {
        return json!({});
    }
    let view = views[0];
    let mut out = Map::new();
    if let Some(zoom) = sheet.attribute(view, "zoomScale") {
        if !zoom.is_empty() && zoom.bytes().all(|b| b.is_ascii_digit()) {
            let n = zoom.parse::<u64>().unwrap_or(0);
            if (10..=400).contains(&n) {
                out.insert("zoom".into(), json!(n));
            }
        }
    }
    for (attribute, key) in [
        ("showGridLines", "gridlines"),
        ("showRowColHeaders", "headings"),
    ] {
        let raw = sheet.attr(view, attribute);
        if ["0", "1", "false", "true"].contains(&raw) {
            out.insert(key.into(), json!(raw == "1" || raw == "true"));
        }
    }
    Value::Object(out)
}

fn package_path(dir_of: &str, target: &str) -> Result<String> {
    let mut segments: Vec<&str> = dir_of.split('/').collect();
    segments.pop();
    if target.starts_with('/') {
        segments.clear();
    }
    for segment in target.split('/') {
        if segment == ".." {
            if segments.pop().is_none() {
                return Err(coded("package-malformed", "Invalid table relationship"));
            }
        } else if !segment.is_empty() && segment != "." {
            segments.push(segment);
        }
    }
    Ok(segments.join("/"))
}

struct Table {
    name: String,
    sheet: String,
    reference: String,
    path: String,
    header_rows: f64,
    totals_rows: f64,
    columns: Vec<String>,
    dom: Dom,
}

struct SheetFacts {
    value: Map<String, Value>,
    dom: Dom,
}

pub(crate) fn command(root: &Path, input: &Value) -> Result<Value> {
    let name = input["name"].as_str().unwrap_or("metadata.xlsx");
    let _ = field(input, "root")?;
    let mut package = Package::open(&local(root, name)?)?;
    if !package.has("xl/workbook.xml") || !package.has("xl/_rels/workbook.xml.rels") {
        return Err(coded("package-bad-input", "Not a supported XLSX workbook"));
    }
    let book = package.dom("xl/workbook.xml")?;
    let rels = package.dom("xl/_rels/workbook.xml.rels")?;
    if !book.root_is(MAIN, "workbook") {
        return Err(coded(
            "package-unsupported",
            "Unsupported workbook namespace/root",
        ));
    }
    let targets = relationships(&rels)?;
    let containers: Vec<usize> = book
        .child_elements(0)
        .into_iter()
        .filter(|i| book.nodes[*i].local == "sheets")
        .collect();
    if containers.len() != 1 || book.nodes[containers[0]].ns != MAIN {
        return Err(coded(
            "internal-missing-field",
            "Missing or foreign workbook sheet inventory",
        ));
    }
    let container = containers[0];
    let sheet_nodes = book.elements(0, "sheet");
    let foreign = book
        .child_elements(container)
        .iter()
        .any(|i| !book.is(*i, "sheet"));
    let nested = book.children(container, "sheet").len() != sheet_nodes.len();
    if sheet_nodes.is_empty() || nested || foreign {
        return Err(coded(
            "package-unsupported",
            "Foreign or nested workbook sheet inventory",
        ));
    }

    let mut reasons = package::read_only_reasons(&mut package);
    let macro_project = package::macro_project(&package);
    let protections = book.elements(0, "workbookProtection");
    if protections.len() > 1
        || protections.iter().any(|p| {
            book.child_elements(0).iter().all(|c| c != p) || !inactive_protection(&book, *p)
        })
    {
        reasons.push("Workbook protection is enabled".to_string());
    }

    // Theme and styles.
    let theme_paths: Vec<Option<String>> = targets
        .values()
        .filter(|r| r.kind == format!("{OFFICE_REL}/theme"))
        .map(|r| r.path.clone())
        .collect();
    let (theme, theme_issues) = styles::theme(&mut package, &theme_paths);
    let mut issues = Issues(Vec::new());
    for issue in &theme_issues {
        issues.add(issue);
    }
    let page_theme = match &theme {
        Some(theme) => {
            let bytes = package.bytes(&theme.part)?;
            let sample = (format_matches!(theme.part.as_str(), XL_THEME_PAGE_THEME_XML))
                .then(|| {
                    ["neutral", "forest", "plum"].into_iter().find(|id| {
                        crate::package_edit::page_theme_sample(id).is_some_and(|xml| {
                            crate::workbook_format::page_theme_matches(bytes.as_slice(), xml)
                        })
                    })
                })
                .flatten();
            json!({
                "revision": sha(&bytes), "sample": sample, "canApply": sample.is_some(),
                "protectedReason": if sample.is_some() { Value::Null } else { json!("Imported or modified themes stay protected.") },
            })
        }
        None if theme_paths.is_empty() => {
            json!({"revision": null, "sample": null, "canApply": true, "protectedReason": null})
        }
        None => json!({
            "revision": null, "sample": null, "canApply": false,
            "protectedReason": "Imported, ambiguous or modified themes stay protected.",
        }),
    };
    let styles_dom = if package.has("xl/styles.xml") {
        Some(package.dom("xl/styles.xml")?)
    } else {
        None
    };
    let read_styles = styles::read(styles_dom.as_ref(), &theme, &mut issues);
    let mut unverified: Vec<String> = Vec::new();
    let mut push = |text: &str| {
        if !unverified.iter().any(|t| t == text) {
            unverified.push(text.to_string());
        }
    };
    for issue in &issues.0 {
        push(issue);
    }

    let date1904 = book
        .first(0, "workbookPr")
        .is_some_and(|p| ["1", "true"].contains(&book.attr(p, "date1904")));

    // Sheets.
    let mut names: Vec<String> = Vec::new();
    let mut paths: Vec<String> = Vec::new();
    let mut protected_sheets = Vec::new();
    let mut sheets: Vec<SheetFacts> = Vec::new();
    for node in &sheet_nodes {
        let identities: Vec<&str> = book.nodes[*node]
            .attrs
            .iter()
            .filter(|a| a.ns == OFFICE_REL && a.local == "id")
            .map(|a| a.value.as_str())
            .collect();
        let relation = (identities.len() == 1)
            .then(|| targets.get(identities[0]))
            .flatten();
        let path = relation
            .filter(|r| r.kind == format!("{OFFICE_REL}/worksheet"))
            .and_then(|r| r.path.clone());
        let Some(path) = path.filter(|p| package.has(p)) else {
            return Err(coded(
                "package-refused",
                "Workbook contains a missing or external sheet",
            ));
        };
        let sheet_name = book.attribute(*node, "name").unwrap_or("").to_string();
        if sheet_name.is_empty()
            || names.contains(&sheet_name.to_lowercase())
            || paths.contains(&path)
        {
            return Err(coded(
                "package-refused",
                "Workbook contains duplicate or missing worksheet identities",
            ));
        }
        names.push(sheet_name.to_lowercase());
        paths.push(path.clone());
        let sheet = package.dom(&path)?;
        if !sheet.elements(0, "sheetProtection").is_empty() {
            protected_sheets.push(sheet_name.clone());
        }
        let special = sheet
            .elements(0, "f")
            .iter()
            .filter(|f| sheet.attribute(**f, "t").is_some())
            .count();
        if special > 0 {
            push(
                "Shared, array, or data-table formulas are preserved but unresolved; their cells cannot be overwritten",
            );
        }
        let extensions = sheet.elements(0, "extLst");
        if !extensions.is_empty() {
            push("Extended worksheet features are preserved but not interpreted");
        }
        let format = sheet.first(0, "sheetFormatPr");
        let mut geometry: Vec<String> = read_styles.normal_issues.clone();
        let mut dimension = |attribute: &str, maximum: f64, integer: bool| -> Option<Value> {
            let format = format?;
            let raw = sheet.attribute(format, attribute)?;
            let value = js_number(raw);
            if !rules::plain_number(raw.trim())
                || !value.is_finite()
                || value < 0.0
                || value > maximum
                || (integer && value.fract() != 0.0)
            {
                geometry.push(format!(
                    "Invalid worksheet {attribute}; its display default is unresolved"
                ));
                return None;
            }
            Some(number_value(value))
        };
        let mut sheet_format = Map::new();
        for (key, attribute, maximum, integer) in [
            ("defaultRowHeight", "defaultRowHeight", 409.5, false),
            ("defaultColumnWidth", "defaultColWidth", 255.0, false),
            ("baseColumnWidth", "baseColWidth", 255.0, true),
        ] {
            if let Some(value) = dimension(attribute, maximum, integer) {
                sheet_format.insert(key.into(), value);
            }
        }
        let hidden_default = format.map_or("", |f| sheet.attr(f, "zeroHeight"));
        if !hidden_default.is_empty() && !["0", "1", "false", "true"].contains(&hidden_default) {
            geometry
                .push("Invalid default row visibility; its display default is unresolved".into());
        }
        sheet_format.insert(
            "rowsHiddenByDefault".into(),
            json!(["1", "true"].contains(&hidden_default)),
        );
        let freeze = sheet.first(0, "pane").map(|p| {
            json!({
                "columns": number_value(js_number(sheet.attr(p, "xSplit"))),
                "rows": number_value(js_number(sheet.attr(p, "ySplit"))),
                "state": sheet.attr(p, "state"),
            })
        });
        let widths: Vec<Value> = sheet
            .elements(0, "col")
            .into_iter()
            .map(|c| {
                json!({
                    "min": number_value(js_number(sheet.attr(c, "min"))),
                    "max": number_value(js_number(sheet.attr(c, "max"))),
                    "width": number_value(js_number(sheet.attr(c, "width"))),
                    "hidden": sheet.attr(c, "hidden") == "1",
                })
            })
            .collect();
        let mut counts = Map::new();
        for element in COUNTED {
            let n = sheet.elements(0, element).len();
            if n > 0 {
                counts.insert(element.into(), json!(n));
            }
        }
        let pivot_link = {
            let slash = path.rfind('/').unwrap_or(0);
            let rels_path = format!("{}/_rels/{}.rels", &path[..slash], &path[slash + 1..]);
            package.has(&rels_path)
                && package
                    .bytes(&rels_path)?
                    .windows(12)
                    .any(|w| w == b"/pivotTable\"")
        };
        let mut value = Map::new();
        value.insert("name".into(), json!(book.attr(*node, "name")));
        value.insert("date1904".into(), json!(date1904));
        let state = book.attr(*node, "state");
        value.insert(
            "state".into(),
            json!(if state.is_empty() { "visible" } else { state }),
        );
        value.insert("path".into(), json!(path));
        value.insert("sheetFormat".into(), Value::Object(sheet_format));
        value.insert("normalFont".into(), read_styles.normal_font.clone());
        value.insert("sourceView".into(), source_view(&sheet));
        if !geometry.is_empty() {
            value.insert("geometryIssues".into(), json!(geometry));
        }
        value.insert(
            "merges".into(),
            json!(
                sheet
                    .elements(0, "mergeCell")
                    .into_iter()
                    .map(|m| sheet.attr(m, "ref").to_string())
                    .collect::<Vec<_>>()
            ),
        );
        value.insert(
            "conditionalRules".into(),
            json!(rules::conditional(&sheet, styles_dom.as_ref())),
        );
        value.insert("validationRules".into(), json!(rules::validation(&sheet)));
        if let Some(freeze) = freeze {
            value.insert("freeze".into(), freeze);
        }
        value.insert("widths".into(), json!(widths));
        value.insert("elements".into(), Value::Object(counts));
        value.insert("specialFormulas".into(), json!(special));
        value.insert("extLstCount".into(), json!(extensions.len()));
        value.insert(
            "extLstFilled".into(),
            json!(extensions.iter().any(|e| !sheet.nodes[*e].kids.is_empty())),
        );
        value.insert("pivotLink".into(), json!(pivot_link));
        sheets.push(SheetFacts { value, dom: sheet });
    }

    // Defined names.
    let named_ranges: Vec<Value> = book
        .elements(0, "definedName")
        .into_iter()
        .map(|n| {
            let mut item = json!({"name": book.attr(n, "name"), "formula": book.text(n)});
            if let Some(raw) = book.attribute(n, "localSheetId") {
                let index = (!raw.is_empty() && raw.bytes().all(|b| b.is_ascii_digit()))
                    .then(|| raw.parse::<u64>().ok())
                    .flatten()
                    .filter(|i| *i <= 4294967295)
                    .and_then(|i| sheets.get(i as usize));
                match index {
                    Some(s) => item["sheet"] = s.value["name"].clone(),
                    None => item["scopeInvalid"] = json!(true),
                }
            }
            item
        })
        .collect();

    // Tables, found through each sheet's relationships.
    let mut tables: Vec<Table> = Vec::new();
    for sheet in &sheets {
        let sheet_path = sheet.value["path"].as_str().unwrap_or("").to_string();
        let slash = sheet_path.rfind('/').unwrap_or(0);
        let rels_path = format!(
            "{}/_rels/{}.rels",
            &sheet_path[..slash],
            &sheet_path[slash + 1..]
        );
        if !package.has(&rels_path) {
            continue;
        }
        let relations = package.dom(&rels_path)?;
        let mut all = Vec::new();
        relations.walk(0, &mut |i| {
            if relations.nodes[i].local == "Relationship" {
                all.push(i);
            }
        });
        for relation in all {
            if !relations.attr(relation, "Type").ends_with("/table")
                || relations.attr(relation, "TargetMode") == "External"
            {
                continue;
            }
            let table_path = package_path(&sheet_path, relations.attr(relation, "Target"))?;
            if !package.has(&table_path) {
                return Err(coded(
                    "package-refused",
                    "Workbook contains a missing table",
                ));
            }
            let dom = package.dom(&table_path)?;
            let display = dom.attr(0, "displayName");
            let columns = dom
                .elements(0, "tableColumn")
                .into_iter()
                .map(|c| dom.attr(c, "name").to_string())
                .collect();
            let header = dom.attr(0, "headerRowCount");
            let totals = dom.attr(0, "totalsRowCount");
            tables.push(Table {
                name: if display.is_empty() {
                    dom.attr(0, "name")
                } else {
                    display
                }
                .to_string(),
                sheet: sheet.value["name"].as_str().unwrap_or("").to_string(),
                reference: dom.attr(0, "ref").to_string(),
                path: table_path,
                header_rows: if header.is_empty() {
                    1.0
                } else {
                    js_number(header)
                },
                totals_rows: if totals.is_empty() {
                    0.0
                } else {
                    js_number(totals)
                },
                columns,
                dom,
            });
        }
    }

    // Each sheet's table list and its filter.
    for sheet in &mut sheets {
        let own = sheet.value["name"].as_str().unwrap_or("").to_string();
        let mine: Vec<&Table> = tables.iter().filter(|t| t.sheet == own).collect();
        sheet.value.insert(
            "tableRanges".into(),
            json!(
                mine.iter()
                    .map(|t| json!({"name": t.name, "ref": t.reference}))
                    .collect::<Vec<_>>()
            ),
        );
        let mut found = Vec::new();
        for filter_node in sheet.dom.children(0, "autoFilter") {
            found.push(filter::inspect(&sheet.dom, filter_node));
        }
        for table in &mine {
            for filter_node in table.dom.children(0, "autoFilter") {
                found.push(filter::inspect(&table.dom, filter_node));
            }
        }
        let imported = sheet
            .dom
            .children(0, "sheetPr")
            .iter()
            .any(|p| ["1", "true"].contains(&sheet.dom.attr(*p, "filterMode")));
        if let Some(chosen) = filter::choose(found, imported) {
            sheet.value.insert("filter".into(), chosen);
        }
    }

    let names_of = |pattern: &dyn Fn(&str) -> bool| package.names.iter().any(|n| pattern(n));
    let features: Vec<String> = package
        .names
        .iter()
        .filter(|n| {
            ["charts/", "media/", "pivot", "embeddings/", "customXml/"]
                .iter()
                .any(|p| n.contains(p))
        })
        .cloned()
        .collect();
    let has_pivot_tables = names_of(&|n| n.contains("pivotTables/"));
    let linked_rows = names_of(&|n| {
        n.contains("externalLinks")
            || n.contains("pivotTables")
            || n.contains("connections.xml")
            || n.contains("slicer")
    });
    let linked_data = names_of(&|n| {
        n.contains("connections.xml") || n.contains("slicer") || n.contains("queryTables/")
    });
    let consolidation = if has_pivot_tables {
        let caches: Vec<String> = package
            .names
            .iter()
            .filter(|n| {
                n.contains("pivotCache/")
                    && n.ends_with(".xml")
                    && !n["pivotCache/".len()..].is_empty()
            })
            .filter(|n| {
                n.rsplit_once("pivotCache/")
                    .is_some_and(|(_, rest)| !rest.contains('/'))
            })
            .cloned()
            .collect();
        let mut found = false;
        for name in caches {
            let text = String::from_utf8_lossy(&package.bytes(&name)?).into_owned();
            found |= text.match_indices("consolidation").any(|(at, _)| {
                let before = &text[..at];
                let open = before.rfind('<');
                let tag_start_ok = open.is_some_and(|o| {
                    let prefix = &before[o + 1..];
                    prefix.is_empty()
                        || (prefix.ends_with(':')
                            && prefix[..prefix.len() - 1]
                                .bytes()
                                .all(|b| b.is_ascii_alphanumeric() || b == b'_'))
                });
                tag_start_ok
                    && text[at + "consolidation".len()..]
                        .starts_with(|c: char| c.is_whitespace() || c == '>')
            });
        }
        found
    } else {
        false
    };

    Ok(json!({
        "reasons": reasons,
        "macroProject": macro_project,
        "unverified": unverified,
        "protectedSheets": protected_sheets,
        "pageTheme": page_theme,
        "cellStyles": read_styles.cell_styles,
        "namedRanges": named_ranges,
        "tables": tables.iter().map(|t| json!({
            "name": t.name, "sheet": t.sheet, "ref": t.reference, "path": t.path,
            "headerRowCount": number_value(t.header_rows), "totalsRowCount": number_value(t.totals_rows),
            "columns": t.columns.iter().map(|c| json!({"name": c})).collect::<Vec<_>>(),
        })).collect::<Vec<_>>(),
        "features": features,
        "package": {
            "linkedRowReferences": linked_rows, "linkedData": linked_data,
            "pivotTables": has_pivot_tables, "pivotConsolidation": consolidation,
        },
        "sheets": sheets.into_iter().map(|s| Value::Object(s.value)).collect::<Vec<_>>(),
    }))
}
