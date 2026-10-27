//! Plain owned legacy notes: real SpreadsheetML comments and fixed hidden VML
//! note boxes. The sealed identity rewrite never changes worksheet data/caches.
use super::{clear_hyperlinks, print_settings::Tree, theme_authoring, *};
const COMMENTS_EDGE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const VML_EDGE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
const COMMENTS_MIME: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
const VML_MIME: &str = "application/vnd.openxmlformats-officedocument.vmlDrawing";
const MAX_NOTES: usize = 256;
#[derive(Clone, Debug, PartialEq, Eq)]
struct Note {
    cell: String,
    author: String,
    text: String,
}
fn digest(bytes: &[u8]) -> String {
    format!("{:x}", Sha256::digest(bytes))
}
fn escape(v: &str) -> String {
    v.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\r', "&#13;")
        .replace('\n', "&#10;")
        .replace('\t', "&#9;")
}
fn text(v: &str, max: usize) -> Result<()> {
    if v.encode_utf16().count() > max
        || v.trim().is_empty()
        || v.chars().any(|c| {
            c < ' ' && !['\r', '\n', '\t'].contains(&c) || ['\u{fffe}', '\u{ffff}'].contains(&c)
        })
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid plain note text/author",
        ));
    }
    Ok(())
}
fn note(cell: &str, author: &str, body: &str) -> Result<Note> {
    let p = position(cell)?;
    if crate::xml::address(p.0, p.1) != cell {
        return Err(crate::coded(
            "note-bad-input",
            "Select a canonical single note cell",
        ));
    }
    text(author, 128)?;
    if author.to_ascii_lowercase().contains("tc=") {
        return Err(crate::coded(
            "note-unsupported",
            "Threaded-comment author tokens are reserved",
        ));
    }
    if author.contains(['\r', '\n', '\t']) {
        return Err(crate::coded(
            "note-bad-input",
            "Note author must be a single-line label",
        ));
    }
    text(body, 4096)?;
    Ok(Note {
        cell: cell.into(),
        author: author.into(),
        text: body.into(),
    })
}
fn parts(sheet: &str) -> (String, String) {
    let id = digest(sheet.as_bytes());
    (
        format!(workbook_format!(XL_COMMENTS_NOTES_XML), &id[..16]),
        format!(workbook_format!(XL_DRAWINGS_NOTES_VML), &id[..16]),
    )
}
fn rel_part(sheet: &str) -> Result<String> {
    let (dir, file) = sheet
        .rsplit_once('/')
        .ok_or_else(|| crate::coded("note-malformed", "Invalid note worksheet path"))?;
    Ok(format!("{dir}/_rels/{file}.rels"))
}
fn part(zip: &mut ZipArchive<crate::cancel::Reader<File>>, name: &str) -> Result<Vec<u8>> {
    part_bytes(zip, name, &serde_json::Map::new())
}
fn relation_tree(zip: &mut ZipArchive<crate::cancel::Reader<File>>, name: &str) -> Result<Tree> {
    if zip.file_names().any(|p| p == name) {
        Tree::parse(&part(zip, name)?)
    } else {
        Tree::parse(
            format!(
                "<Relationships xmlns=\"{}\"/>",
                String::from_utf8_lossy(PACKAGE_REL_NS)
            )
            .as_bytes(),
        )
    }
}
fn comments(notes: &[Note]) -> String {
    comments_xml(notes, false)
}
// Older editor-created records used this optional attribute. Accept their exact
// canonical form for migration, but never emit it: Excel repairs these records
// and discards their notes when the attribute is present in the tested profile.
fn comments_xml(notes: &[Note], legacy_shape_ids: bool) -> String {
    let mut authors = Vec::new();
    for n in notes {
        if !authors.contains(&n.author) {
            authors.push(n.author.clone());
        }
    }
    let a = authors
        .iter()
        .map(|a| format!("<author xml:space=\"preserve\">{}</author>", escape(a)))
        .collect::<String>();
    let c=notes.iter().enumerate().map(|(i,n)|{let shape_id=if legacy_shape_ids{format!(" shapeId=\"{}\"",1025+i)}else{String::new()};format!("<comment ref=\"{}\" authorId=\"{}\"{shape_id}><text><t xml:space=\"preserve\">{}</t></text></comment>",n.cell,authors.iter().position(|a|a==&n.author).unwrap(),escape(&n.text))}).collect::<String>();
    format!(
        "<comments xmlns=\"{}\"><authors>{a}</authors><commentList>{c}</commentList></comments>",
        String::from_utf8_lossy(NS)
    )
}
fn vml(notes: &[Note]) -> Result<String> {
    let mut xml = String::from(
        r##"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>"##,
    );
    for (i, n) in notes.iter().enumerate() {
        let (row, col) = position(&n.cell)?;
        let left = (col + 1).min(16381);
        let top = row.min(1048571);
        xml.push_str(&format!(r##"<v:shape id="_x0000_s{}" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:144pt;height:79pt;z-index:{};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>{left},15,{top},2,{},15,{},2</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>{row}</x:Row><x:Column>{col}</x:Column></x:ClientData></v:shape>"##,1025+i,i+1,left+2,top+4));
    }
    xml.push_str("</xml>");
    Ok(xml)
}
fn parse_comments(bytes: &[u8]) -> Result<Vec<Note>> {
    let tree = Tree::parse(bytes)?;
    tree.root(NS, "comments")?;
    theme_authoring::attributes(&tree, 0, &[])?;
    let a = tree
        .direct(0, "authors")?
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing note authors"))?;
    let c = tree
        .direct(0, "commentList")?
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing note list"))?;
    if tree.nodes[0].children != vec![a, c] || tree.nodes[c].children.len() > MAX_NOTES {
        return Err(crate::coded(
            "note-limit",
            "Extended/oversized legacy notes stay protected",
        ));
    }
    let authors = tree.nodes[a]
        .children
        .iter()
        .map(|i| {
            let n = &tree.nodes[*i];
            if n.ns != NS || n.name != "author" || !n.children.is_empty() {
                return Err(crate::coded(
                    "note-protected",
                    "Imported/rich note authors stay protected",
                ));
            }
            text(&n.text, 128)?;
            Ok(n.text.clone())
        })
        .collect::<Result<Vec<_>>>()?;
    let mut result = Vec::new();
    let mut prior = None;
    for i in &tree.nodes[c].children {
        let n = &tree.nodes[*i];
        if n.ns != NS || n.name != "comment" {
            return Err(crate::coded("note-unsupported", "Foreign note record"));
        }
        let id = n
            .attrs
            .get("authorId")
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing note author index"))?
            .parse::<usize>()
            .map_err(err)?;
        let author = authors
            .get(id)
            .ok_or_else(|| crate::coded("note-malformed", "Invalid note author index"))?;
        let t = tree
            .direct(*i, "text")?
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing note text"))?;
        let body = tree
            .direct(t, "t")?
            .ok_or_else(|| crate::coded("note-protected", "Rich note text stays protected"))?;
        if n.children != vec![t]
            || tree.nodes[t].children != vec![body]
            || !tree.nodes[body].children.is_empty()
        {
            return Err(crate::coded(
                "note-protected",
                "Extended note text stays protected",
            ));
        }
        let value = note(
            n.attrs
                .get("ref")
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing note cell"))?,
            author,
            &tree.nodes[body].text,
        )?;
        let pos = position(&value.cell)?;
        if prior.is_some_and(|p| pos <= p) {
            return Err(crate::coded(
                "note-bad-input",
                "Notes require sorted unique cells",
            ));
        }
        prior = Some(pos);
        result.push(value);
    }
    // These are this editor's one declared plain-note format, not an importer for
    // arbitrary VML/rich text. Real standard comments remain portable; unknowns are preserved.
    if result.is_empty()
        || (comments(&result).as_bytes() != bytes
            && comments_xml(&result, true).as_bytes() != bytes)
    {
        return Err(crate::coded(
            "note-protected",
            "Imported/extended note records stay protected",
        ));
    }
    Ok(result)
}
struct Graph {
    notes: Vec<Note>,
    revision: Option<String>,
    edges: Vec<usize>,
    read_only: Option<String>,
    comments: String,
    vml: String,
}
fn graph(
    zip: &mut ZipArchive<crate::cancel::Reader<File>>,
    sheet: &str,
    tree: &Tree,
    rels: &Tree,
) -> Result<Graph> {
    tree.root(NS, "worksheet")?;
    if tree.nodes.iter().any(|n| {
        n.ns == NS
            && ["extLst", "controls", "oleObjects", "legacyDrawingHF"].contains(&n.name.as_str())
    }) {
        return Err(crate::coded(
            "note-protected",
            "Protected or extended worksheet notes stay read-only",
        ));
    }
    let (comment_part, vml_part) = parts(sheet);
    let edges = theme_authoring::relationships(rels)?;
    let comments_edge = edges
        .iter()
        .filter(|i| rels.nodes[**i].attrs["Type"].ends_with("/comments"))
        .copied()
        .collect::<Vec<_>>();
    let vml_edge = edges
        .iter()
        .filter(|i| rels.nodes[**i].attrs["Type"].ends_with("/vmlDrawing"))
        .copied()
        .collect::<Vec<_>>();
    if edges
        .iter()
        .any(|i| rels.nodes[*i].attrs["Type"].ends_with("/threadedComment"))
    {
        return Err(crate::coded(
            "note-protected",
            "Threaded comments and their compatibility records stay protected",
        ));
    }
    let legacy = tree.direct(0, "legacyDrawing")?;
    let (_, source_legacy, _) =
        clear_hyperlinks::worksheet_tag_hash(zip, sheet, None, b"legacyDrawing")?;
    let names = clear_hyperlinks::package_names(zip)?;
    let workbook = Tree::parse(&part(zip, "xl/workbook.xml")?)?;
    workbook.root(NS, "workbook")?;
    let read_only = if tree
        .nodes
        .iter()
        .any(|n| n.ns == NS && n.name == "sheetProtection")
        // An element with no attributes locks nothing; some writers emit it empty.
        || workbook.nodes.iter().any(|n| {
            n.ns == NS
                && n.name == "workbookProtection"
                && n.attrs.keys().any(|key| key != "xmlns" && !key.starts_with("xmlns:"))
        })
        || names.iter().any(|n| {
            n.starts_with("_xmlsignatures/")
                || ["vbaproject.bin", "vbaproject.xml"]
                    .iter()
                    .any(|s| n.to_ascii_lowercase().ends_with(s))
        }) {
        Some("Protected, signed or macro note sources stay read-only".into())
    } else {
        None
    };
    if comments_edge.is_empty() && vml_edge.is_empty() && legacy.is_none() {
        if !source_legacy.is_empty() {
            return Err(crate::coded(
                "note-protected",
                "Unrecognized source legacy drawing stays protected",
            ));
        }
        if names.contains(&comment_part)
            || names.contains(&vml_part)
        {
            return Err(crate::coded(
                "note-protected",
                "Orphaned private note parts stay protected",
            ));
        }
        return Ok(Graph {
            notes: vec![],
            revision: None,
            edges: vec![],
            read_only,
            comments: comment_part,
            vml: vml_part,
        });
    }
    if comments_edge.len() != 1 || vml_edge.len() != 1 || legacy.is_none() {
        return Err(crate::coded(
            "note-protected",
            "Imported/incomplete note graphs stay protected",
        ));
    }
    let own = [comments_edge[0], vml_edge[0]];
    for (i, kind, target, id) in [
        (
            own[0],
            COMMENTS_EDGE,
            &comment_part,
            workbook_format!(NOTES),
        ),
        (own[1], VML_EDGE, &vml_part, workbook_format!(NOTESVML)),
    ] {
        let e = &rels.nodes[i];
        if e.attrs["Type"] != kind
            || !(e.attrs["Id"] == id
                || (id == workbook_format!(NOTES) && format_matches!(e.attrs["Id"], NOTES))
                || (id == workbook_format!(NOTESVML) && format_matches!(e.attrs["Id"], NOTESVML)))
            || e.attrs.get("TargetMode").is_some_and(|m| m != "Internal")
            || theme_authoring::target(sheet, &e.attrs["Target"])? != *target
        {
            return Err(crate::coded(
                "note-protected",
                "Imported/external note graphs stay protected",
            ));
        }
    }
    let actual = Tree::parse(&source_legacy)?;
    actual.root(NS, "legacyDrawing")?;
    actual.clean(0, &["r:id"])?;
    if !actual.nodes[0]
        .attrs
        .get("r:id")
        .is_some_and(|id| format_matches!(id.as_str(), NOTESVML))
    {
        return Err(crate::coded(
            "note-unsupported",
            "Source note VML has a different owner",
        ));
    }
    let l = legacy.unwrap();
    tree.clean(l, &["r:id"])?;
    if !tree.nodes[l]
        .attrs
        .get("r:id")
        .is_some_and(|id| format_matches!(id.as_str(), NOTESVML))
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Note VML ownership differs from the worksheet",
        ));
    }
    let (_, _, uses) = clear_hyperlinks::worksheet_tag_hash(
        zip,
        sheet,
        Some(&BTreeSet::from([
            workbook_format!(NOTES).into(),
            workbook_format!(NOTESVML).into(),
        ])),
        b"legacyDrawing",
    )?;
    if !uses.is_empty() {
        return Err(crate::coded(
            "note-refused",
            "Note edges have other worksheet consumers",
        ));
    }
    let mut incoming = BTreeMap::from([(comment_part.clone(), 0usize), (vml_part.clone(), 0usize)]);
    for name in names.iter().filter(|p| p.ends_with(".rels")) {
        let rt = Tree::parse(&part(zip, name)?)?;
        let owner = if name == "_rels/.rels" {
            String::new()
        } else {
            let (dir, file) = name
                .rsplit_once("/_rels/")
                .ok_or_else(|| crate::coded("note-malformed", "Invalid note relationship part"))?;
            format!(
                "{dir}/{}",
                file.strip_suffix(".rels").ok_or_else(|| crate::coded(
                    "note-malformed",
                    "Invalid note relationship suffix"
                ))?
            )
        };
        for i in theme_authoring::relationships(&rt)? {
            let e = &rt.nodes[i];
            if e.attrs.get("TargetMode").is_some_and(|m| m == "External") {
                continue;
            }
            let to = theme_authoring::target(&owner, &e.attrs["Target"])?;
            if let Some(count) = incoming.get_mut(&to) {
                *count += 1;
            }
        }
    }
    if incoming.values().any(|c| *c != 1)
        || names.contains(&rel_part(&comment_part)?)
        || names.contains(&rel_part(&vml_part)?)
    {
        return Err(crate::coded(
            "note-protected",
            "Shared or linked note parts stay protected",
        ));
    }
    let types = Tree::parse(&part(zip, "[Content_Types].xml")?)?;
    types.root(TYPE_NS, "Types")?;
    for (target, mime) in [(&comment_part, COMMENTS_MIME), (&vml_part, VML_MIME)] {
        let matches = types.nodes[0]
            .children
            .iter()
            .filter(|i| types.nodes[**i].attrs.get("PartName") == Some(&format!("/{target}")))
            .copied()
            .collect::<Vec<_>>();
        if matches.len() != 1 {
            return Err(crate::coded(
                "note-ambiguous",
                "Ambiguous note content type",
            ));
        }
        let n = &types.nodes[matches[0]];
        types.clean(matches[0], &["PartName", "ContentType"])?;
        if n.ns != TYPE_NS
            || n.name != "Override"
            || n.attrs.get("ContentType").map(String::as_str) != Some(mime)
        {
            return Err(crate::coded(
                "note-unsupported",
                "Foreign note content type",
            ));
        }
    }
    let bytes = part(zip, &comment_part)?;
    let notes = parse_comments(&bytes)?;
    let draw = part(zip, &vml_part)?;
    if draw != vml(&notes)?.as_bytes() {
        return Err(crate::coded(
            "note-protected",
            "Modified/foreign VML notes stay protected",
        ));
    }
    let revision = digest(format!("{sheet}\0{}\0{}", digest(&bytes), digest(&draw)).as_bytes());
    Ok(Graph {
        notes,
        revision: Some(revision),
        edges: own.to_vec(),
        read_only,
        comments: comment_part,
        vml: vml_part,
    })
}
fn definitions(root: &Path, source: &str, sheet_name: &str) -> Result<(String, Tree, Tree, Graph)> {
    let inventory = crate::used_range_transfer::worksheet_inventory(&local(root, source)?)?;
    let sheet = inventory
        .iter()
        .find(|s| s.name == sheet_name)
        .ok_or_else(|| crate::coded("internal-unknown-reference", "Unknown note worksheet"))?
        .path
        .clone();
    let mut metadata = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "metadata.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    let tree = Tree::parse(&part(&mut metadata, &sheet)?)?;
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, source)?).map_err(err)?,
    ))
    .map_err(err)?;
    let rels = relation_tree(&mut zip, &rel_part(&sheet)?)?;
    let g = graph(&mut zip, &sheet, &tree, &rels)?;
    Ok((sheet, tree, rels, g))
}
pub(super) struct Prepared {
    pub plan: Value,
    pub command: Value,
    pub changed: bool,
    pub prior: Option<String>,
    pub remaining: usize,
}
impl Prepared {
    pub fn result(&self) -> Value {
        json!({"schema":workbook_format!(SCHEMA_NOTE_WRITER_V1),"file":"package.xlsx","contentChanged":false,"changed":self.changed,"priorNoteRevision":self.prior,"remainingNotes":self.remaining})
    }
}
fn prepare_from(root: &Path, source: &str, command: &Value) -> Result<Prepared> {
    let remove = command["type"] == "removeNote";
    theme_authoring::fields(
        command,
        &["type", "sheet", "cell", "expectedNoteRevision"],
        if remove { &[] } else { &["note"] },
    )?;
    if !remove && (command["type"] != "setNote" || command.get("note").is_none()) {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid note command",
        ));
    }
    let cell = field(command, "cell")?;
    let pos = position(cell)?;
    if crate::xml::address(pos.0, pos.1) != cell {
        return Err(crate::coded(
            "note-bad-input",
            "Select a canonical note cell",
        ));
    }
    let next = if remove {
        None
    } else {
        theme_authoring::fields(&command["note"], &["author", "text"], &[])?;
        Some(note(
            cell,
            field(&command["note"], "author")?,
            field(&command["note"], "text")?,
        )?)
    };
    let expected = &command["expectedNoteRevision"];
    if !expected.is_null()
        && !expected.as_str().is_some_and(|v| {
            v.len() == 64
                && v.bytes()
                    .all(|b| b.is_ascii_digit() || matches!(b, b'a'..=b'f'))
        })
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid prior note revision",
        ));
    }
    let (sheet, tree, rels, g) = definitions(root, source, field(command, "sheet")?)?;
    if let Some(reason) = &g.read_only {
        return Err(crate::coded("note-refused", reason));
    }
    if *expected != json!(g.revision) {
        return Err(crate::coded(
            "note-stale",
            "Notes changed since the captured revision",
        ));
    }
    for n in &tree.nodes {
        if n.ns == NS && n.name == "mergeCell" {
            let value = n
                .attrs
                .get("ref")
                .ok_or_else(|| crate::coded("note-malformed", "Invalid merge target"))?;
            let bounds = crate::bulk::bounds(value)?;
            if pos.0 >= bounds.0.0
                && pos.0 <= bounds.1.0
                && pos.1 >= bounds.0.1
                && pos.1 <= bounds.1.1
                && pos != bounds.0
            {
                return Err(crate::coded(
                    "note-bad-input",
                    "Select the top-left cell of a merged note target",
                ));
            }
        }
    }
    let mut notes = g.notes.clone();
    if let Some(i) = notes.iter().position(|n| n.cell == cell) {
        if let Some(next) = next {
            notes[i] = next;
        } else {
            notes.remove(i);
        }
    } else if let Some(next) = next {
        notes.push(next);
    }
    notes.sort_by_key(|n| position(&n.cell).unwrap());
    if notes.len() > MAX_NOTES {
        return Err(crate::coded(
            "note-limit",
            "The owned sheet note limit is 256",
        ));
    }
    let empty =
        json!({"kind":command["type"],"replacements":{},"fragments":{},"copies":{},"removed":[]});
    if notes == g.notes {
        return Ok(Prepared {
            plan: empty,
            command: command.clone(),
            changed: false,
            prior: g.revision,
            remaining: notes.len(),
        });
    }
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, source)?).map_err(err)?,
    ))
    .map_err(err)?;
    let mut replacements = serde_json::Map::new();
    let mut removed = Vec::new();
    let types = Tree::parse(&part(&mut zip, "[Content_Types].xml")?)?;
    types.root(TYPE_NS, "Types")?;
    let type_indexes = types.nodes[0]
        .children
        .iter()
        .filter(|i| {
            [format!("/{}", g.comments), format!("/{}", g.vml)].contains(
                types.nodes[**i]
                    .attrs
                    .get("PartName")
                    .unwrap_or(&String::new()),
            )
        })
        .copied()
        .collect::<Vec<_>>();
    if g.notes.is_empty() && !type_indexes.is_empty() {
        return Err(crate::coded(
            "note-protected",
            "Orphaned note content types stay protected",
        ));
    }
    let revised_types = Tree::parse(&types.mask(&type_indexes)?)?;
    let revised_rels = Tree::parse(&rels.mask(&g.edges)?)?;
    let mut types_xml = String::from_utf8(revised_types.bytes.clone()).map_err(err)?;
    let mut rels_xml = String::from_utf8(revised_rels.bytes.clone()).map_err(err)?;
    let fragment = if notes.is_empty() {
        removed.extend([g.comments.clone(), g.vml.clone()]);
        String::new()
    } else {
        if theme_authoring::relationships(&revised_rels)?
            .iter()
            .any(|i| {
                [workbook_format!(NOTES), workbook_format!(NOTESVML)]
                    .contains(&revised_rels.nodes[*i].attrs["Id"].as_str())
            })
        {
            return Err(crate::coded(
                "note-refused",
                "Note relationship ids already have another owner",
            ));
        }
        replacements.insert(g.comments.clone(), json!(comments(&notes)));
        replacements.insert(g.vml.clone(), json!(vml(&notes)?));
        let a = format!(workbook_format!(XML_3EA451F026),
            String::from_utf8_lossy(PACKAGE_REL_NS),
            g.comments,
            String::from_utf8_lossy(PACKAGE_REL_NS),
            g.vml,
            COMMENTS_EDGE = COMMENTS_EDGE,
            VML_EDGE = VML_EDGE
        );
        rels_xml = theme_authoring::append(&revised_rels, &a)?;
        let a = format!(
            "<Override xmlns=\"{}\" PartName=\"/{}\" ContentType=\"{COMMENTS_MIME}\"/><Override xmlns=\"{}\" PartName=\"/{}\" ContentType=\"{VML_MIME}\"/>",
            String::from_utf8_lossy(TYPE_NS),
            g.comments,
            String::from_utf8_lossy(TYPE_NS),
            g.vml
        );
        types_xml = theme_authoring::append(&revised_types, &a)?;
        format!(workbook_format!(XML_EC5CB73564),
            String::from_utf8_lossy(NS),
            String::from_utf8_lossy(REL_NS)
        )
    };
    replacements.insert(rel_part(&sheet)?, json!(rels_xml));
    replacements.insert("[Content_Types].xml".into(), json!(types_xml));
    let plan = json!({"kind":command["type"],"replacements":replacements,"fragments":{sheet:{"legacyDrawing":fragment}},"copies":{},"removed":removed});
    Ok(Prepared {
        plan,
        command: command.clone(),
        changed: true,
        prior: g.revision,
        remaining: notes.len(),
    })
}
pub(super) fn prepare(root: &Path, input: &Value) -> Result<Prepared> {
    theme_authoring::fields(
        input,
        &[
            "action",
            "operation",
            "root",
            "version",
            "source",
            "rewriteInputId",
            "command",
        ],
        &["diagnostics"],
    )?;
    if input["action"] != "workbook"
        || input["operation"] != "package-edit"
        || input["source"] != "export.xlsx"
        || input.get("diagnostics").is_some_and(|v| !v.is_boolean())
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid native note request",
        ));
    }
    prepare_from(root, "export.xlsx", &input["command"])
}
pub(super) fn verify_output(
    root: &Path,
    source: &str,
    output: &str,
    prepared: &Prepared,
) -> Result<()> {
    let mut before = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, source)?).map_err(err)?,
    ))
    .map_err(err)?;
    let mut after = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, output)?).map_err(err)?,
    ))
    .map_err(err)?;
    let mut names = clear_hyperlinks::package_names(&before)?;
    let replacements = prepared.plan["replacements"]
        .as_object()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing note replacements"))?;
    for n in prepared.plan["removed"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing note removals"))?
    {
        names.remove(
            n.as_str()
                .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid note removal"))?,
        );
    }
    names.extend(replacements.keys().cloned());
    if names != clear_hyperlinks::package_names(&after)? {
        return Err(crate::coded(
            "note-stale",
            "Note writer changed unrelated package inventory",
        ));
    }
    for name in names {
        if let Some(fragment) = prepared.plan["fragments"].get(&name) {
            let (a, _, _) =
                clear_hyperlinks::worksheet_tag_hash(&mut before, &name, None, b"legacyDrawing")?;
            let (b, raw, _) =
                clear_hyperlinks::worksheet_tag_hash(&mut after, &name, None, b"legacyDrawing")?;
            if a != b || raw != field(fragment, "legacyDrawing")?.as_bytes() {
                return Err(crate::coded(
                    "note-stale",
                    "Note writer changed non-note worksheet contents",
                ));
            }
        } else if let Some(xml) = replacements.get(&name) {
            if clear_hyperlinks::member_hash(&mut after, &name)?
                != Sha256::digest(
                    xml.as_str()
                        .ok_or_else(|| {
                            crate::coded("internal-invalid-field", "Invalid note replacement")
                        })?
                        .as_bytes(),
                )
                .to_vec()
            {
                return Err(crate::coded(
                    "internal-mismatch",
                    "Note output differs from the native model",
                ));
            }
        } else if clear_hyperlinks::member_hash(&mut before, &name)?
            != clear_hyperlinks::member_hash(&mut after, &name)?
        {
            return Err(crate::coded(
                "note-stale",
                "Note writer changed unrelated package content",
            ));
        }
    }
    Ok(())
}
pub(super) fn verify_receipt_package(root: &Path, command: &Value, output: &str) -> Result<()> {
    let p = prepare_from(root, "export.xlsx", command)?;
    verify_output(root, "export.xlsx", output, &p)
}
pub(super) fn read(root: &Path, input: &Value) -> Result<Value> {
    theme_authoring::fields(
        input,
        &[
            "action",
            "operation",
            "root",
            "version",
            "sourceRevision",
            "sheet",
        ],
        &["diagnostics"],
    )?;
    if input["action"] != "workbook"
        || input["operation"] != "notes"
        || input.get("diagnostics").is_some_and(|v| !v.is_boolean())
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid native note inspection",
        ));
    }
    let version = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid note inspection version"))?;
    let db = crate::database(root)?;
    if !db
        .query_row(
            "SELECT EXISTS(SELECT 1 FROM versions WHERE id=?1)",
            [version],
            |r| r.get::<_, bool>(0),
        )
        .map_err(err)?
    {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown note inspection version",
        ));
    }
    let index: Value =
        serde_json::from_slice(&fs::read(local(root, "index.json")?).map_err(err)?).map_err(err)?;
    if index["revision"] != input["sourceRevision"]
        || crate::hash(&local(root, "source.xlsx")?)? != field(input, "sourceRevision")?
    {
        return Err(crate::coded(
            "note-stale",
            "Note inspection source revision changed",
        ));
    }
    let sheet = field(input, "sheet")?;
    let inventory = crate::used_range_transfer::worksheet_inventory(&local(root, "source.xlsx")?)?;
    if !inventory.iter().any(|s| s.name == sheet) {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown note worksheet",
        ));
    }
    match definitions(root, "source.xlsx", sheet) {
        Ok((_, _, _, g)) => Ok(
            json!({"schema":workbook_format!(SCHEMA_LEGACY_NOTES_V1),"sheet":sheet,"noteRevision":g.revision,"notes":g.notes.iter().map(|n|json!({"cell":n.cell,"author":n.author,"text":n.text})).collect::<Vec<_>>(),"canEdit":g.read_only.is_none(),"protectedReason":g.read_only}),
        ),
        Err(reason) => Ok(
            json!({"schema":workbook_format!(SCHEMA_LEGACY_NOTES_V1),"sheet":sheet,"noteRevision":null,"notes":null,"canEdit":false,"protectedReason":reason}),
        ),
    }
}
#[cfg(test)]
mod tests;
