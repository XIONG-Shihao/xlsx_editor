//! The zip package as the questions see it: part names, part bytes, and the
//! passive checks on its content types and relationships (what makes a
//! workbook read-only because of macros, ActiveX, signatures or a broken graph).
use super::dom::Dom;
use crate::{Result, coded, err};
use std::{collections::BTreeSet, fs::File, io::Read, path::Path};
use zip::ZipArchive;

const TYPES: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

pub(super) struct Package {
    zip: ZipArchive<crate::cancel::Reader<File>>,
    pub names: Vec<String>,
}

impl Package {
    pub fn open(path: &Path) -> Result<Package> {
        let file = File::open(path).map_err(err)?;
        let zip = ZipArchive::new(crate::cancel::Reader(file))
            .map_err(|e| coded("package-malformed", e))?;
        let mut zip = zip;
        let mut names = Vec::new();
        for index in 0..zip.len() {
            let member = zip
                .by_index_raw(index)
                .map_err(|e| coded("package-malformed", e))?;
            names.push(member.name().to_string());
        }
        Ok(Package { zip, names })
    }
    pub fn has(&self, name: &str) -> bool {
        self.names.iter().any(|n| n == name)
    }
    pub fn bytes(&mut self, name: &str) -> Result<Vec<u8>> {
        let mut member = self
            .zip
            .by_name(name)
            .map_err(|e| coded("package-malformed", e))?;
        let mut data = Vec::with_capacity(member.size() as usize);
        member.read_to_end(&mut data).map_err(err)?;
        Ok(data)
    }
    pub fn size(&mut self, name: &str) -> Option<u64> {
        self.zip.by_name(name).ok().map(|m| m.size())
    }
    pub fn dom(&mut self, name: &str) -> Result<Dom> {
        Dom::parse(&self.bytes(name)?)
    }
}

/// A package graph part: strict UTF-8, no entities, the expected root.
fn graph(bytes: &[u8], namespace: &str, root: &str) -> Result<Dom> {
    std::str::from_utf8(bytes).map_err(|_| coded("internal-invalid-field", "Invalid UTF-8"))?;
    let dom = Dom::parse(bytes)?;
    if !dom.root_is(namespace, root) {
        return Err(coded("package-malformed", "Invalid package graph namespace/root"));
    }
    Ok(dom)
}
/// The element children of a graph node; loose text among them is refused.
fn elements(dom: &Dom, of: usize) -> Result<Vec<usize>> {
    if dom.has_loose_text(of) {
        return Err(coded("internal-invalid-field", "Invalid package graph mixed text"));
    }
    Ok(dom.child_elements(of))
}
// Macros do not make a workbook read-only: the macro project, its signature
// parts and the macro-enabled content type are carried byte for byte, as Excel
// does when it edits a macro workbook. Controls and package signatures stay
// guarded: an edit would have to move the controls or would break the signature.
fn guarded_type(content_type: &str) -> bool {
    let t = content_type.to_ascii_lowercase();
    t.starts_with("application/vnd.ms-office.activex")
        || t.starts_with("application/vnd.openxmlformats-package.digital-signature-")
}
fn guarded_relation(kind: &str) -> bool {
    let k = kind.to_ascii_lowercase();
    let last = k.rsplit_once('/').map(|(_, last)| last);
    last.is_some_and(|last| last == "activexcontrolbinary" || last == "control") || k.contains("/digital-signature/")
}
/// True when the package carries a macro project (VBA).
pub(super) fn macro_project(package: &Package) -> bool {
    package.names.iter().any(|n| {
        let n = n.to_ascii_lowercase();
        n == "vbaproject.bin" || n.ends_with("/vbaproject.bin")
    })
}
fn content_type_shape(value: &str) -> bool {
    // ^[^\s;/]+\/[^\s;]+(?:;[^\r\n]*)?$
    let (main, parameters) = match value.split_once(';') {
        Some((main, rest)) => (main, Some(rest)),
        None => (value, None),
    };
    let Some((kind, sub)) = main.split_once('/') else {
        return false;
    };
    let plain = |s: &str| !s.is_empty() && !s.chars().any(|c| c.is_whitespace() || c == ';');
    plain(kind)
        && !kind.contains('/')
        && plain(sub)
        && parameters.is_none_or(|p| !p.contains(['\r', '\n']))
}
fn relationship_type_shape(value: &str) -> bool {
    // ^[A-Za-z][A-Za-z0-9+.-]*:[^\s]+$
    let Some((scheme, rest)) = value.split_once(':') else {
        return false;
    };
    let mut chars = scheme.chars();
    chars.next().is_some_and(|c| c.is_ascii_alphabetic())
        && chars.all(|c| c.is_ascii_alphanumeric() || "+.-".contains(c))
        && !rest.is_empty()
        && !rest.chars().any(char::is_whitespace)
}

fn content_types(package: &mut Package, guard: &mut impl FnMut()) -> Result<()> {
    if !package.has("[Content_Types].xml") {
        return Err(coded("internal-missing-field", "Missing package content types"));
    }
    let dom = graph(&package.bytes("[Content_Types].xml")?, TYPES, "Types")?;
    let mut seen = BTreeSet::new();
    for child in elements(&dom, 0)? {
        let node = &dom.nodes[child];
        if node.ns != TYPES
            || !["Default", "Override"].contains(&node.local.as_str())
            || !elements(&dom, child)?.is_empty()
        {
            return Err(coded("package-malformed", "Invalid package content type entry"));
        }
        let key_attribute = if node.local == "Default" { "Extension" } else { "PartName" };
        let key = format!("{}:{}", node.local, dom.attr(child, key_attribute));
        let kind = dom.attr(child, "ContentType");
        if kind.is_empty()
            || kind.trim() != kind
            || !content_type_shape(kind)
            || key.ends_with(':')
            || !seen.insert(key)
        {
            return Err(coded(
                "internal-missing-field",
                "Missing, malformed or duplicate package content type",
            ));
        }
        if guarded_type(kind) {
            guard();
        }
    }
    Ok(())
}
fn relationships(bytes: &[u8], guard: &mut impl FnMut()) -> Result<()> {
    let dom = graph(bytes, RELS, "Relationships")?;
    let mut seen = BTreeSet::new();
    for child in elements(&dom, 0)? {
        if dom.nodes[child].ns != RELS
            || dom.nodes[child].local != "Relationship"
            || !elements(&dom, child)?.is_empty()
        {
            return Err(coded("package-malformed", "Invalid package relationship entry"));
        }
        let (id, kind, target) = (
            dom.attr(child, "Id"),
            dom.attr(child, "Type"),
            dom.attr(child, "Target"),
        );
        if id.is_empty()
            || kind.is_empty()
            || !relationship_type_shape(kind)
            || target.is_empty()
            || !seen.insert(id.to_string())
        {
            return Err(coded(
                "internal-missing-field",
                "Missing, malformed or duplicate package relationship",
            ));
        }
        // Even a missing, external or malformed sensitive target counts: this
        // reader never follows the target and cannot edit it.
        if guarded_relation(kind) {
            guard();
        }
    }
    Ok(())
}

/// Why the package alone must be opened read-only, in the order found.
pub(super) fn read_only_reasons(package: &mut Package) -> Vec<String> {
    let mut reasons: Vec<String> = Vec::new();
    let macros = "ActiveX controls and signed workbooks are read-only";
    let guard = |reasons: &mut Vec<String>| reasons.push(macros.to_string());
    let lowered = |n: &String| n.to_ascii_lowercase();
    if package.names.iter().any(|n| {
        let n = lowered(n);
        n.contains("activex") || n.contains("_xmlsignatures")
    }) {
        guard(&mut reasons);
    }
    let mut found = 0usize;
    let result = content_types(package, &mut || found += 1);
    for _ in 0..found {
        guard(&mut reasons);
    }
    if result.is_err() {
        reasons.push("The package content type graph is invalid; editing is not qualified".into());
    }
    let rels: Vec<String> = package
        .names
        .iter()
        .filter(|n| n.ends_with(".rels"))
        .cloned()
        .collect();
    for name in rels {
        let mut found = 0usize;
        let result = package
            .bytes(&name)
            .and_then(|bytes| relationships(&bytes, &mut || found += 1));
        for _ in 0..found {
            guard(&mut reasons);
        }
        if result.is_err() {
            reasons.push("The package relationship graph is invalid; editing is not qualified".into());
        }
    }
    let mut seen = BTreeSet::new();
    reasons.retain(|r| seen.insert(r.clone()));
    reasons
}
