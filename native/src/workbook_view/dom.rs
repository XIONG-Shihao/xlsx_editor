//! A small namespace-aware XML tree for reading workbook parts. It keeps what
//! the questions about a workbook need (names, attributes, text) and nothing
//! about byte positions; the writers use their own readers.
use crate::{Result, coded, err};
use quick_xml::{
    NsReader,
    events::Event,
    name::ResolveResult,
};

pub(super) const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

pub(super) struct Attr {
    pub ns: String,
    pub local: String,
    /// The name as written, with its prefix.
    pub qname: String,
    pub value: String,
}
pub(super) enum Kid {
    Element(usize),
    Text(String),
    /// A comment or processing instruction: it counts as a child, nothing more.
    Other,
}
pub(super) struct El {
    pub ns: String,
    pub local: String,
    pub attrs: Vec<Attr>,
    pub kids: Vec<Kid>,
}
pub(super) struct Dom {
    pub nodes: Vec<El>,
}

impl Dom {
    /// Reads one XML part. Entities and document type declarations are refused.
    pub fn parse(bytes: &[u8]) -> Result<Dom> {
        let bytes = bytes.strip_prefix(b"\xEF\xBB\xBF").unwrap_or(bytes);
        let lower = String::from_utf8_lossy(bytes).to_ascii_lowercase();
        if lower.contains("<!doctype") || lower.contains("<!entity") {
            return Err(coded("package-malformed", "XML entities are unsupported"));
        }
        let malformed = |error: &dyn std::fmt::Display| coded("package-malformed", error);
        let mut reader = NsReader::from_reader(bytes);
        let mut buffer = Vec::new();
        let mut nodes: Vec<El> = Vec::new();
        let mut stack: Vec<usize> = Vec::new();
        loop {
            crate::cancel::check()?;
            let (binding, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|e| malformed(&e))?;
            match event {
                Event::Start(ref element) | Event::Empty(ref element) => {
                    let ns = match binding {
                        ResolveResult::Bound(ns) => String::from_utf8_lossy(ns.as_ref()).into_owned(),
                        _ => String::new(),
                    };
                    let mut attrs = Vec::new();
                    for attribute in element.attributes().with_checks(true) {
                        let attribute = attribute.map_err(|e| malformed(&e))?;
                        let qname = std::str::from_utf8(attribute.key.as_ref())
                            .map_err(err)?
                            .to_string();
                        if qname == "xmlns" || qname.starts_with("xmlns:") {
                            continue;
                        }
                        let (binding, local) = reader.resolve_attribute(attribute.key);
                        let ns = match binding {
                            ResolveResult::Bound(ns) => {
                                String::from_utf8_lossy(ns.as_ref()).into_owned()
                            }
                            _ => String::new(),
                        };
                        attrs.push(Attr {
                            ns,
                            local: String::from_utf8_lossy(local.as_ref()).into_owned(),
                            qname,
                            value: attribute.unescape_value().map_err(|e| malformed(&e))?.into_owned(),
                        });
                    }
                    let id = nodes.len();
                    if let Some(parent) = stack.last() {
                        nodes[*parent].kids.push(Kid::Element(id));
                    } else if id != 0 {
                        return Err(coded("package-malformed", "Multiple XML roots"));
                    }
                    nodes.push(El {
                        ns,
                        local: String::from_utf8_lossy(element.local_name().as_ref()).into_owned(),
                        attrs,
                        kids: Vec::new(),
                    });
                    if matches!(event, Event::Start(_)) {
                        stack.push(id);
                    }
                }
                Event::End(_) => {
                    stack.pop();
                }
                Event::Text(ref text) => {
                    if let Some(parent) = stack.last() {
                        let value = crate::xml::xml_text(text.as_ref()).map_err(|e| malformed(&e))?;
                        nodes[*parent].kids.push(Kid::Text(value.into_owned()));
                    }
                }
                Event::CData(ref text) => {
                    if let Some(parent) = stack.last() {
                        let value = String::from_utf8_lossy(text.as_ref()).into_owned();
                        nodes[*parent].kids.push(Kid::Text(value));
                    }
                }
                Event::Comment(_) | Event::PI(_) => {
                    if let Some(parent) = stack.last() {
                        nodes[*parent].kids.push(Kid::Other);
                    }
                }
                Event::Eof => break,
                _ => {}
            }
        }
        if nodes.is_empty() || !stack.is_empty() {
            return Err(coded("package-malformed", "Incomplete XML document"));
        }
        Ok(Dom { nodes })
    }

    /// Element children, in order.
    pub fn child_elements(&self, of: usize) -> Vec<usize> {
        self.nodes[of]
            .kids
            .iter()
            .filter_map(|kid| match kid {
                Kid::Element(i) => Some(*i),
                _ => None,
            })
            .collect()
    }
    /// Element children in the workbook namespace with this local name.
    pub fn children(&self, of: usize, local: &str) -> Vec<usize> {
        self.child_elements(of)
            .into_iter()
            .filter(|i| self.is(*i, local))
            .collect()
    }
    pub fn is(&self, node: usize, local: &str) -> bool {
        self.nodes[node].ns == MAIN && self.nodes[node].local == local
    }
    /// Every descendant in the workbook namespace with this local name, in document order.
    pub fn elements(&self, of: usize, local: &str) -> Vec<usize> {
        let mut found = Vec::new();
        self.walk(of, &mut |i| {
            if i != of && self.is(i, local) {
                found.push(i);
            }
        });
        found
    }
    pub fn first(&self, of: usize, local: &str) -> Option<usize> {
        let mut found = None;
        self.walk(of, &mut |i| {
            if found.is_none() && i != of && self.is(i, local) {
                found = Some(i);
            }
        });
        found
    }
    pub fn walk(&self, of: usize, visit: &mut impl FnMut(usize)) {
        visit(of);
        for kid in &self.nodes[of].kids {
            if let Kid::Element(i) = kid {
                self.walk(*i, visit);
            }
        }
    }
    /// The attribute by its written name; an absent attribute reads as empty.
    pub fn attr(&self, node: usize, name: &str) -> &str {
        self.attribute(node, name).unwrap_or("")
    }
    pub fn attribute(&self, node: usize, name: &str) -> Option<&str> {
        self.nodes[node]
            .attrs
            .iter()
            .find(|a| a.qname == name)
            .map(|a| a.value.as_str())
    }
    pub fn text(&self, node: usize) -> String {
        let mut out = String::new();
        self.collect(node, &mut out);
        out
    }
    fn collect(&self, node: usize, out: &mut String) {
        for kid in &self.nodes[node].kids {
            match kid {
                Kid::Text(t) => out.push_str(t),
                Kid::Element(i) => self.collect(*i, out),
                Kid::Other => {}
            }
        }
    }
    /// Text directly inside the element that is more than whitespace.
    pub fn has_loose_text(&self, node: usize) -> bool {
        self.nodes[node]
            .kids
            .iter()
            .any(|kid| matches!(kid, Kid::Text(t) if !t.trim().is_empty()))
    }
    pub fn root_is(&self, ns: &str, local: &str) -> bool {
        self.nodes[0].ns == ns && self.nodes[0].local == local
    }
}
