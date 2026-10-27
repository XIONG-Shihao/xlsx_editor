//! Change a few values in an XML part and leave every other byte as it was.
//!
//! A part is read once into a flat list of elements in document order. Edits
//! are chosen by an element's place in that list and then written by streaming
//! the same part again, so attribute order, prefixes, comments and spacing of
//! everything not edited stay exactly as stored.
use crate::{Result, err};
use quick_xml::{
    NsReader, Reader,
    events::{BytesEnd, BytesStart, Event},
};
use std::collections::BTreeMap;

const LIMIT_BYTES: usize = 64 * 1024 * 1024;
const LIMIT_NODES: usize = 2_000_000;

pub(super) struct Node {
    pub ns: Vec<u8>,
    pub name: String,
    pub qname: String,
    /// Attributes as written (prefix included), in their stored order.
    pub attrs: Vec<(String, String)>,
    pub text: String,
    pub kids: Vec<usize>,
}
pub(super) struct Dom {
    pub nodes: Vec<Node>,
}
impl Dom {
    pub fn parse(bytes: &[u8]) -> Result<Self> {
        if bytes.len() > LIMIT_BYTES {
            return Err(crate::coded("package-limit", "Package XML exceeds its bound"));
        }
        let mut reader = NsReader::from_reader(bytes);
        let mut buffer = Vec::new();
        let mut nodes: Vec<Node> = Vec::new();
        let mut stack: Vec<usize> = Vec::new();
        loop {
            crate::cancel::check()?;
            let (binding, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
            match event {
                Event::Start(ref element) | Event::Empty(ref element) => {
                    if nodes.len() >= LIMIT_NODES || stack.len() > 256 {
                        return Err(crate::coded("package-limit", "Package XML exceeds its bound"));
                    }
                    let mut attrs = Vec::new();
                    for attribute in element.attributes().with_checks(false) {
                        let attribute = attribute.map_err(err)?;
                        attrs.push((
                            String::from_utf8_lossy(attribute.key.as_ref()).into_owned(),
                            attribute.unescape_value().map_err(err)?.into_owned(),
                        ));
                    }
                    let index = nodes.len();
                    nodes.push(Node {
                        ns: match binding {
                            quick_xml::name::ResolveResult::Bound(ns) => ns.as_ref().to_vec(),
                            _ => Vec::new(),
                        },
                        name: String::from_utf8_lossy(element.local_name().as_ref()).into_owned(),
                        qname: String::from_utf8_lossy(element.name().as_ref()).into_owned(),
                        attrs,
                        text: String::new(),
                        kids: Vec::new(),
                    });
                    if let Some(parent) = stack.last() {
                        nodes[*parent].kids.push(index);
                    }
                    if matches!(event, Event::Start(_)) {
                        stack.push(index);
                    }
                }
                Event::End(_) => {
                    stack
                        .pop()
                        .ok_or_else(|| crate::coded("package-malformed", "Unbalanced XML"))?;
                }
                Event::Text(ref text) => {
                    if let Some(index) = stack.last() {
                        let value = text.unescape().map_err(err)?;
                        nodes[*index].text.push_str(&value);
                    }
                }
                Event::CData(ref text) => {
                    if let Some(index) = stack.last() {
                        nodes[*index]
                            .text
                            .push_str(std::str::from_utf8(text).map_err(err)?);
                    }
                }
                Event::DocType(_) => {
                    return Err(crate::coded(
                        "package-malformed",
                        "XML entities are unsupported",
                    ));
                }
                Event::Eof => break,
                _ => {}
            }
            buffer.clear();
        }
        if nodes.is_empty() || !stack.is_empty() {
            return Err(crate::coded("package-malformed", "Invalid XML part"));
        }
        Ok(Self { nodes })
    }
    pub fn attr(&self, node: usize, key: &str) -> Option<&str> {
        self.nodes[node]
            .attrs
            .iter()
            .find(|(name, _)| name == key)
            .map(|(_, value)| value.as_str())
    }
    /// The value of the attribute with this local name, whatever its prefix.
    pub fn attr_local(&self, node: usize, local: &str) -> Option<&str> {
        self.nodes[node]
            .attrs
            .iter()
            .find(|(name, _)| name.rsplit(':').next() == Some(local) && !name.starts_with("xmlns"))
            .map(|(_, value)| value.as_str())
    }
    pub fn is(&self, node: usize, ns: &[u8], name: &str) -> bool {
        self.nodes[node].ns == ns && self.nodes[node].name == name
    }
    /// Every element with this name, in document order.
    pub fn all(&self, ns: &[u8], name: &str) -> Vec<usize> {
        (0..self.nodes.len())
            .filter(|index| self.is(*index, ns, name))
            .collect()
    }
    pub fn kids_named(&self, node: usize, ns: &[u8], name: &str) -> Vec<usize> {
        self.nodes[node]
            .kids
            .iter()
            .copied()
            .filter(|kid| self.is(*kid, ns, name))
            .collect()
    }
}

#[derive(Clone, Default)]
pub(super) enum Change {
    #[default]
    Keep,
    /// New values for some attributes (added at the end when absent).
    Attrs(BTreeMap<String, String>),
    /// New text for an element that holds only text.
    Text(String),
    /// The element and everything inside it.
    Drop,
}
#[derive(Clone, Default)]
pub(super) struct Edit {
    pub change: Change,
    /// Raw XML written just before the element, and just after it.
    pub before: String,
    pub after: String,
}
impl Edit {
    pub fn attrs<I: IntoIterator<Item = (&'static str, String)>>(pairs: I) -> Self {
        Self {
            change: Change::Attrs(pairs.into_iter().map(|(k, v)| (k.to_owned(), v)).collect()),
            ..Default::default()
        }
    }
    pub fn text(text: String) -> Self {
        Self {
            change: Change::Text(text),
            ..Default::default()
        }
    }
    pub fn drop() -> Self {
        Self {
            change: Change::Drop,
            ..Default::default()
        }
    }
}

pub(super) fn escape(text: &str, attribute: bool) -> String {
    let mut out = String::with_capacity(text.len());
    for character in text.chars() {
        match character {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' if attribute => out.push_str("&quot;"),
            '\n' if attribute => out.push_str("&#10;"),
            '\r' => out.push_str("&#13;"),
            '\t' if attribute => out.push_str("&#9;"),
            other => out.push(other),
        }
    }
    out
}

fn restart(element: &BytesStart<'_>, values: &BTreeMap<String, String>) -> Result<BytesStart<'static>> {
    let mut changed = BytesStart::new(String::from_utf8_lossy(element.name().as_ref()).into_owned());
    let mut used = Vec::new();
    for attribute in element.attributes().with_checks(false) {
        let attribute = attribute.map_err(err)?;
        let key = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
        match values.get(&key) {
            Some(value) => {
                changed.push_attribute((key.as_str(), value.as_str()));
                used.push(key);
            }
            None => {
                let value = attribute.unescape_value().map_err(err)?.into_owned();
                changed.push_attribute((key.as_str(), value.as_str()));
            }
        }
    }
    for (key, value) in values {
        if !used.contains(key) {
            changed.push_attribute((key.as_str(), value.as_str()));
        }
    }
    Ok(changed)
}

fn put(out: &mut Vec<u8>, event: Event<'_>) -> Result<()> {
    quick_xml::Writer::new(out).write_event(event).map_err(err)
}

/// Write `bytes` again with the edits applied. An edit is addressed by the
/// element's place among all elements (start order), as `Dom::parse` lists them.
pub(super) fn apply(bytes: &[u8], edits: &BTreeMap<usize, Edit>) -> Result<Vec<u8>> {
    let mut reader = Reader::from_reader(bytes);
    let mut out: Vec<u8> = Vec::with_capacity(bytes.len() + 256);
    let mut buffer = Vec::new();
    let mut ordinal = 0usize;
    // Elements still open that have text to write after them (their place).
    let mut open: Vec<Option<usize>> = Vec::new();
    // An element being left out, with how deep inside it the reader is.
    let mut skipping: Option<(usize, usize)> = None;
    loop {
        crate::cancel::check()?;
        let event = reader.read_event_into(&mut buffer).map_err(err)?;
        if let Some((at, depth)) = &mut skipping {
            match event {
                Event::Start(_) => {
                    *depth += 1;
                    ordinal += 1;
                }
                Event::Empty(_) => ordinal += 1,
                Event::End(_) => {
                    *depth -= 1;
                    if *depth == 0 {
                        out.extend_from_slice(edits[at].after.as_bytes());
                        skipping = None;
                    }
                }
                Event::Eof => return Err(crate::coded("package-malformed", "Unbalanced XML")),
                _ => {}
            }
            buffer.clear();
            continue;
        }
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let empty = matches!(event, Event::Empty(_));
                let at = ordinal;
                ordinal += 1;
                let Some(edit) = edits.get(&at) else {
                    put(&mut out, event.clone())?;
                    if !empty {
                        open.push(None);
                    }
                    buffer.clear();
                    continue;
                };
                out.extend_from_slice(edit.before.as_bytes());
                match &edit.change {
                    Change::Drop => {}
                    Change::Text(text) => {
                        let name = String::from_utf8_lossy(element.name().as_ref()).into_owned();
                        put(&mut out, Event::Start(element.to_owned()))?;
                        out.extend_from_slice(escape(text, false).as_bytes());
                        put(&mut out, Event::End(BytesEnd::new(name)))?;
                    }
                    Change::Attrs(values) => {
                        let changed = restart(element, values)?;
                        put(
                            &mut out,
                            if empty {
                                Event::Empty(changed)
                            } else {
                                Event::Start(changed)
                            },
                        )?;
                    }
                    Change::Keep => put(&mut out, event.clone())?,
                }
                let whole = matches!(edit.change, Change::Drop | Change::Text(_));
                if empty {
                    out.extend_from_slice(edit.after.as_bytes());
                } else if whole {
                    skipping = Some((at, 1));
                } else {
                    open.push(if edit.after.is_empty() { None } else { Some(at) });
                }
            }
            Event::End(ref end) => {
                put(&mut out, Event::End(end.to_owned()))?;
                if let Some(Some(at)) = open.pop() {
                    out.extend_from_slice(edits[&at].after.as_bytes());
                }
            }
            Event::Eof => break,
            other => put(&mut out, other.into_owned())?,
        }
        buffer.clear();
    }
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn edits_leave_the_rest_of_the_part_as_written() {
        let xml = "<?xml version=\"1.0\"?>\n<r:Rels xmlns:r=\"urn:x\" a='1'><r:Rel Id=\"1\" T=\"a&amp;b\"/><r:Rel Id=\"2\"><r:in>x</r:in></r:Rel><r:Rel Id=\"3\">old &amp; text</r:Rel><!-- note --></r:Rels>";
        let dom = Dom::parse(xml.as_bytes()).unwrap();
        let rels = dom.all(b"urn:x", "Rel");
        assert_eq!(rels.len(), 3);
        let mut edits = BTreeMap::new();
        edits.insert(rels[0], Edit::drop());
        edits.insert(rels[1], Edit::drop());
        edits.insert(rels[2], Edit::text("new <t>".into()));
        let out = String::from_utf8(apply(xml.as_bytes(), &edits).unwrap()).unwrap();
        assert_eq!(
            out,
            "<?xml version=\"1.0\"?>\n<r:Rels xmlns:r=\"urn:x\" a='1'><r:Rel Id=\"3\">new &lt;t&gt;</r:Rel><!-- note --></r:Rels>"
        );
    }
    #[test]
    fn attributes_change_in_place_and_text_is_added_around_an_element() {
        let xml = "<a xmlns=\"urn:x\"><b k=\"1\" z=\"2\"/><c k=\"1\">t</c></a>";
        let dom = Dom::parse(xml.as_bytes()).unwrap();
        let b = dom.all(b"urn:x", "b")[0];
        let c = dom.all(b"urn:x", "c")[0];
        let mut edits = BTreeMap::new();
        edits.insert(b, Edit::attrs([("k", "9".to_owned()), ("n", "new".to_owned())]));
        let mut around = Edit::attrs([("k", "7".to_owned())]);
        around.before = "<p/>".into();
        around.after = "<q/>".into();
        edits.insert(c, around);
        let out = String::from_utf8(apply(xml.as_bytes(), &edits).unwrap()).unwrap();
        assert_eq!(
            out,
            "<a xmlns=\"urn:x\"><b k=\"9\" z=\"2\" n=\"new\"/><p/><c k=\"7\">t</c><q/></a>"
        );
    }
}
