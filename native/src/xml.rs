use crate::{Result, err};
use quick_xml::events::BytesStart;
use std::{borrow::Cow, collections::BTreeMap};

pub fn xml_text(bytes: &[u8]) -> Result<Cow<'_, str>> {
    let text = std::str::from_utf8(bytes).map_err(err)?;
    if !text.contains('\r') {
        return quick_xml::escape::unescape(text).map_err(err);
    }
    let normalized = text.replace("\r\n", "\n").replace('\r', "\n");
    Ok(Cow::Owned(
        quick_xml::escape::unescape(&normalized)
            .map_err(err)?
            .into_owned(),
    ))
}

/// Common element names borrow static storage rather than allocating for each
/// cell/value event. Unknown names retain the existing namespace/stack behavior.
pub fn node_name(bound: bool, element: &BytesStart<'_>) -> Cow<'static, str> {
    if !bound {
        return Cow::Borrowed("");
    }
    let local = element.local_name();
    let known = match local.as_ref() {
        b"worksheet" => "worksheet",
        b"sst" => "sst",
        b"si" => "si",
        b"sheetData" => "sheetData",
        b"row" => "row",
        b"c" => "c",
        b"v" => "v",
        b"f" => "f",
        b"is" => "is",
        b"t" => "t",
        b"r" => "r",
        b"rPh" => "rPh",
        b"cols" => "cols",
        b"col" => "col",
        b"dimension" => "dimension",
        b"extLst" => "extLst",
        _ => return Cow::Owned(String::from_utf8_lossy(local.as_ref()).into_owned()),
    };
    Cow::Borrowed(known)
}

/// Attribute views never outlive the current XML event buffer.
pub fn borrowed_attrs<'a>(element: &'a BytesStart<'_>) -> Result<BTreeMap<&'a str, Cow<'a, str>>> {
    element
        .attributes()
        .map(|attribute| {
            let attribute = attribute.map_err(err)?;
            Ok((
                std::str::from_utf8(attribute.key.0).map_err(err)?,
                attribute.unescape_value().map_err(err)?,
            ))
        })
        .collect()
}

pub const NS: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub fn attrs(element: &BytesStart<'_>) -> Result<BTreeMap<String, String>> {
    element
        .attributes()
        .map(|attribute| {
            let attribute = attribute.map_err(err)?;
            Ok((
                std::str::from_utf8(attribute.key.as_ref())
                    .map_err(err)?
                    .to_string(),
                attribute.unescape_value().map_err(err)?.into_owned(),
            ))
        })
        .collect()
}
pub fn position(reference: &str) -> Result<(u32, u32)> {
    let mut row = 0u32;
    let mut column = 0u32;
    let mut in_row = false;
    for byte in reference.bytes() {
        if byte == b'$' {
            continue;
        }
        if byte.is_ascii_alphabetic() && !in_row {
            column = column
                .checked_mul(26)
                .and_then(|v| v.checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1)))
                .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid cell"))?;
        } else if byte.is_ascii_digit() {
            in_row = true;
            row = row
                .checked_mul(10)
                .and_then(|v| v.checked_add(u32::from(byte - b'0')))
                .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid cell"))?;
        } else {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid cell address",
            ));
        }
    }
    if !(1..=1048576).contains(&row) || !(1..=16384).contains(&column) {
        return Err(crate::coded(
            "package-unsupported",
            "Cell is outside the worksheet",
        ));
    }
    Ok((row - 1, column - 1))
}
pub fn address(row: u32, col: u32) -> String {
    let mut letters = Vec::new();
    let mut n = col + 1;
    while n > 0 {
        letters.push((b'A' + ((n - 1) % 26) as u8) as char);
        n = (n - 1) / 26;
    }
    format!("{}{}", letters.iter().rev().collect::<String>(), row + 1)
}
/// SpreadsheetML escapes are decoded once, including UTF-16 surrogate pairs.
pub fn spreadsheet_text(value: &str) -> Result<Cow<'_, str>> {
    if !value.contains("_x") {
        return Ok(Cow::Borrowed(value));
    }
    let bytes = value.as_bytes();
    let mut units = Vec::new();
    let mut start = 0;
    let mut i = 0;
    while i + 7 <= bytes.len() {
        if bytes[i] == b'_'
            && bytes[i + 1] == b'x'
            && bytes[i + 6] == b'_'
            && bytes[i + 2..i + 6].iter().all(u8::is_ascii_hexdigit)
        {
            units.extend(value[start..i].encode_utf16());
            units.push(u16::from_str_radix(&value[i + 2..i + 6], 16).map_err(err)?);
            i += 7;
            start = i;
        } else {
            i += 1;
        }
    }
    if start == 0 {
        return Ok(Cow::Borrowed(value));
    }
    units.extend(value[start..].encode_utf16());
    String::from_utf16(&units)
        .map(Cow::Owned)
        .map_err(|_| crate::coded("package-refused", "Unpaired SpreadsheetML text escape"))
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn borrowed_text_preserves_xml_and_spreadsheet_escapes() {
        assert!(matches!(xml_text(b"123.45").unwrap(), Cow::Borrowed(_)));
        assert_eq!(xml_text(b"a\r\nb\rc&amp;d&#13;").unwrap(), "a\nb\nc&d\r");
        assert!(xml_text(b"&unknown;").is_err());
        assert!(xml_text(&[0xff]).is_err());
        assert!(matches!(
            spreadsheet_text("plain 🙂").unwrap(),
            Cow::Borrowed(_)
        ));
        assert_eq!(spreadsheet_text("_x005F_x0041_").unwrap(), "_x0041_");
        assert_eq!(spreadsheet_text("_xD83D__xDE42_").unwrap(), "🙂");
        assert!(spreadsheet_text("_xD83D_").is_err());
    }
    #[test]
    fn borrowed_attribute_and_name_views_preserve_validation() {
        let element = BytesStart::from_content("c r=\"A1\" t=\"str\" custom=\"a&amp;b\"", 1);
        let borrowed = borrowed_attrs(&element).unwrap();
        assert!(matches!(borrowed["r"], Cow::Borrowed(_)));
        assert_eq!(borrowed["custom"], "a&b");
        let owned = borrowed
            .iter()
            .map(|(k, v)| (k.to_string(), v.to_string()))
            .collect::<BTreeMap<_, _>>();
        assert_eq!(owned, attrs(&element).unwrap());
        assert_eq!(node_name(true, &element), "c");
        assert_eq!(node_name(false, &element), "");
        assert_eq!(node_name(true, &BytesStart::new("custom")), "custom");
        assert!(borrowed_attrs(&BytesStart::from_content("c r=\"A1\" r=\"A2\"", 1)).is_err());
    }
}
