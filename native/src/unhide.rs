//! Qualified visibility changes only; all cell and unrelated events pass through.
use crate::{Result, xml::attrs};
use quick_xml::events::BytesStart;
fn positive(value: Option<&String>, fallback: f64) -> Result<f64> {
    let v = match value {
        Some(s) => s.parse::<f64>().map_err(|_| {
            crate::coded(
                "internal-invalid-field",
                "Invalid source visibility dimension",
            )
        })?,
        None => fallback,
    };
    if !v.is_finite() || v < 0.0 {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid source visibility dimension",
        ));
    }
    Ok(if v == 0.0 { fallback } else { v })
}
fn flag(value: Option<&String>) -> Result<bool> {
    match value.map(String::as_str) {
        None | Some("0" | "false") => Ok(false),
        Some("1" | "true") => Ok(true),
        _ => Err(crate::coded(
            "internal-invalid-field",
            "Invalid source visibility flag",
        )),
    }
}
/** Preserve positive custom dimensions and opaque attributes; zero geometry uses source defaults. */
pub(crate) fn visible(
    e: &BytesStart<'_>,
    kind: &[u8],
    row_default: f64,
    col_default: f64,
) -> Result<BytesStart<'static>> {
    let mut a = attrs(e)?;
    match kind {
        b"sheetFormatPr" => {
            if flag(a.get("zeroHeight"))? {
                a.insert("zeroHeight".into(), "0".into());
            }
            if a.get("defaultRowHeight")
                .is_some_and(|s| s.parse::<f64>().ok() == Some(0.0))
            {
                a.insert("defaultRowHeight".into(), "15".into());
            }
            if a.get("defaultColWidth")
                .is_some_and(|s| s.parse::<f64>().ok() == Some(0.0))
            {
                a.insert("defaultColWidth".into(), "8.43".into());
            }
        }
        b"row" => {
            if flag(a.get("hidden"))? {
                a.insert("hidden".into(), "0".into());
            }
            let height = positive(a.get("ht"), row_default)?;
            if a.get("ht")
                .is_some_and(|s| s.parse::<f64>().ok() == Some(0.0))
            {
                a.insert("ht".into(), height.to_string());
            }
        }
        b"col" => {
            if flag(a.get("hidden"))? {
                a.insert("hidden".into(), "0".into());
            }
            let width = positive(a.get("width"), col_default)?;
            if a.get("width")
                .is_some_and(|s| s.parse::<f64>().ok() == Some(0.0))
            {
                a.insert("width".into(), width.to_string());
            }
        }
        _ => {
            return Err(crate::coded(
                "internal-unknown-reference",
                "Unknown visibility element",
            ));
        }
    }
    let mut out = BytesStart::new(String::from_utf8_lossy(e.name().as_ref()).to_string());
    for (k, v) in a {
        out.push_attribute((k.as_str(), v.as_str()));
    }
    Ok(out)
}
pub(crate) fn defaults(e: &BytesStart<'_>) -> Result<(f64, f64)> {
    let a = attrs(e)?;
    Ok((
        positive(a.get("defaultRowHeight"), 15.0)?,
        positive(a.get("defaultColWidth"), 8.43)?,
    ))
}
#[cfg(test)]
mod tests {
    use super::*;
    fn element(text: &str) -> BytesStart<'static> {
        BytesStart::from_content(text.to_string(), text.find(' ').unwrap_or(text.len()))
    }
    fn values(e: BytesStart<'static>) -> std::collections::BTreeMap<String, String> {
        attrs(&e).unwrap()
    }
    #[test]
    fn hidden_and_zero_dimensions_use_qualified_defaults() {
        let row = values(
            visible(
                &element("s:row r=\"700000\" hidden=\"true\" ht=\"0\" s=\"7\""),
                b"row",
                22.0,
                12.0,
            )
            .unwrap(),
        );
        assert_eq!(row["hidden"], "0");
        assert_eq!(row["ht"], "22");
        assert_eq!(row["r"], "700000");
        assert_eq!(row["s"], "7");
        let col = values(
            visible(
                &element("col min=\"4\" max=\"9\" hidden=\"1\" width=\"0\""),
                b"col",
                22.0,
                12.0,
            )
            .unwrap(),
        );
        assert_eq!(col["width"], "12");
        assert_eq!(col["hidden"], "0");
    }
    #[test]
    fn positive_custom_geometry_and_escaped_attributes_are_preserved() {
        let row=values(visible(&element("row r=\"2\" hidden=\"1\" ht=\"37.5\" xmlns:x=\"urn:test\" x:note=\"A &amp; B\""),b"row",15.0,8.43).unwrap());
        assert_eq!(row["ht"], "37.5");
        assert_eq!(row["x:note"], "A & B");
    }
    #[test]
    fn default_hidden_rows_and_zero_default_geometry_become_visible() {
        let e = element(
            "sheetFormatPr zeroHeight=\"1\" defaultRowHeight=\"0\" defaultColWidth=\"0\" baseColWidth=\"8\"",
        );
        assert_eq!(defaults(&e).unwrap(), (15.0, 8.43));
        let a = values(visible(&e, b"sheetFormatPr", 15.0, 8.43).unwrap());
        assert_eq!(a["zeroHeight"], "0");
        assert_eq!(a["defaultRowHeight"], "15");
        assert_eq!(a["defaultColWidth"], "8.43");
    }
    #[test]
    fn ambiguous_flags_and_invalid_dimensions_refuse() {
        assert!(visible(&element("row r=\"2\" hidden=\"maybe\""), b"row", 15.0, 8.43).is_err());
        for n in ["-1", "NaN", "inf"] {
            assert!(visible(&element(&format!("col width=\"{n}\"")), b"col", 15.0, 8.43).is_err());
        }
    }
}
