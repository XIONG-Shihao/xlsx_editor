//! Closed, explicit series fills/lines; style identity comes from saved DrawingML.
use super::managed_shape::same;
use super::print_settings::Tree;
use super::*;
const COLORFUL: [&str; 8] = [
    "2563EB", "E86125", "16A34A", "9333EA", "CA8A04", "0891B2", "DC2626", "64748B",
];
const MONO: [&str; 8] = [
    "217346", "42875F", "66A27F", "85B998", "175431", "327149", "50946A", "749F84",
];
fn expected(style: &str, index: usize) -> String {
    let color = if style == "monochrome" {
        MONO[index]
    } else {
        COLORFUL[index]
    };
    let fill = if style == "outline" {
        "<a:noFill/>".to_string()
    } else {
        format!("<a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill>")
    };
    let dash = if style == "outline" { "dash" } else { "solid" };
    format!(
        "<c:spPr xmlns:c=\"{}\" xmlns:a=\"{}\">{fill}<a:ln w=\"19050\"><a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill><a:prstDash val=\"{dash}\"/></a:ln></c:spPr>",
        String::from_utf8_lossy(CHART_NS),
        String::from_utf8_lossy(A_NS)
    )
}
pub(super) fn check(bytes: &[u8]) -> Result<Option<&'static str>> {
    let t = Tree::parse(bytes)?;
    t.root(CHART_NS, "chartSpace")?;
    let props = t
        .nodes
        .iter()
        .enumerate()
        .filter(|(_, n)| n.ns == CHART_NS && n.name == "spPr")
        .map(|(i, _)| i)
        .collect::<Vec<_>>();
    if props.is_empty() {
        if t.nodes.iter().any(|n| {
            n.ns == A_NS
                && ["solidFill", "srgbClr", "noFill", "ln", "prstDash"].contains(&n.name.as_str())
        }) {
            return Err(crate::coded(
                "chart-protected",
                "Chart formatting outside series is protected",
            ));
        }
        return Ok(None);
    }
    let series = t
        .nodes
        .iter()
        .enumerate()
        .filter(|(_, n)| n.ns == CHART_NS && n.name == "ser")
        .map(|(i, _)| i)
        .collect::<Vec<_>>();
    if series.is_empty() || series.len() > 8 || props.len() != series.len() {
        return Err(crate::coded(
            "chart-protected",
            "Incomplete/foreign chart series formatting is protected",
        ));
    }
    for style in ["colorful", "monochrome", "outline"] {
        let mut valid = true;
        for (index, i) in series.iter().enumerate() {
            let own = t.nodes[*i]
                .children
                .iter()
                .copied()
                .filter(|n| t.nodes[*n].ns == CHART_NS && t.nodes[*n].name == "spPr")
                .collect::<Vec<_>>();
            let expected = Tree::parse(expected(style, index).as_bytes())?;
            if own.len() != 1 || !same(&t, own[0], &expected, 0) {
                valid = false;
                break;
            }
        }
        if valid {
            return Ok(Some(style));
        }
    }
    Err(crate::coded(
        "chart-protected",
        "Foreign/extended chart preset is protected",
    ))
}
#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn admits_three_real_series_presets_and_protects_modified_properties() {
        for style in ["colorful", "monochrome", "outline"] {
            let series = (0..8)
                .map(|i| format!("<c:ser>{}</c:ser>", expected(style, i)))
                .collect::<String>();
            let xml = format!(
                "<c:chartSpace xmlns:c=\"{}\"><c:chart><c:plotArea><c:barChart>{series}</c:barChart></c:plotArea></c:chart></c:chartSpace>",
                String::from_utf8_lossy(CHART_NS)
            );
            assert_eq!(check(xml.as_bytes()).unwrap(), Some(style));
            for changed in [
                xml.replace("19050", "19051"),
                xml.replace("srgbClr", "schemeClr"),
                xml.replace("prstDash", "custDash"),
            ] {
                assert!(check(changed.as_bytes()).is_err());
            }
        }
    }
    #[test]
    fn preserves_unstyled_profile_and_refuses_partial_or_root_formatting() {
        let c = String::from_utf8_lossy(CHART_NS);
        let a = String::from_utf8_lossy(A_NS);
        assert_eq!(
            check(format!("<c:chartSpace xmlns:c=\"{c}\"><c:chart/></c:chartSpace>").as_bytes())
                .unwrap(),
            None
        );
        assert!(
            check(
                format!(
                    "<c:chartSpace xmlns:c=\"{c}\" xmlns:a=\"{a}\"><a:solidFill/></c:chartSpace>"
                )
                .as_bytes()
            )
            .is_err()
        );
        assert!(
            check(
                format!(
                    "<c:chartSpace xmlns:c=\"{c}\"><c:ser>{}</c:ser><c:ser/></c:chartSpace>",
                    expected("colorful", 0)
                )
                .as_bytes()
            )
            .is_err()
        );
    }
}
