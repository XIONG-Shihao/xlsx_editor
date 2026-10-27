use super::*;
use serde_json::json;

#[test]
fn custom_format_independent_documentation_corpus_has_exact_text_color_and_admission() {
    let corpus: Value = serde_json::from_str(include_str!(
        "../../../../tests/support/custom-format-vectors.json"
    ))
    .unwrap();
    assert_eq!(
        corpus["schema"],
        workbook_format!(SCHEMA_CUSTOM_FORMAT_VECTORS_V1)
    );
    assert_eq!(corpus["source"], "documentation-derived");
    let mut ids = std::collections::BTreeSet::new();
    for case in corpus["cases"].as_array().unwrap() {
        assert!(ids.insert(case["id"].as_str().unwrap()));
        let before = case["value"].clone();
        let actual = format_custom_number(
            &case["value"],
            case["format"].as_str().unwrap(),
            case["date1904"].as_bool().unwrap_or(false),
        )
        .unwrap();
        let actual = actual.map_or(Value::Null, |result| serde_json::to_value(result).unwrap());
        assert_eq!(actual, case["expected"], "{}", case["id"]);
        assert_eq!(case["value"], before);
    }
    assert_eq!(ids.len(), 99);
}
#[test]
fn custom_format_resource_bounds_scalar_fences_and_unicode_are_explicit() {
    assert!(
        format_custom_number(&json!(7), &"0".repeat(CODE_UNITS + 1), false)
            .unwrap()
            .is_none()
    );
    assert!(
        format_custom_number(
            &json!(7),
            &format!("0.{}", "0".repeat(FRACTION_PLACES + 1)),
            false
        )
        .unwrap()
        .is_none()
    );
    assert!(
        format_custom_number(&json!("x".repeat(OUTPUT_UNITS + 1)), "@", false)
            .unwrap()
            .is_none()
    );
    for value in [Value::Null, json!(true), json!({}), json!([])] {
        assert!(
            format_custom_number(&value, "0.00", false)
                .unwrap()
                .is_none()
        );
    }
    for code in ["", "0\u{0}", "0\"\u{fffe}\"", "0\\\u{ffff}"] {
        assert!(
            format_custom_number(&json!(7), code, false)
                .unwrap()
                .is_none()
        );
    }
    assert!(serde_json::from_str::<Value>(r#""\ud800""#).is_err());
    let mixed = format_custom_number(&json!(8.9), "0.#0", false)
        .unwrap()
        .unwrap();
    assert_eq!(mixed.text, "8.90");
    let unicode = format_custom_number(&json!("雪 😀"), "0;0;0;\"day: \"@", false)
        .unwrap()
        .unwrap();
    assert_eq!(unicode.text, "day: 雪 😀");
    let negative_zero = json!(-0.0_f64);
    assert!(negative_zero.as_f64().unwrap().is_sign_negative());
    assert_eq!(
        format_custom_number(&negative_zero, "$#,##0.0000;[Red]($#,##0.0000)", false)
            .unwrap()
            .unwrap()
            .text,
        "$0.0000"
    );
    assert!(negative_zero.as_f64().unwrap().is_sign_negative());
}
#[test]
fn custom_format_repeated_text_refuses_before_an_oversized_output_append() {
    let value = json!("x".repeat(32767));
    let before = value.clone();
    assert!(
        format_custom_number(&value, &"@".repeat(CODE_UNITS), false)
            .unwrap()
            .is_none()
    );
    let admitted = format_custom_number(&value, "@@", false).unwrap().unwrap();
    assert_eq!(admitted.text.encode_utf16().count(), 65534);
    let exact = format_custom_number(&value, "@@\"..\"", false)
        .unwrap()
        .unwrap();
    assert_eq!(exact.text.encode_utf16().count(), OUTPUT_UNITS);
    assert!(
        format_custom_number(&value, "@@\"...\"", false)
            .unwrap()
            .is_none()
    );
    assert_eq!(value, before);
    let unicode = json!(format!("{}x", "😀".repeat(16383)));
    assert_eq!(unicode.as_str().unwrap().encode_utf16().count(), 32767);
    assert_eq!(
        format_custom_number(&unicode, "@@", false)
            .unwrap()
            .unwrap()
            .text
            .encode_utf16()
            .count(),
        65534
    );
    assert!(
        format_custom_number(&unicode, &"@".repeat(CODE_UNITS), false)
            .unwrap()
            .is_none()
    );
}
#[test]
fn custom_format_legacy_rounding_and_builtin_date_contract_stay_distinct() {
    for (value, format, expected) in [
        (1.005, 2, "1.00"),
        (2.675, 2, "2.67"),
        (0.615, 2, "0.61"),
        (-2.5, 1, "-3"),
        (2.675, 4, "2.68"),
        (-0.0, 4, "-0.00"),
    ] {
        assert_eq!(
            super::super::formatted(value, &json!(format), false).unwrap(),
            expected
        );
    }
    for (serial, calendar, expected) in [
        (0.0, true, "1904-01-01"),
        (59.0, false, "1900-02-28"),
        (60.0, false, "1900-02-29"),
        (61.0, false, "1900-03-01"),
        (-1.0, false, "1899-12-30"),
    ] {
        assert_eq!(
            super::super::formatted(serial, &json!(14), calendar).unwrap(),
            expected
        );
    }
    assert_eq!(
        super::super::formatted(7.0, &json!("14"), false).unwrap(),
        "7"
    );
    assert_eq!(
        super::super::formatted(7.0, &json!("[Red]0.00"), false).unwrap(),
        "7.00"
    );
}
#[test]
fn custom_format_cell_authority_precedes_typed_text_sections_and_named_colors() {
    let code = json!("0.00;[Red](0.00);\"zero\";\"text: \"@");
    let display = |cell: &Value| super::super::cell(cell, &code, false).unwrap();
    assert_eq!(
        display(&json!({"value":"#N/A","dataType":"text"})),
        "text: #N/A"
    );
    assert_eq!(display(&json!({"value":"#N/A","dataType":"error"})), "#N/A");
    assert_eq!(display(&json!({"value":"#N/A","error":"#N/A"})), "#N/A");
    assert_eq!(
        display(
            &json!({"value":9007199254740992_f64,"rawValue":"9007199254740993","dataType":"number"})
        ),
        "9007199254740993"
    );
    assert_eq!(display(&json!({"value":null})), "");
    assert_eq!(display(&json!({"value":"","dataType":"text"})), "text: ");
    assert_eq!(
        display(&json!({"formula":"1","calculation":{"status":"calculated","value":"June"}})),
        "text: June"
    );
    assert_eq!(
        display(&json!({"formula":"1","calculation":{"status":"calculated","value":null}})),
        ""
    );
    assert_eq!(
        display(&json!({"formula":"1","calculation":{"status":"error","error":"#N/A"}})),
        "#N/A"
    );
    assert_eq!(
        display(&json!({"formula":"1","calculation":{"status":"circular"}})),
        "#CYCLE!"
    );
    assert_eq!(
        display(&json!({"formula":"1","calculation":{"status":"unresolved"}})),
        "Unresolved"
    );
    assert_eq!(display(&json!({"formula":"1"})), "Calculating…");
    assert_eq!(
        display(
            &json!({"formula":"1","calculation":{"pending":true},"cachedResult":{"status":"calculated","value":"June"}})
        ),
        "Stored: text: June"
    );
    assert_eq!(
        display(
            &json!({"formula":"1","calculation":{"pending":true},"cachedResult":{"status":"error","error":"#N/A"}})
        ),
        "Stored: #N/A"
    );
    assert_eq!(
        display(
            &json!({"formula":"1","calculation":{"pending":true},"cachedResult":{"status":"error"}})
        ),
        "Stored: #ERROR!"
    );
    assert_eq!(
        display(
            &json!({"formula":"1","calculation":{"pending":true},"cachedResult":{"status":"calculated","value":9007199254740992_f64,"rawValue":"9007199254740993"}})
        ),
        "Stored: 9007199254740993"
    );
    assert_eq!(
        display(&json!({"formula":"1","calculation":{"pending":true}})),
        "Pending calculation"
    );
}
