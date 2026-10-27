//! Bounded numeric scalar dates, qualified against paired Excel 16.113.3 files.
//! Integer civil-day arithmetic never consults a clock, locale or timezone.
use crate::Result;

fn civil_days(mut year: i64, month: i64, day: i64) -> i64 {
    year -= i64::from(month <= 2);
    let era = year.div_euclid(400);
    let yoe = year - era * 400;
    let mp = month + if month > 2 { -3 } else { 9 };
    era * 146097 + yoe * 365 + yoe / 4 - yoe / 100 + (153 * mp + 2) / 5 + day - 1 - 719468
}
fn civil_parts(days: i64) -> (i64, i64, i64) {
    let z = days + 719468;
    let era = z.div_euclid(146097);
    let doe = z - era * 146097;
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    let year = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let day = doy - (153 * mp + 2) / 5 + 1;
    let month = mp + if mp < 10 { 3 } else { -9 };
    (year + i64::from(month <= 2), month, day)
}
pub(super) fn calculate(name: &str, values: &[f64], date1904: bool) -> Result<f64> {
    if values.len()
        != if name == "DATE" {
            3
        } else if ["EDATE", "EOMONTH"].contains(&name) {
            2
        } else {
            1
        }
        || values.iter().any(|value| !value.is_finite())
    {
        return Err(crate::coded(
            "formula-bad-input",
            "Date functions require finite numeric scalar inputs",
        ));
    }
    let origin1900 = civil_days(1899, 12, 31);
    let origin1904 = civil_days(1904, 1, 1);
    let origin = if date1904 { origin1904 } else { origin1900 };
    let maximum = if date1904 { 2957003 } else { 2958465 };
    if name == "DATE" {
        let year = values[0].floor();
        if !(0.0..=9999.0).contains(&year) {
            return Err("Excel error: #NUM!".into());
        }
        let month = values[1].floor();
        let day = values[2].floor();
        if month.abs() > 1200.0 || day.abs() > 36600.0 {
            return Err(crate::coded(
                "formula-limit",
                "Date normalization exceeds qualified month/day bounds",
            ));
        }
        let mut year = year as i64;
        if year <= 1899 {
            year += 1900;
        }
        let mut month = month as i64;
        let shift = (month - 1).div_euclid(12);
        year += shift;
        month -= shift * 12;
        let first = civil_days(year, month, 1);
        let serial =
            first - origin + i64::from(!date1904 && first >= civil_days(1900, 3, 1)) + day as i64
                - 1;
        return if serial < 0 || serial > maximum {
            Err("Excel error: #NUM!".into())
        } else {
            Ok(serial as f64)
        };
    }
    let serial = values[0];
    if serial < 0.0 || serial >= (maximum + 1) as f64 {
        return Err("Excel error: #NUM!".into());
    }
    let whole = serial.floor() as i64;
    let parts = if !date1904 && whole == 0 {
        (1900, 1, 0)
    } else if !date1904 && whole == 60 {
        (1900, 2, 29)
    } else {
        civil_parts(origin + whole - i64::from(!date1904 && whole > 60))
    };
    if name == "EDATE" || name == "EOMONTH" {
        let months = values[1].trunc();
        if months.abs() > 120000.0 {
            return Err(crate::coded(
                "formula-limit",
                "Month shift exceeds the qualified bound",
            ));
        }
        let months = months as i64;
        if !date1904 && whole == 0 && months == -1 {
            return Err("Excel error: #NUM!".into());
        }
        if !date1904
            && whole == 0
            && !(name == "EDATE" && [0, 1].contains(&months) || name == "EOMONTH" && months == 0)
            || !date1904
                && whole == 60
                && !(name == "EDATE" && months == 1 || name == "EOMONTH" && months == 0)
        {
            return Err(crate::coded(
                "formula-refused",
                "Uncaptured fictional/origin month-date mode",
            ));
        }
        let (mut year, mut month, mut day) = parts;
        let shift = (month - 1 + months).div_euclid(12);
        year += shift;
        month = month - 1 + months - shift * 12 + 1;
        if !(1..=9999).contains(&year) {
            return Err("Excel error: #NUM!".into());
        }
        let last = civil_days(year, month + 1, 1) - civil_days(year, month, 1);
        day = if name == "EOMONTH" {
            last
        } else {
            day.min(last)
        };
        let target = civil_days(year, month, day);
        let result = target - origin + i64::from(!date1904 && target >= civil_days(1900, 3, 1));
        if result < 0 && !(date1904 && whole == 0 && months == -1) {
            return Err(crate::coded(
                "formula-refused",
                "Uncaptured month-date result before calendar origin",
            ));
        }
        return if result > maximum {
            Err("Excel error: #NUM!".into())
        } else {
            Ok(result as f64)
        };
    }
    match name {
        "YEAR" => Ok(parts.0 as f64),
        "MONTH" => Ok(parts.1 as f64),
        "DAY" => Ok(parts.2 as f64),
        _ => Err(crate::coded(
            "formula-unsupported",
            "Unqualified date function",
        )),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn month_functions_match_actual_paired_numeric_script_oracles() {
        for (system, raw) in [
            (
                false,
                include_str!("../../../fixtures/excel-month-1900-reference.json"),
            ),
            (
                true,
                include_str!("../../../fixtures/excel-month-1904-reference.json"),
            ),
        ] {
            let oracle: serde_json::Value = serde_json::from_str(raw).unwrap();
            for (reference, item) in oracle["cells"].as_object().unwrap() {
                let view = serde_json::json!({"revision":"monthly-native","sheets":[{"name":"Cases","date1904":system,"cells":{reference:{"formula":item["formula"]}}}]});
                let result = super::super::evaluate(&view);
                if let Some(expected) = item["expected"].as_f64() {
                    assert_eq!(
                        result.as_ref().unwrap()["results"][0]["value"].as_f64(),
                        Some(expected),
                        "{system}:{reference}"
                    );
                } else if let Ok(value) = result {
                    assert_eq!(
                        value["results"][0]["error"], item["observedError"],
                        "{system}:{reference}"
                    );
                }
            }
        }
        assert!(
            calculate("EDATE", &[60.0, 0.0], false)
                .unwrap_err()
                .contains("Uncaptured")
        );
        assert!(
            calculate("EDATE", &[30.0, -2.0], true)
                .unwrap_err()
                .contains("Uncaptured")
        );
    }
    #[test]
    fn origins_fictional_day_and_bounds_are_explicit() {
        assert_eq!(calculate("DATE", &[1900.0, 3.0, 0.0], false).unwrap(), 60.0);
        assert_eq!(calculate("DAY", &[60.75], false).unwrap(), 29.0);
        assert_eq!(calculate("DAY", &[0.5], false).unwrap(), 0.0);
        assert_eq!(calculate("YEAR", &[0.5], false).unwrap(), 1900.0);
        assert_eq!(calculate("DATE", &[1904.0, 1.0, 1.0], true).unwrap(), 0.0);
        assert_eq!(calculate("DAY", &[0.5], true).unwrap(), 1.0);
        assert_eq!(
            calculate("DATE", &[9999.0, 12.0, 31.0], true).unwrap(),
            2957003.0
        );
        assert!(
            calculate("DATE", &[9999.0, 12.0, 32.0], false)
                .unwrap_err()
                .contains("#NUM!")
        );
        assert!(
            calculate("DATE", &[2024.0, 1201.0, 1.0], false)
                .unwrap_err()
                .contains("bounds")
        );
    }
}
