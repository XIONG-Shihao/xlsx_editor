//! Pinned evaluator arithmetic contract: safe-integer IEEE shortcut, otherwise
//! decimal operands from their shortest spellings, 20 significant digits with
//! half-up rounding, then one correctly-rounded conversion back to binary64.
use crate::Result;
use num_bigint::BigInt;
use num_rational::BigRational;
use num_traits::{Signed, Zero};

pub fn number(n: f64) -> Result<f64> {
    if !n.is_finite() {
        return Err("Excel error: #NUM!".into());
    }
    if n == 0.0 && n.is_sign_negative() {
        return Err(crate::coded(
            "formula-bad-input",
            "Negative zero requires compatibility engine",
        ));
    }
    Ok(n)
}
fn ten(n: u32) -> BigInt {
    BigInt::from(10).pow(n)
}
fn decimal(n: f64) -> Result<BigRational> {
    number(n)?;
    let text = n.to_string();
    let (parts, exp) = if let Some((a, b)) = text.split_once(['e', 'E']) {
        (a, b.parse::<i32>().map_err(crate::err)?)
    } else {
        (text.as_str(), 0)
    };
    let decimals = parts
        .split_once('.')
        .map(|(_, b)| b.len() as i32)
        .unwrap_or(0);
    let digits = parts
        .replace('.', "")
        .parse::<BigInt>()
        .map_err(crate::err)?;
    let power = exp - decimals;
    Ok(if power >= 0 {
        BigRational::from_integer(digits * ten(power as u32))
    } else {
        BigRational::new(digits, ten((-power) as u32))
    })
}
fn rounded(n: BigRational) -> Result<f64> {
    rounded_precision(n, 20)
}
/// JavaScript Number.toPrecision(15) rounds the exact binary64 magnitude,
/// choosing the larger decimal on a tie. Rust formatting uses ties-to-even.
pub fn series_precision(n: f64) -> Result<f64> {
    rounded_precision(
        BigRational::from_float(n)
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid series number"))?,
        15,
    )
}
fn rounded_precision(n: BigRational, precision: i32) -> Result<f64> {
    if n.is_zero() {
        return Ok(0.0);
    }
    let negative = n.is_negative();
    let numerator = n.numer().abs();
    let denominator = n.denom();
    let mut exponent = numerator.to_string().len() as i32 - denominator.to_string().len() as i32;
    let below = if exponent >= 0 {
        numerator < denominator * ten(exponent as u32)
    } else {
        &numerator * ten((-exponent) as u32) < *denominator
    };
    if below {
        exponent -= 1
    }
    let shift = precision - 1 - exponent;
    let (a, b) = if shift >= 0 {
        (numerator * ten(shift as u32), denominator.clone())
    } else {
        (numerator, denominator * ten((-shift) as u32))
    };
    let mut quotient = &a / &b;
    let remainder = a % &b;
    if remainder * 2 >= b {
        quotient += 1
    }
    let text = format!("{}{quotient}e{}", if negative { "-" } else { "" }, -shift);
    number(text.parse::<f64>().map_err(crate::err)?)
}
fn arithmetic(a: f64, b: f64, op: u8) -> Result<f64> {
    number(a)?;
    number(b)?;
    if op == b'/' && b == 0.0 {
        return Err("Excel error: #DIV/0!".into());
    }
    let raw = match op {
        b'+' => a + b,
        b'-' => a - b,
        b'*' => a * b,
        b'/' => a / b,
        _ => unreachable!(),
    };
    if raw.is_finite() && raw.fract() == 0.0 && raw.abs() <= 9007199254740991.0 {
        return number(raw);
    }
    let a = decimal(a)?;
    let b = decimal(b)?;
    rounded(match op {
        b'+' => a + b,
        b'-' => a - b,
        b'*' => a * b,
        b'/' => a / b,
        _ => unreachable!(),
    })
}
pub fn add(a: f64, b: f64) -> Result<f64> {
    arithmetic(a, b, b'+')
}
#[cfg(test)]
fn subtract(a: f64, b: f64) -> Result<f64> {
    arithmetic(a, b, b'-')
}
pub fn multiply(a: f64, b: f64) -> Result<f64> {
    arithmetic(a, b, b'*')
}
pub fn divide(a: f64, b: f64) -> Result<f64> {
    arithmetic(a, b, b'/')
}
pub fn round(value: f64, digits: f64) -> Result<f64> {
    number(value)?;
    number(digits)?;
    let digits = digits.trunc();
    if !(0.0..=12.0).contains(&digits) {
        return Err(crate::coded(
            "formula-bad-input",
            "ROUND precision requires compatibility engine",
        ));
    }
    let factor = 10_f64.powi(digits as i32);
    let scaled = multiply(value.abs(), factor)?;
    let adjusted = scaled + f64::EPSILON * scaled.abs().max(1.0);
    let low = adjusted.floor();
    let rounded = if adjusted - low >= 0.5 {
        low + 1.0
    } else {
        low
    };
    number(value.signum() * rounded / factor)
}
#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn decimal_contract() {
        assert_eq!(add(0.1, 0.2).unwrap(), 0.3);
        assert_eq!(
            subtract(1_000_000_000.0 + 2_f64.powi(-20), 1_000_000_000.0).unwrap(),
            0.000001
        );
        assert_eq!(divide(1.0, 3.0).unwrap(), 1.0 / 3.0);
        assert_eq!(round(-1.005, 2.0).unwrap(), -1.01);
        assert!(divide(1.0, 0.0).is_err());
    }
}
