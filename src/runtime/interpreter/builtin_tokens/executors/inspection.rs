//! 检验函数执行器

use crate::runtime::{RuntimeError, Value};
use super::super::token::BuiltinToken;
use chrono::Timelike;

pub fn execute(token: BuiltinToken, args: &[Value]) -> Result<Option<Value>, RuntimeError> {
    let result = match token {
        BuiltinToken::IsNumeric => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            Value::Boolean(is_numeric(&args[0]))
        }
        BuiltinToken::IsEmpty => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            Value::Boolean(matches!(args[0], Value::Empty))
        }
        BuiltinToken::IsNull => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            Value::Boolean(matches!(args[0], Value::Null))
        }
        BuiltinToken::IsObject => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            Value::Boolean(matches!(args[0], Value::Object(_)))
        }
        BuiltinToken::IsDate => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            // 直接实现，避免函数调用问题
            let result = match &args[0] {
                Value::String(s) => {
                    let trimmed = s.trim();
                    if trimmed.is_empty() {
                        false
                    } else {
                        try_parse_date(trimmed).is_some() || try_parse_time(trimmed).is_some()
                    }
                }
                Value::Number(n) => {
                    *n >= -657434.0 && *n <= 2958465.0
                }
                _ => false,
            };
            Value::Boolean(result)
        }
        BuiltinToken::VarType => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            Value::Number(var_type(&args[0]) as f64)
        }
        BuiltinToken::TypeName => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            Value::String(type_name(&args[0]).to_string())
        }
        _ => return Ok(None),
    };
    Ok(Some(result))
}

#[allow(dead_code)]
fn is_numeric(v: &Value) -> bool {
    matches!(v, Value::Number(_) | Value::Boolean(_) | Value::Empty)
        || matches!(v, Value::String(s) if s.parse::<f64>().is_ok())
}

#[allow(dead_code)]
fn var_type(v: &Value) -> u16 {
    match v {
        Value::Empty => 0,
        Value::Null => 1,
        Value::Number(_) => 2,
        Value::String(_) => 8,
        Value::Boolean(_) => 11,
        Value::Array(_) => 8192,
        Value::Object(_) => 9,
        Value::Nothing => 1,
    }
}

#[allow(dead_code)]
fn type_name(v: &Value) -> &'static str {
    match v {
        Value::Empty => "Empty",
        Value::Null => "Null",
        Value::Number(_) => "Double",
        Value::String(_) => "String",
        Value::Boolean(_) => "Boolean",
        Value::Array(_) => "Variant()",
        Value::Object(_) => "Object",
        Value::Nothing => "Nothing",
    }
}

/// 尝试解析日期字符串
/// 支持多种日期格式，与 datetime.rs 中的 parse_vbscript_date 类似
fn try_parse_date(date_str: &str) -> Option<chrono::NaiveDateTime> {
    use chrono::NaiveDate;

    // 移除 # 包围（VBScript 日期字面量 #2024-01-01#）
    let clean_str = if date_str.starts_with('#') && date_str.ends_with('#') {
        &date_str[1..date_str.len()-1]
    } else {
        date_str
    };

    let clean_str = clean_str.trim();

    // 尝试各种日期格式
    let date_formats = [
        // ISO 格式
        "%Y-%m-%d",
        "%Y/%m/%d",
        // 美式格式
        "%m/%d/%Y",
        "%m-%d-%Y",
        // 短年份格式
        "%y-%m-%d",
        "%y/%m/%d",
    ];

    // 先尝试纯日期
    for format in &date_formats {
        if let Ok(date) = NaiveDate::parse_from_str(clean_str, format) {
            return Some(date.and_hms_opt(0, 0, 0).unwrap_or_else(|| {
                chrono::NaiveDateTime::new(date, chrono::NaiveTime::from_hms_opt(0, 0, 0).unwrap())
            }));
        }
    }

    // 尝试日期时间格式（包含时间部分）
    let datetime_formats = [
        "%Y-%m-%d %H:%M:%S",
        "%Y/%m/%d %H:%M:%S",
        "%m/%d/%Y %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y/%m/%d %H:%M",
        "%m/%d/%Y %H:%M",
    ];

    for format in &datetime_formats {
        if let Ok(dt) = chrono::NaiveDateTime::parse_from_str(clean_str, format) {
            return Some(dt);
        }
    }

    None
}

/// 尝试解析时间字符串
/// 支持多种时间格式，包括 AM/PM 标记
fn try_parse_time(time_str: &str) -> Option<chrono::NaiveDateTime> {
    let trimmed = time_str.trim().to_lowercase();

    // 处理 AM/PM 标记
    let (time_part, is_pm, has_period) = if trimmed.ends_with("pm") {
        (trimmed[..trimmed.len()-2].trim().to_string(), true, true)
    } else if trimmed.ends_with("am") {
        (trimmed[..trimmed.len()-2].trim().to_string(), false, true)
    } else {
        (trimmed.to_string(), false, false)
    };

    // 尝试各种时间格式
    let time_formats = [
        "%H:%M:%S",    // 14:30:45
        "%H:%M",       // 14:30
        "%I:%M:%S",    // 2:30:45 (12小时制)
        "%I:%M",       // 2:30 (12小时制)
        "%H:%M:%S %.f", // 带毫秒
    ];

    for format in &time_formats {
        if let Ok(mut time) = chrono::NaiveTime::parse_from_str(&time_part, format) {
            // 如果是 12 小时制且是 PM，且小时不是 12
            if has_period && is_pm && time.hour() < 12 {
                time = chrono::NaiveTime::from_hms_opt(
                    time.hour() + 12,
                    time.minute(),
                    time.second()
                ).unwrap_or(time);
            }
            // 如果是 12 小时制且是 AM，且小时是 12
            if has_period && !is_pm && time.hour() == 12 {
                time = chrono::NaiveTime::from_hms_opt(
                    0,
                    time.minute(),
                    time.second()
                ).unwrap_or(time);
            }
            // 将时间转换为日期时间（使用一个基准日期）
            let date = chrono::NaiveDate::from_ymd_opt(1899, 12, 30)?;
            return Some(chrono::NaiveDateTime::new(date, time));
        }
    }

    None
}
