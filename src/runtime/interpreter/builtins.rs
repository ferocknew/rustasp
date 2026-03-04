//! 内置函数模块
//!
//! 实现 VBScript 内置函数，如类型转换、字符串处理等

use crate::runtime::{RuntimeError, Value, ValueConversion};

/// 调用内置函数
///
/// 处理类似 CInt(x) 的单参数内置函数调用
pub fn call_builtin_function(name: &str, arg: &Value) -> Option<Result<Value, RuntimeError>> {
    match name {
        // 类型转换函数
        "cint" | "cbyte" | "cbool" => {
            Some(Ok(Value::Number(ValueConversion::to_number(arg) as i32 as f64)))
        }
        "clng" | "csng" => Some(Ok(Value::Number(ValueConversion::to_number(arg)))),
        "cdbl" => Some(Ok(Value::Number(ValueConversion::to_number(arg)))),
        "cstr" => Some(Ok(Value::String(ValueConversion::to_string(arg)))),
        "cdate" => {
            // TODO: 实现日期转换
            Some(Ok(arg.clone()))
        }
        "int" | "fix" => {
            let n = ValueConversion::to_number(arg);
            Some(Ok(Value::Number(n.trunc())))
        }
        "abs" => {
            let n = ValueConversion::to_number(arg);
            Some(Ok(Value::Number(n.abs())))
        }
        "sgn" => {
            let n = ValueConversion::to_number(arg);
            Some(Ok(Value::Number(if n > 0.0 {
                1.0
            } else if n < 0.0 {
                -1.0
            } else {
                0.0
            })))
        }
        "sqr" => {
            let n = ValueConversion::to_number(arg);
            Some(Ok(Value::Number(n.sqrt())))
        }
        "len" => {
            let s = ValueConversion::to_string(arg);
            Some(Ok(Value::Number(s.len() as f64)))
        }
        "trim" | "ltrim" | "rtrim" => {
            let s = ValueConversion::to_string(arg);
            let result = match name {
                "trim" => s.trim().to_string(),
                "ltrim" => s.trim_start().to_string(),
                "rtrim" => s.trim_end().to_string(),
                _ => s,
            };
            Some(Ok(Value::String(result)))
        }
        "ucase" | "lcase" => {
            let s = ValueConversion::to_string(arg);
            let result = match name {
                "ucase" => s.to_uppercase(),
                "lcase" => s.to_lowercase(),
                _ => s,
            };
            Some(Ok(Value::String(result)))
        }
        "chr" => {
            let n = ValueConversion::to_number(arg) as u32;
            Some(Ok(Value::String(
                char::from_u32(n).unwrap_or('\0').to_string(),
            )))
        }
        "asc" => {
            let s = ValueConversion::to_string(arg);
            let code = s.chars().next().map(|c| c as u8 as f64).unwrap_or(0.0);
            Some(Ok(Value::Number(code)))
        }
        "isnumeric" => {
            let is_num = match arg {
                Value::Number(_) => true,
                Value::Boolean(_) => true,
                Value::String(s) => s.parse::<f64>().is_ok(),
                Value::Empty => true,
                Value::Null => false,
                Value::Nothing => false,
                Value::Array(_) => false,
                Value::Object(_) => false,
            };
            Some(Ok(Value::Boolean(is_num)))
        }
        "isempty" => Some(Ok(Value::Boolean(matches!(arg, Value::Empty)))),
        "isnull" => Some(Ok(Value::Boolean(matches!(arg, Value::Null)))),
        "isarray" => Some(Ok(Value::Boolean(matches!(arg, Value::Array(_))))),
        "isobject" => Some(Ok(Value::Boolean(matches!(arg, Value::Object(_))))),
        "isdate" => {
            // TODO: 实现日期检测
            Some(Ok(Value::Boolean(false)))
        }
        _ => None,
    }
}
