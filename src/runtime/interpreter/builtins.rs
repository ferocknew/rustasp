//! 内置函数模块
//!
//! 实现 VBScript 内置函数，如类型转换、字符串处理等

use crate::runtime::{RuntimeError, Value, ValueConversion};

/// 获取配置的日期格式字符串
fn get_datetime_format() -> (String, String, String) {
    // 从环境变量读取配置，使用默认值
    let now_format = std::env::var("NOW_FORMAT")
        .unwrap_or_else(|_| "YYYY/MM/DD HH:MM:SS".to_string());
    let date_format = std::env::var("DATE_FORMAT")
        .unwrap_or_else(|_| "YYYY/MM/DD".to_string());
    let time_format = std::env::var("TIME_FORMAT")
        .unwrap_or_else(|_| "HH:MM:SS".to_string());

    // 转换格式字符串：YYYY/MM/DD HH:MM:SS -> %Y/%m/%d %H:%M:%S
    let now_format_str = convert_vbscript_format(&now_format);
    let date_format_str = convert_vbscript_format(&date_format);
    let time_format_str = convert_vbscript_format(&time_format);

    (now_format_str, date_format_str, time_format_str)
}

/// 将 VBScript 日期格式转换为 strftime 格式
fn convert_vbscript_format(format: &str) -> String {
    format
        .replace("YYYY", "%Y")
        .replace("YY", "%y")
        .replace("MM", "%m")
        .replace("DD", "%d")
        .replace("HH", "%H")
        .replace("MM", "%M")  // 分钟，注意会与月份冲突，需要先替换月份
        .replace("SS", "%S")
}

/// 调用内置函数（多参数版本）
pub fn call_builtin_function_multi(name: &str, args: &[Value]) -> Option<Result<Value, RuntimeError>> {
    let name_lower = name.to_lowercase();

    match name_lower.as_str() {
        // 日期时间函数（无参数）
        "now" => {
            let now = chrono::Local::now();
            let (now_format, _, _) = get_datetime_format();
            Some(Ok(Value::String(now.format(&now_format).to_string())))
        }
        "date" => {
            let now = chrono::Local::now();
            let (_, date_format, _) = get_datetime_format();
            Some(Ok(Value::String(now.format(&date_format).to_string())))
        }
        "time" => {
            let now = chrono::Local::now();
            let (_, _, time_format) = get_datetime_format();
            Some(Ok(Value::String(now.format(&time_format).to_string())))
        }
        // 字符串函数 - 多参数
        "left" => {
            if args.len() >= 2 {
                let s = ValueConversion::to_string(&args[0]);
                let n = ValueConversion::to_number(&args[1]) as usize;
                let result = s.chars().take(n).collect::<String>();
                Some(Ok(Value::String(result)))
            } else {
                None
            }
        }
        "right" => {
            if args.len() >= 2 {
                let s = ValueConversion::to_string(&args[0]);
                let n = ValueConversion::to_number(&args[1]) as usize;
                let result = s.chars().rev().take(n).collect::<String>()
                    .chars().rev().collect::<String>();
                Some(Ok(Value::String(result)))
            } else {
                None
            }
        }
        "mid" => {
            if args.len() >= 2 {
                let s = ValueConversion::to_string(&args[0]);
                let start = (ValueConversion::to_number(&args[1]) as usize).saturating_sub(1);
                let length = if args.len() >= 3 {
                    ValueConversion::to_number(&args[2]) as usize
                } else {
                    s.len()
                };
                let result = s.chars().skip(start).take(length).collect::<String>();
                Some(Ok(Value::String(result)))
            } else {
                None
            }
        }
        "round" => {
            if args.len() >= 1 {
                let n = ValueConversion::to_number(&args[0]);
                let decimals = if args.len() >= 2 {
                    ValueConversion::to_number(&args[1]) as i32
                } else {
                    0
                };
                let multiplier = 10_f64.powi(decimals);
                Some(Ok(Value::Number((n * multiplier).round() / multiplier)))
            } else {
                None
            }
        }
        // 单参数函数 - 委托给旧版本
        _ => {
            if args.len() == 1 {
                call_builtin_function(name, &args[0])
            } else {
                None
            }
        }
    }
}

/// 调用内置函数
///
/// 处理类似 CInt(x) 的单参数内置函数调用
pub fn call_builtin_function(name: &str, arg: &Value) -> Option<Result<Value, RuntimeError>> {
    match name.to_lowercase().as_str() {
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
