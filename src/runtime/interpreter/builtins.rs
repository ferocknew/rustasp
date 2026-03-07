//! 内置函数模块
//!
//! 实现 VBScript 内置函数，如类型转换、字符串处理等

use crate::runtime::{RuntimeError, Value, ValueConversion};
use chrono::Datelike;
use chrono::Timelike;

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
        "year" => {
            let now = chrono::Local::now();
            Some(Ok(Value::Number(now.year() as f64)))
        }
        "month" => {
            let now = chrono::Local::now();
            Some(Ok(Value::Number(now.month() as f64)))
        }
        "day" => {
            let now = chrono::Local::now();
            Some(Ok(Value::Number(now.day() as f64)))
        }
        "hour" => {
            let now = chrono::Local::now();
            Some(Ok(Value::Number(now.hour() as f64)))
        }
        "minute" => {
            let now = chrono::Local::now();
            Some(Ok(Value::Number(now.minute() as f64)))
        }
        "second" => {
            let now = chrono::Local::now();
            Some(Ok(Value::Number(now.second() as f64)))
        }
        "weekday" => {
            let now = chrono::Local::now();
            // VBScript: 1=Sunday, 2=Monday, ..., 7=Saturday
            // chrono: 0=Monday, ..., 6=Sunday
            let weekday = now.weekday().num_days_from_sunday();
            Some(Ok(Value::Number((weekday + 1) as f64)))
        }
        "weekdayname" => {
            // WeekdayName(weekday, abbreviate, firstdayofweek)
            if args.len() < 1 {
                return None;
            }
            let weekday = ValueConversion::to_number(&args[0]) as usize % 7;
            let abbreviate = if args.len() >= 2 {
                ValueConversion::to_bool(&args[1])
            } else {
                false
            };

            // VBScript: 0=Sunday(default) or 1=Monday as first day
            // WeekdayName expects 1=Sunday, 2=Monday, ..., 7=Saturday
            let names = if abbreviate {
                ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
            } else {
                ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
            };

            let index = if weekday == 0 { 6 } else { weekday - 1 };
            let name = names.get(index).unwrap_or(&"");
            Some(Ok(Value::String(name.to_string())))
        }
        "monthname" => {
            // MonthName(month, abbreviate)
            if args.len() < 1 {
                return None;
            }
            let month = (ValueConversion::to_number(&args[0]) as usize - 1) % 12;
            let abbreviate = if args.len() >= 2 {
                ValueConversion::to_bool(&args[1])
            } else {
                false
            };

            let names = if abbreviate {
                ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
            } else {
                ["January", "February", "March", "April", "May", "June",
                 "July", "August", "September", "October", "November", "December"]
            };

            let name = names.get(month).unwrap_or(&"");
            Some(Ok(Value::String(name.to_string())))
        }
        "timer" => {
            // Timer 返回自午夜以来的秒数
            let now = chrono::Local::now();
            let seconds_since_midnight = now.hour() as f64 * 3600.0
                + now.minute() as f64 * 60.0
                + now.second() as f64
                + now.nanosecond() as f64 / 1_000_000_000.0;
            Some(Ok(Value::Number(seconds_since_midnight)))
        }
        "dateadd" => {
            // DateAdd(interval, number, date)
            if args.len() < 3 {
                return None;
            }
            let interval = ValueConversion::to_string(&args[0]).to_lowercase();
            let number = ValueConversion::to_number(&args[1]) as i64;

            // 简化实现：使用当前时间
            let mut date = chrono::Local::now();

            match interval.as_str() {
                "yyyy" => date = date + chrono::Duration::days(number * 365),
                "q" => date = date + chrono::Duration::days(number * 91),
                "m" => date = date + chrono::Duration::days(number * 30),
                "y" => date = date + chrono::Duration::days(number),
                "d" => date = date + chrono::Duration::days(number),
                "w" => date = date + chrono::Duration::days(number * 7),
                "ww" => date = date + chrono::Duration::days(number * 7),
                "h" => date = date + chrono::Duration::hours(number),
                "n" => date = date + chrono::Duration::minutes(number),
                "s" => date = date + chrono::Duration::seconds(number),
                _ => return Some(Ok(Value::Empty)),
            }

            let (now_format, _, _) = get_datetime_format();
            Some(Ok(Value::String(date.format(&now_format).to_string())))
        }
        "datediff" => {
            // DateDiff(interval, date1, date2)
            // 简化实现：返回两个日期之间的差值
            if args.len() < 3 {
                return None;
            }
            let interval = ValueConversion::to_string(&args[0]).to_lowercase();
            let date1 = chrono::Local::now();
            let date2 = chrono::Local::now();

            let diff = match interval.as_str() {
                "d" | "y" => (date2 - date1).num_days(),
                "h" => (date2 - date1).num_hours(),
                "n" => (date2 - date1).num_minutes(),
                "s" => (date2 - date1).num_seconds(),
                _ => 0,
            };

            Some(Ok(Value::Number(diff as f64)))
        }
        "datepart" => {
            // DatePart(interval, date)
            // 简化实现：返回日期的指定部分
            if args.len() < 2 {
                return None;
            }
            let interval = ValueConversion::to_string(&args[0]).to_lowercase();
            let date = chrono::Local::now();

            let result = match interval.as_str() {
                "yyyy" => date.year() as f64,
                "q" => ((date.month() - 1) / 3 + 1) as f64,
                "m" => date.month() as f64,
                "d" => date.day() as f64,
                "h" => date.hour() as f64,
                "n" => date.minute() as f64,
                "s" => date.second() as f64,
                _ => 0.0,
            };

            Some(Ok(Value::Number(result)))
        }
        // 随机数函数（无参数或单参数）
        "rnd" => {
            use std::time::{SystemTime, UNIX_EPOCH};
            let n = if args.len() >= 1 {
                ValueConversion::to_number(&args[0])
            } else {
                1.0 // 默认返回下一个随机数
            };
            let seed = if n < 0.0 {
                (n.abs() as u32) as u64
            } else {
                SystemTime::now().duration_since(UNIX_EPOCH)
                    .unwrap_or_default()
                    .as_nanos() as u64
            };
            // 使用更简单的随机数生成，避免溢出
            let random = (seed % 2147483647) as f64 / 2147483647.0;
            Some(Ok(Value::Number(random)))
        }
        "randomize" => {
            // Randomize 函数：初始化随机数生成器
            // 在当前实现中不需要做任何事
            Some(Ok(Value::Empty))
        }
        // 字符串函数 - 多参数
        "instr" => {
            // InStr([start, ]string1, string2[, compare])
            // 返回 string2 在 string1 中首次出现的位置
            let (start, string1, string2) = if args.len() >= 3 {
                // 有 start 参数
                (Some(ValueConversion::to_number(&args[0]) as usize),
                 ValueConversion::to_string(&args[1]),
                 ValueConversion::to_string(&args[2]))
            } else if args.len() == 2 {
                // 没有 start 参数
                (None, ValueConversion::to_string(&args[0]), ValueConversion::to_string(&args[1]))
            } else {
                return None;
            };

            // 从指定位置开始搜索（VBScript 位置从 1 开始）
            let search_str = if let Some(pos) = start {
                if pos > string1.len() {
                    return Some(Ok(Value::Number(0.0)));
                } else if pos > 1 {
                    &string1.chars().skip(pos - 1).collect::<String>()
                } else {
                    &string1
                }
            } else {
                &string1
            };

            // 查找子串
            let pos = search_str.find(&string2);
            let result = if let Some(p) = pos {
                // VBScript 位置从 1 开始
                let base_pos = start.unwrap_or(1);
                (base_pos + p) as f64
            } else {
                0.0
            };
            Some(Ok(Value::Number(result)))
        }
        "strcomp" => {
            // StrComp(string1, string2[, compare])
            // 比较两个字符串，返回 -1, 0, 或 1
            if args.len() < 2 {
                return None;
            }
            let string1 = ValueConversion::to_string(&args[0]);
            let string2 = ValueConversion::to_string(&args[1]);

            let result = if string1 == string2 {
                0.0
            } else if string1 < string2 {
                -1.0
            } else {
                1.0
            };
            Some(Ok(Value::Number(result)))
        }
        "string" => {
            // String(number, character) - 返回重复的字符
            if args.len() >= 2 {
                let n = ValueConversion::to_number(&args[0]) as usize;
                let ch = ValueConversion::to_string(&args[1]).chars().next().unwrap_or(' ');
                let result = ch.to_string().repeat(n.min(1000000));
                Some(Ok(Value::String(result)))
            } else {
                None
            }
        }
        "space" => {
            let n = if args.len() >= 1 {
                ValueConversion::to_number(&args[0]) as usize
            } else {
                0
            };
            let spaces = " ".repeat(n.min(1000000));
            Some(Ok(Value::String(spaces)))
        }
        "instrrev" => {
            // InStrRev(string1, string2[, start[, compare]])
            // 从字符串末尾开始查找
            let (string1, string2, start) = if args.len() >= 2 {
                let s1 = ValueConversion::to_string(&args[0]);
                let s2 = ValueConversion::to_string(&args[1]);
                let start = if args.len() >= 3 {
                    Some(ValueConversion::to_number(&args[2]) as usize)
                } else {
                    None
                };
                (s1, s2, start)
            } else {
                return None;
            };

            // 从末尾查找子串
            let search_str = if let Some(pos) = start {
                if pos > string1.len() {
                    return Some(Ok(Value::Number(0.0)));
                } else if pos > 0 {
                    &string1[..pos.min(string1.len())]
                } else {
                    &string1
                }
            } else {
                &string1
            };

            // 使用 rfind 从右查找
            let pos = search_str.rfind(&string2);
            let result = if let Some(p) = pos {
                (p + 1) as f64  // VBScript 位置从 1 开始
            } else {
                0.0
            };
            Some(Ok(Value::Number(result)))
        }
        "split" => {
            // Split(string[, delimiter[, count[, compare]]])
            // 将字符串分割为数组
            if args.len() < 1 {
                return None;
            }
            let string = ValueConversion::to_string(&args[0]);
            let delimiter = if args.len() >= 2 {
                ValueConversion::to_string(&args[1])
            } else {
                " ".to_string()
            };

            let parts: Vec<Value> = if delimiter.is_empty() {
                // 空分隔符，返回单个字符数组
                string.chars().map(|c| Value::String(c.to_string())).collect()
            } else {
                string.split(&delimiter).map(|s| Value::String(s.to_string())).collect()
            };

            Some(Ok(Value::Array(parts)))
        }
        "join" => {
            // Join(array[, delimiter])
            // 将数组合并为字符串
            if args.len() < 1 {
                return None;
            }
            let delimiter = if args.len() >= 2 {
                ValueConversion::to_string(&args[1])
            } else {
                " ".to_string()
            };

            let result = match &args[0] {
                Value::Array(arr) => {
                    let parts: Vec<String> = arr.iter()
                        .map(|v| ValueConversion::to_string(v))
                        .collect();
                    Value::String(parts.join(&delimiter))
                }
                _ => Value::String(String::new())
            };
            Some(Ok(result))
        }
        "ubound" => {
            if args.len() < 1 {
                return None;
            }
            match &args[0] {
                Value::Array(arr) => {
                    if arr.is_empty() {
                        Some(Ok(Value::Number(-1.0)))
                    } else {
                        Some(Ok(Value::Number((arr.len() - 1) as f64)))
                    }
                }
                _ => Some(Ok(Value::Number(-1.0)))
            }
        }
        "lbound" => {
            if args.len() < 1 {
                return None;
            }
            match &args[0] {
                Value::Array(_) => Some(Ok(Value::Number(0.0))),
                _ => Some(Ok(Value::Number(0.0)))
            }
        }
        "filter" => {
            if args.len() < 2 {
                return None;
            }
            let filter_value = ValueConversion::to_string(&args[1]);
            let include = if args.len() >= 3 {
                ValueConversion::to_bool(&args[2])
            } else {
                true
            };

            match &args[0] {
                Value::Array(arr) => {
                    let filtered: Vec<Value> = arr.iter()
                        .filter(|v| {
                            let s = ValueConversion::to_string(&**v);
                            let contains = s.contains(&filter_value);
                            if include { contains } else { !contains }
                        })
                        .cloned()
                        .collect();
                    Some(Ok(Value::Array(filtered)))
                }
                _ => Some(Ok(Value::Array(vec![])))
            }
        }
        "array" => {
            Some(Ok(Value::Array(args.to_vec())))
        }
        "replace" => {
            // Replace(string, find, replacewith[, start[, count[, compare]]])
            if args.len() < 3 {
                return None;
            }
            let string = ValueConversion::to_string(&args[0]);
            let find = ValueConversion::to_string(&args[1]);
            let replace_with = ValueConversion::to_string(&args[2]);
            let start = if args.len() >= 4 {
                ValueConversion::to_number(&args[3]) as usize
            } else {
                1
            };
            let count = if args.len() >= 5 {
                ValueConversion::to_number(&args[4]) as usize
            } else {
                usize::MAX
            };

            // 获取要搜索的子串
            let search_str = if start > 1 && start <= string.len() {
                &string[start - 1..]
            } else if start > string.len() {
                return Some(Ok(Value::String(string)));
            } else {
                &string
            };

            // 执行替换
            let result = if count == usize::MAX {
                search_str.replacen(&find, &replace_with, usize::MAX)
            } else {
                search_str.replacen(&find, &replace_with, count)
            };

            // 组合结果
            let final_result = if start > 1 && start <= string.len() {
                format!("{}{}", &string[..start - 1], result)
            } else {
                result
            };

            Some(Ok(Value::String(final_result)))
        }
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
        "clng" | "csng" | "ccur" => Some(Ok(Value::Number(ValueConversion::to_number(arg)))),
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
        "sin" => {
            let n = ValueConversion::to_number(arg);
            Some(Ok(Value::Number(n.sin())))
        }
        "cos" => {
            let n = ValueConversion::to_number(arg);
            Some(Ok(Value::Number(n.cos())))
        }
        "tan" => {
            let n = ValueConversion::to_number(arg);
            Some(Ok(Value::Number(n.tan())))
        }
        "atn" => {
            let n = ValueConversion::to_number(arg);
            Some(Ok(Value::Number(n.atan())))
        }
        "exp" => {
            let n = ValueConversion::to_number(arg);
            Some(Ok(Value::Number(n.exp())))
        }
        "log" => {
            let n = ValueConversion::to_number(arg);
            Some(Ok(Value::Number(n.ln())))
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
        "strreverse" => {
            let s = ValueConversion::to_string(arg);
            let reversed: String = s.chars().rev().collect();
            Some(Ok(Value::String(reversed)))
        }
        "space" => {
            let n = ValueConversion::to_number(arg) as usize;
            let spaces = " ".repeat(n.min(1000000)); // 限制最大长度防止内存溢出
            Some(Ok(Value::String(spaces)))
        }
        "string" => {
            // String(number, character) - 返回重复的字符
            let n = ValueConversion::to_number(arg) as usize;
            // 对于 String 函数，第一个参数是数字，但我们这里只有单参数
            // 简化实现：返回 n 个空格
            let spaces = " ".repeat(n.min(1000000));
            Some(Ok(Value::String(spaces)))
        }
        "chr" => {
            let n = ValueConversion::to_number(arg) as u32;
            Some(Ok(Value::String(
                char::from_u32(n).unwrap_or('\0').to_string(),
            )))
        }
        "chrw" => {
            // ChrW - 返回 Unicode 字符（与 Chr 相同，因为 Rust 使用 UTF-8）
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
        "ascw" => {
            // AscW - 返回 Unicode 码点（16 位）
            let s = ValueConversion::to_string(arg);
            let code = s.chars().next().map(|c| c as u32 as f64).unwrap_or(0.0);
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
        "hex" => {
            let n = ValueConversion::to_number(arg) as i64;
            Some(Ok(Value::String(format!("{:X}", n))))
        }
        "oct" => {
            let n = ValueConversion::to_number(arg) as i64;
            Some(Ok(Value::String(format!("{:o}", n))))
        }
        "vartype" => {
            // VarType 返回值类型的数值代码
            let type_code = match arg {
                Value::Empty => 0,
                Value::Null => 1,
                Value::Boolean(_) => 11,
                Value::Number(_) => 5,  // Double
                Value::String(_) => 8,
                Value::Array(_) => 8192,
                Value::Object(_) => 9,
                Value::Nothing => 0,
            };
            Some(Ok(Value::Number(type_code as f64)))
        }
        "typename" => {
            // TypeName 返回值类型名称
            let type_name = match arg {
                Value::Empty => "Empty",
                Value::Null => "Null",
                Value::Boolean(_) => "Boolean",
                Value::Number(_) => "Double",
                Value::String(_) => "String",
                Value::Array(_) => "Array",
                Value::Object(_) => "Object",
                Value::Nothing => "Nothing",
            };
            Some(Ok(Value::String(type_name.to_string())))
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
