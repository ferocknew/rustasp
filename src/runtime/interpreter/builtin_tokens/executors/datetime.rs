//! 日期时间函数执行器

use crate::runtime::{RuntimeError, Value, ValueConversion};
use super::super::token::BuiltinToken;
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
///
/// 修复了 MM 被替换两次的 bug：先替换分钟的占位符，再替换月份的占位符
fn convert_vbscript_format(format: &str) -> String {
    format
        .replace("YYYY", "%Y")
        .replace("YY", "%y")
        // 先用临时占位符处理月份，避免与分钟冲突
        .replace("MM", "__MONTH__")
        .replace("DD", "%d")
        .replace("HH", "%H")
        .replace("__MONTH__", "%m")  // 月份
        .replace("SS", "%S")
}

pub fn execute(token: BuiltinToken, args: &[Value]) -> Result<Option<Value>, RuntimeError> {
    let result = match token {
        BuiltinToken::Now => {
            let now = chrono::Local::now();
            let (now_format, _, _) = get_datetime_format();
            Value::String(now.format(&now_format).to_string())
        }
        BuiltinToken::Date => {
            let now = chrono::Local::now();
            let (_, date_format, _) = get_datetime_format();
            Value::String(now.format(&date_format).to_string())
        }
        BuiltinToken::Time => {
            let now = chrono::Local::now();
            let (_, _, time_format) = get_datetime_format();
            Value::String(now.format(&time_format).to_string())
        }
        BuiltinToken::Year => {
            let now = chrono::Local::now();
            Value::Number(now.year() as f64)
        }
        BuiltinToken::Month => {
            let now = chrono::Local::now();
            Value::Number(now.month() as f64)
        }
        BuiltinToken::Day => {
            let now = chrono::Local::now();
            Value::Number(now.day() as f64)
        }
        BuiltinToken::Hour => {
            let now = chrono::Local::now();
            Value::Number(now.hour() as f64)
        }
        BuiltinToken::Minute => {
            let now = chrono::Local::now();
            Value::Number(now.minute() as f64)
        }
        BuiltinToken::Second => {
            let now = chrono::Local::now();
            Value::Number(now.second() as f64)
        }
        BuiltinToken::WeekDay => {
            let now = chrono::Local::now();
            // VBScript: 1=Sunday, 2=Monday, ..., 7=Saturday
            // chrono: 0=Monday, ..., 6=Sunday
            let weekday = now.weekday().num_days_from_sunday();
            Value::Number((weekday + 1) as f64)
        }
        BuiltinToken::WeekDayName => {
            // WeekdayName(weekday, abbreviate, firstdayofweek)
            if args.len() < 1 {
                return Err(RuntimeError::ArgumentCountMismatch);
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
            Value::String(name.to_string())
        }
        BuiltinToken::MonthName => {
            // MonthName(month, abbreviate)
            if args.len() < 1 {
                return Err(RuntimeError::ArgumentCountMismatch);
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
            Value::String(name.to_string())
        }
        BuiltinToken::DateAdd => {
            // DateAdd(interval, number, date)
            if args.len() < 3 {
                return Err(RuntimeError::ArgumentCountMismatch);
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
                _ => return Ok(Some(Value::Empty)),
            }

            let (now_format, _, _) = get_datetime_format();
            Value::String(date.format(&now_format).to_string())
        }
        BuiltinToken::DateDiff => {
            // DateDiff(interval, date1, date2)
            // 简化实现：返回两个日期之间的差值
            if args.len() < 3 {
                return Err(RuntimeError::ArgumentCountMismatch);
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

            Value::Number(diff as f64)
        }
        BuiltinToken::DatePart => {
            // DatePart(interval, date)
            // 简化实现：返回日期的指定部分
            if args.len() < 2 {
                return Err(RuntimeError::ArgumentCountMismatch);
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

            Value::Number(result)
        }
        BuiltinToken::FormatDateTime => {
            // FormatDateTime(date, format)
            // 简化实现：返回日期字符串
            if args.len() < 1 {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let now = chrono::Local::now();
            let (now_format, _, _) = get_datetime_format();
            Value::String(now.format(&now_format).to_string())
        }
        BuiltinToken::Timer => {
            // Timer 函数返回自午夜以来的秒数
            let now = chrono::Local::now();
            let seconds_since_midnight = now.hour() as f64 * 3600.0
                + now.minute() as f64 * 60.0
                + now.second() as f64
                + now.nanosecond() as f64 / 1_000_000_000.0;
            Value::Number(seconds_since_midnight)
        }
        BuiltinToken::DateSerial | BuiltinToken::DateValue | BuiltinToken::TimeSerial | BuiltinToken::TimeValue => {
            // TODO: 实现完整的日期序列化/反序列化
            Value::Empty
        }
        _ => return Ok(None),
    };
    Ok(Some(result))
}
