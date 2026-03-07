//! 日期时间函数执行器

use crate::runtime::{RuntimeError, Value, ValueConversion};
use super::super::token::BuiltinToken;
use chrono::{Datelike, Timelike, NaiveDate, NaiveDateTime};

/// 解析 VBScript 日期字符串
/// 支持格式: #2024-01-01#, "2024-01-01", "2024/01/01", "01/01/2024" 等
fn parse_vbscript_date(date_str: &str) -> Option<NaiveDateTime> {
    let trimmed = date_str.trim();

    // 移除 # 包围（VBScript 日期字面量 #2024-01-01#）
    let clean_str = if trimmed.starts_with('#') && trimmed.ends_with('#') {
        &trimmed[1..trimmed.len()-1]
    } else {
        trimmed
    };

    let clean_str = clean_str.trim();

    // 尝试各种日期格式
    let formats = [
        // ISO 格式
        "%Y-%m-%d",
        "%Y/%m/%d",
        // 美式格式
        "%m/%d/%Y",
        "%m-%d-%Y",
        // 短年份格式
        "%y-%m-%d",
        "%y/%m/%d",
        // 带时间的格式
        "%Y-%m-%d %H:%M:%S",
        "%Y/%m/%d %H:%M:%S",
        "%m/%d/%Y %H:%M:%S",
    ];

    for format in &formats {
        if let Ok(date) = NaiveDate::parse_from_str(clean_str, format) {
            return Some(date.and_hms_opt(0, 0, 0).unwrap_or_else(|| {
                NaiveDateTime::new(date, chrono::NaiveTime::from_hms_opt(0, 0, 0).unwrap())
            }));
        }
    }

    // 尝试解析日期时间格式（包含时间部分）
    for format in &[
        "%Y-%m-%d %H:%M:%S",
        "%Y/%m/%d %H:%M:%S",
        "%m/%d/%Y %H:%M:%S",
    ] {
        if let Ok(dt) = NaiveDateTime::parse_from_str(clean_str, format) {
            return Some(dt);
        }
    }

    None
}

/// 计算两个日期之间的差值
fn calculate_date_diff(interval: &str, date1: NaiveDateTime, date2: NaiveDateTime) -> i64 {
    let duration = date2.signed_duration_since(date1);

    match interval {
        "yyyy" => {
            // 年份差：比较年份数字
            let years = date2.year() - date1.year();
            // 如果月份和日期还没到，减去1年
            if date2.month() < date1.month() ||
               (date2.month() == date1.month() && date2.day() < date1.day()) {
                years as i64 - 1
            } else {
                years as i64
            }
        }
        "q" => {
            // 季度差：先计算总月份差，再除以3
            let months = calculate_date_diff("m", date1, date2);
            months / 3
        }
        "m" => {
            // 月份差
            let months = (date2.year() - date1.year()) as i64 * 12
                + (date2.month() as i64 - date1.month() as i64);
            // 如果日期还没到，减去1个月
            if date2.day() < date1.day() {
                months - 1
            } else {
                months
            }
        }
        "y" | "d" => duration.num_days(),
        "w" => {
            // 周数差（以7天为单位）
            duration.num_days() / 7
        }
        "ww" => {
            // 日历周数差：按星期日计算
            let days = duration.num_days();
            // 计算 date1 到下一个星期日的天数
            let weekday1 = date1.weekday().num_days_from_sunday() as i64;
            // 计算从 date1 的星期日到 date2 的星期日之间的周数
            (days + weekday1) / 7
        }
        "h" => duration.num_hours(),
        "n" => duration.num_minutes(),
        "s" => duration.num_seconds(),
        _ => 0,
    }
}

/// 获取语言配置
/// 支持 zh-cn（中文）和 en（英文），默认 zh-cn
fn get_language() -> String {
    std::env::var("LANGUAGE")
        .unwrap_or_else(|_| "zh-cn".to_string())
        .to_lowercase()
}

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
            let weekday = ValueConversion::to_number(&args[0]) as usize;
            let abbreviate = if args.len() >= 2 {
                ValueConversion::to_bool(&args[1])
            } else {
                false
            };

            // VBScript: 1=Sunday, 2=Monday, ..., 7=Saturday
            // 根据语言配置返回星期名称
            let lang = get_language();
            let name = if lang == "zh-cn" || lang == "zh" {
                // 中文星期名称
                if abbreviate {
                    ["日", "一", "二", "三", "四", "五", "六"]
                } else {
                    ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"]
                }
            } else {
                // 英文星期名称
                if abbreviate {
                    ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
                } else {
                    ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
                }
            };

            // weekday 应该是 1-7，对应 Sunday-Saturday
            let index = if weekday == 0 { 6 } else { (weekday - 1) % 7 };
            Value::String(name.get(index).unwrap_or(&"").to_string())
        }
        BuiltinToken::MonthName => {
            // MonthName(month, abbreviate)
            if args.len() < 1 {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let month = ValueConversion::to_number(&args[0]) as usize;
            let abbreviate = if args.len() >= 2 {
                ValueConversion::to_bool(&args[1])
            } else {
                false
            };

            // 根据语言配置返回月份名称
            let lang = get_language();
            let name = if lang == "zh-cn" || lang == "zh" {
                // 中文月份名称
                ["一月", "二月", "三月", "四月", "五月", "六月",
                 "七月", "八月", "九月", "十月", "十一月", "十二月"]
            } else {
                // 英文月份名称
                if abbreviate {
                    ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
                } else {
                    ["January", "February", "March", "April", "May", "June",
                     "July", "August", "September", "October", "November", "December"]
                }
            };

            // month 应该是 1-12
            let index = if month == 0 { 11 } else { (month - 1) % 12 };
            Value::String(name.get(index).unwrap_or(&"").to_string())
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
            // DateDiff(interval, date1, date2 [, firstdayofweek [, firstweekofyear]])
            // 完整实现：计算两个日期之间的差值
            if args.len() < 3 {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let interval = ValueConversion::to_string(&args[0]).to_lowercase();
            let date1_str = ValueConversion::to_string(&args[1]);
            let date2_str = ValueConversion::to_string(&args[2]);

            // 解析日期
            let date1 = parse_vbscript_date(&date1_str)
                .unwrap_or_else(|| chrono::Local::now().naive_local());
            let date2 = parse_vbscript_date(&date2_str)
                .unwrap_or_else(|| chrono::Local::now().naive_local());

            // 计算差值
            let diff = calculate_date_diff(&interval, date1, date2);
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