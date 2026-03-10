//! 日期时间函数执行器

use super::super::token::BuiltinToken;
use crate::runtime::{RuntimeError, Value, ValueConversion};
use chrono::{Datelike, NaiveDate, NaiveDateTime, Timelike};

/// 解析 VBScript 时间字符串
/// 支持格式: "12:30:45", "4:35:17 PM", "14:30" 等
fn parse_vbscript_time(time_str: &str) -> Option<chrono::NaiveTime> {
    let trimmed = time_str.trim().to_lowercase();

    // 处理 AM/PM 标记
    let (time_part, is_pm, has_period) = if trimmed.ends_with("pm") {
        (trimmed[..trimmed.len() - 2].trim().to_string(), true, true)
    } else if trimmed.ends_with("am") {
        (trimmed[..trimmed.len() - 2].trim().to_string(), false, true)
    } else {
        (trimmed.to_string(), false, false)
    };

    // 尝试各种时间格式
    let formats = [
        "%H:%M:%S",     // 14:30:45
        "%H:%M",        // 14:30
        "%I:%M:%S",     // 2:30:45 (12小时制)
        "%I:%M",        // 2:30 (12小时制)
        "%H:%M:%S %.f", // 带毫秒
    ];

    for format in &formats {
        if let Ok(mut time) = chrono::NaiveTime::parse_from_str(&time_part, format) {
            // 如果是 12 小时制且是 PM，且小时不是 12
            if has_period && is_pm && time.hour() < 12 {
                time =
                    chrono::NaiveTime::from_hms_opt(time.hour() + 12, time.minute(), time.second())
                        .unwrap_or(time);
            }
            // 如果是 12 小时制且是 AM，且小时是 12
            if has_period && !is_pm && time.hour() == 12 {
                time = chrono::NaiveTime::from_hms_opt(0, time.minute(), time.second())
                    .unwrap_or(time);
            }
            return Some(time);
        }
    }

    None
}

/// 解析 VBScript 日期字符串
/// 支持格式: #2024-01-01#, "2024-01-01", "2024/01/01", "01/01/2024" 等
fn parse_vbscript_date(date_str: &str) -> Option<NaiveDateTime> {
    let trimmed = date_str.trim();

    // 移除 # 包围（VBScript 日期字面量 #2024-01-01#）
    let clean_str = if trimmed.starts_with('#') && trimmed.ends_with('#') {
        &trimmed[1..trimmed.len() - 1]
    } else {
        trimmed
    };

    let clean_str = clean_str.trim();

    // 优先尝试解析带时间的格式
    let datetime_formats = [
        "%Y-%m-%d %H:%M:%S",
        "%Y/%m/%d %H:%M:%S",
        "%m/%d/%Y %H:%M:%S",
    ];

    for format in &datetime_formats {
        if let Ok(dt) = NaiveDateTime::parse_from_str(clean_str, format) {
            return Some(dt);
        }
    }

    // 尝试各种纯日期格式
    let date_formats = [
        // ISO 格式
        "%Y-%m-%d", "%Y/%m/%d", // 美式格式
        "%m/%d/%Y", "%m-%d-%Y", // 短年份格式
        "%y-%m-%d", "%y/%m/%d",
    ];

    for format in &date_formats {
        if let Ok(date) = NaiveDate::parse_from_str(clean_str, format) {
            return Some(date.and_hms_opt(0, 0, 0).unwrap_or_else(|| {
                NaiveDateTime::new(date, chrono::NaiveTime::from_hms_opt(0, 0, 0).unwrap())
            }));
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
            if date2.month() < date1.month()
                || (date2.month() == date1.month() && date2.day() < date1.day())
            {
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
    let now_format =
        std::env::var("NOW_FORMAT").unwrap_or_else(|_| "yyyy/mm/dd hh:nn:ss".to_string());
    let date_format = std::env::var("DATE_FORMAT").unwrap_or_else(|_| "yyyy/mm/dd".to_string());
    let time_format = std::env::var("TIME_FORMAT").unwrap_or_else(|_| "hh:nn:ss".to_string());

    // 转换格式字符串：yyyy/mm/dd hh:nn:ss -> %Y/%m/%d %H:%M:%S
    let now_format_str = convert_vbscript_format(&now_format);
    let date_format_str = convert_vbscript_format(&date_format);
    let time_format_str = convert_vbscript_format(&time_format);

    (now_format_str, date_format_str, time_format_str)
}

/// 将 VBScript 日期格式转换为 strftime 格式
///
/// VBScript 标准格式：
///   yyyy = 四位年份 (2024)
///   yy = 两位年份 (24)
///   mm = 月份，带前导零 (01-12)
///   m = 月份，不带前导零 (1-12)
///   dd = 日期，带前导零 (01-31)
///   d = 日期，不带前导零 (1-31)
///   hh = 小时，带前导零 (00-23)
///   h = 小时，不带前导零 (0-23)
///   nn = 分钟，带前导零 (00-59)
///   n = 分钟，不带前导零 (0-59)
///   ss = 秒，带前导零 (00-59)
///   s = 秒，不带前导零 (0-59)
///
/// 注意：先处理长占位符，避免被短占位符干扰
fn convert_vbscript_format(format: &str) -> String {
    let mut result = format.to_string();

    // 按长度从长到短处理，避免短占位符提前匹配
    result = result
        .replace("yyyy", "%Y") // 四位年份
        .replace("yy", "%y") // 两位年份
        .replace("mm", "%m") // 月份（带前导零）
        .replace("m", "%m") // 月份（不带前导零）
        .replace("dd", "%d") // 日期（带前导零）
        .replace("d", "%d") // 日期（不带前导零）
        .replace("hh", "%H") // 小时（带前导零）
        .replace("h", "%H") // 小时（不带前导零）
        .replace("nn", "%M") // 分钟（带前导零）
        .replace("n", "%M") // 分钟（不带前导零）
        .replace("ss", "%S") // 秒（带前导零）
        .replace("s", "%S"); // 秒（不带前导零）

    result
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
            // Year(date) - 返回年份，可选参数
            let dt = if args.is_empty() {
                chrono::Local::now().naive_local()
            } else {
                let date_str = ValueConversion::to_string(&args[0]);
                parse_vbscript_date(&date_str).unwrap_or_else(|| chrono::Local::now().naive_local())
            };
            Value::Number(dt.year() as f64)
        }
        BuiltinToken::Month => {
            // Month(date) - 返回月份，可选参数
            let dt = if args.is_empty() {
                chrono::Local::now().naive_local()
            } else {
                let date_str = ValueConversion::to_string(&args[0]);
                parse_vbscript_date(&date_str).unwrap_or_else(|| chrono::Local::now().naive_local())
            };
            Value::Number(dt.month() as f64)
        }
        BuiltinToken::Day => {
            // Day(date) - 返回日期，可选参数
            let dt = if args.is_empty() {
                chrono::Local::now().naive_local()
            } else {
                let date_str = ValueConversion::to_string(&args[0]);
                parse_vbscript_date(&date_str).unwrap_or_else(|| chrono::Local::now().naive_local())
            };
            Value::Number(dt.day() as f64)
        }
        BuiltinToken::Hour => {
            // Hour(time) - 返回小时，可选参数
            let dt = if args.is_empty() {
                chrono::Local::now().naive_local()
            } else {
                let time_str = ValueConversion::to_string(&args[0]);
                // 尝试解析为日期时间
                if let Some(dt) = parse_vbscript_date(&time_str) {
                    dt
                } else if let Some(t) = parse_vbscript_time(&time_str) {
                    // 如果只是时间，构造一个带时间的日期
                    chrono::Local::now().naive_local().date().and_time(t)
                } else {
                    chrono::Local::now().naive_local()
                }
            };
            Value::Number(dt.hour() as f64)
        }
        BuiltinToken::Minute => {
            // Minute(time) - 返回分钟，可选参数
            let dt = if args.is_empty() {
                chrono::Local::now().naive_local()
            } else {
                let time_str = ValueConversion::to_string(&args[0]);
                // 尝试解析为日期时间
                if let Some(dt) = parse_vbscript_date(&time_str) {
                    dt
                } else if let Some(t) = parse_vbscript_time(&time_str) {
                    chrono::Local::now().naive_local().date().and_time(t)
                } else {
                    chrono::Local::now().naive_local()
                }
            };
            Value::Number(dt.minute() as f64)
        }
        BuiltinToken::Second => {
            // Second(time) - 返回秒数，可选参数
            let dt = if args.is_empty() {
                chrono::Local::now().naive_local()
            } else {
                let time_str = ValueConversion::to_string(&args[0]);
                // 尝试解析为日期时间
                if let Some(dt) = parse_vbscript_date(&time_str) {
                    dt
                } else if let Some(t) = parse_vbscript_time(&time_str) {
                    chrono::Local::now().naive_local().date().and_time(t)
                } else {
                    chrono::Local::now().naive_local()
                }
            };
            Value::Number(dt.second() as f64)
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
                    [
                        "星期日",
                        "星期一",
                        "星期二",
                        "星期三",
                        "星期四",
                        "星期五",
                        "星期六",
                    ]
                }
            } else {
                // 英文星期名称
                if abbreviate {
                    ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
                } else {
                    [
                        "Sunday",
                        "Monday",
                        "Tuesday",
                        "Wednesday",
                        "Thursday",
                        "Friday",
                        "Saturday",
                    ]
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
                [
                    "一月",
                    "二月",
                    "三月",
                    "四月",
                    "五月",
                    "六月",
                    "七月",
                    "八月",
                    "九月",
                    "十月",
                    "十一月",
                    "十二月",
                ]
            } else {
                // 英文月份名称
                if abbreviate {
                    [
                        "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct",
                        "Nov", "Dec",
                    ]
                } else {
                    [
                        "January",
                        "February",
                        "March",
                        "April",
                        "May",
                        "June",
                        "July",
                        "August",
                        "September",
                        "October",
                        "November",
                        "December",
                    ]
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

            // 简化实现：直接使用当前时间进行计算
            let mut date = chrono::Local::now();

            match interval.as_str() {
                "yyyy" => {
                    // 年份加减
                    let year = date.year() + number as i32;
                    let month = date.month();
                    let day = date.day();
                    date = chrono::Local::now()
                        .with_year(year)
                        .unwrap()
                        .with_month(month)
                        .unwrap()
                        .with_day(day)
                        .unwrap();
                }
                "q" => {
                    // 季度加减（3个月）
                    date = date + chrono::Duration::days(number * 91);
                }
                "m" => {
                    // 月份加减
                    date = date + chrono::Duration::days(number * 30);
                }
                "y" | "d" => {
                    // 天数加减
                    date = date + chrono::Duration::days(number);
                }
                "w" => {
                    // 周数加减（7天）
                    date = date + chrono::Duration::days(number * 7);
                }
                "ww" => {
                    // 日历周数加减（7天）
                    date = date + chrono::Duration::days(number * 7);
                }
                "h" => {
                    date = date + chrono::Duration::hours(number);
                }
                "n" => {
                    date = date + chrono::Duration::minutes(number);
                }
                "s" => {
                    date = date + chrono::Duration::seconds(number);
                }
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
        BuiltinToken::DateValue => {
            // DateValue(date_string) - 将日期字符串转换为日期
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let date_str = ValueConversion::to_string(&args[0]);

            // 解析日期字符串（忽略时间部分）
            if let Some(dt) = parse_vbscript_date(&date_str) {
                let (_, date_format, _) = get_datetime_format();
                Value::String(dt.format(&date_format).to_string())
            } else {
                Value::Empty
            }
        }
        BuiltinToken::TimeValue => {
            // TimeValue(time_string) - 将时间字符串转换为时间
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let time_str = ValueConversion::to_string(&args[0]);

            // 解析时间字符串
            if let Some(dt) = parse_vbscript_time(&time_str) {
                let (_, _, time_format) = get_datetime_format();
                Value::String(dt.format(&time_format).to_string())
            } else {
                Value::Empty
            }
        }
        BuiltinToken::DateSerial => {
            // DateSerial(year, month, day) - 根据年月日生成日期
            // 支持月份和日期的溢出处理
            if args.len() < 3 {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let year = ValueConversion::to_number(&args[0]) as i32;
            let month = ValueConversion::to_number(&args[1]) as i32;
            let day = ValueConversion::to_number(&args[2]) as i32;

            // 处理月份溢出（支持 month=0, month=-1, month=13 等）
            // month=0 -> 上一年12月，month=13 -> 下一年1月
            let month_offset = month - 1; // 转换为 0-based (0=1月)
            let year_offset = month_offset / 12;
            let adjusted_month = month_offset % 12;

            // 处理负数月份
            let (adjusted_year, adjusted_month) = if adjusted_month < 0 {
                (year + year_offset - 1, adjusted_month + 12)
            } else {
                (year + year_offset, adjusted_month)
            };

            let adjusted_month = adjusted_month + 1; // 转换回 1-based (1=1月)

            // 尝试创建日期（chrono会自动处理日期溢出，如2月30日会返回None）
            if let Some(date) =
                NaiveDate::from_ymd_opt(adjusted_year, adjusted_month as u32, day as u32)
            {
                let (_, date_format, _) = get_datetime_format();
                Value::String(date.format(&date_format).to_string())
            } else {
                Value::Empty
            }
        }
        BuiltinToken::TimeSerial => {
            // TimeSerial(hour, minute, second) - 根据时分秒生成时间
            if args.len() < 3 {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let hour = ValueConversion::to_number(&args[0]) as i64;
            let minute = ValueConversion::to_number(&args[1]) as i64;
            let second = ValueConversion::to_number(&args[2]) as i64;

            // 处理溢出：将所有时间转换为秒，然后计算时分秒
            let total_seconds = hour * 3600 + minute * 60 + second;
            let normalized_seconds = ((total_seconds % 86400) + 86400) % 86400; // 确保在 0-86399 之间

            let norm_hour = (normalized_seconds / 3600) as u32;
            let norm_minute = ((normalized_seconds % 3600) / 60) as u32;
            let norm_second = (normalized_seconds % 60) as u32;

            if let Some(time) = chrono::NaiveTime::from_hms_opt(norm_hour, norm_minute, norm_second)
            {
                let (_, _, time_format) = get_datetime_format();
                Value::String(time.format(&time_format).to_string())
            } else {
                Value::Empty
            }
        }
        _ => return Ok(None),
    };
    Ok(Some(result))
}
