//! 内置函数 Token 识别模块
//!
//! 使用 Token Key 而非字符串匹配，提高内置函数调用性能
//! 将函数名映射为整数 ID，通过 match/switch 快速分发

use crate::runtime::{RuntimeError, Value, ValueConversion};
use rand::Rng;
use std::collections::HashMap;

/// 内置函数 Token ID
/// 每个内置函数对应一个唯一的整数 ID，用于快速匹配
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u16)]
pub enum BuiltinToken {
    // ========== 数学函数 (1-50) ==========
    Abs = 1,
    Sqr = 2,
    Sin = 3,
    Cos = 4,
    Tan = 5,
    Atn = 6,
    Log = 7,
    Exp = 8,
    Int = 9,
    Fix = 10,
    Round = 11,
    Rnd = 12,
    Sgn = 13,

    // ========== 类型转换函数 (51-100) ==========
    CStr = 51,
    CInt = 52,
    CLng = 53,
    CSng = 54,
    CDbl = 55,
    CBool = 56,
    CDate = 57,
    CByte = 58,

    // ========== 字符串函数 (101-150) ==========
    Len = 101,
    Trim = 102,
    LTrim = 103,
    RTrim = 104,
    Left = 105,
    Right = 106,
    Mid = 107,
    UCase = 108,
    LCase = 109,
    InStr = 110,
    InStrRev = 111,
    StrComp = 112,
    Replace = 113,
    Split = 114,
    Join = 115,
    StrReverse = 116,
    Space = 117,
    String_ = 118,  // String 是 Rust 关键字，加下划线
    Asc = 119,
    Chr = 120,
    AscW = 121,
    ChrW = 122,

    // ========== 日期时间函数 (151-200) ==========
    Now = 151,
    Date = 152,
    Time = 153,
    Year = 154,
    Month = 155,
    Day = 156,
    Hour = 157,
    Minute = 158,
    Second = 159,
    WeekDay = 160,
    DateAdd = 161,
    DateDiff = 162,
    DatePart = 163,
    DateSerial = 164,
    DateValue = 165,
    TimeSerial = 166,
    TimeValue = 167,
    FormatDateTime = 168,
    MonthName = 169,
    WeekDayName = 170,

    // ========== 数组函数 (201-220) ==========
    Array = 201,
    UBound = 202,
    LBound = 203,
    Filter = 204,
    IsArray = 205,

    // ========== 检验函数 (221-250) ==========
    IsNumeric = 221,
    IsDate = 222,
    IsEmpty = 223,
    IsNull = 224,
    IsObject = 225,
    IsNothing = 226,
    TypeName = 227,
    VarType = 228,

    // ========== 交互函数 (251-270) ==========
    MsgBox = 251,
    InputBox = 252,

    // ========== 其他函数 (271-300) ==========
    CreateObject = 271,
    GetObject = 272,
    Eval = 273,
    Execute = 274,
    RGB = 275,

    // 未知函数
    Unknown = 0,
}

/// Token 注册表
/// 维护函数名到 Token ID 的映射
pub struct TokenRegistry {
    /// 函数名到 Token 的映射（小写存储，不区分大小写）
    map: HashMap<String, BuiltinToken>,
}

impl TokenRegistry {
    /// 创建新的 Token 注册表并初始化所有内置函数
    pub fn new() -> Self {
        let mut registry = Self {
            map: HashMap::new(),
        };
        registry.init_all_tokens();
        registry
    }

    /// 初始化所有内置函数 Token
    fn init_all_tokens(&mut self) {
        // 数学函数
        self.register("abs", BuiltinToken::Abs);
        self.register("sqr", BuiltinToken::Sqr);
        self.register("sin", BuiltinToken::Sin);
        self.register("cos", BuiltinToken::Cos);
        self.register("tan", BuiltinToken::Tan);
        self.register("atn", BuiltinToken::Atn);
        self.register("log", BuiltinToken::Log);
        self.register("exp", BuiltinToken::Exp);
        self.register("int", BuiltinToken::Int);
        self.register("fix", BuiltinToken::Fix);
        self.register("round", BuiltinToken::Round);
        self.register("rnd", BuiltinToken::Rnd);
        self.register("sgn", BuiltinToken::Sgn);

        // 类型转换函数
        self.register("cstr", BuiltinToken::CStr);
        self.register("cint", BuiltinToken::CInt);
        self.register("clng", BuiltinToken::CLng);
        self.register("csng", BuiltinToken::CSng);
        self.register("cdbl", BuiltinToken::CDbl);
        self.register("cbool", BuiltinToken::CBool);
        self.register("cdate", BuiltinToken::CDate);
        self.register("cbyte", BuiltinToken::CByte);

        // 字符串函数
        self.register("len", BuiltinToken::Len);
        self.register("trim", BuiltinToken::Trim);
        self.register("ltrim", BuiltinToken::LTrim);
        self.register("rtrim", BuiltinToken::RTrim);
        self.register("left", BuiltinToken::Left);
        self.register("right", BuiltinToken::Right);
        self.register("mid", BuiltinToken::Mid);
        self.register("ucase", BuiltinToken::UCase);
        self.register("lcase", BuiltinToken::LCase);
        self.register("instr", BuiltinToken::InStr);
        self.register("instrrev", BuiltinToken::InStrRev);
        self.register("strcomp", BuiltinToken::StrComp);
        self.register("replace", BuiltinToken::Replace);
        self.register("split", BuiltinToken::Split);
        self.register("join", BuiltinToken::Join);
        self.register("strreverse", BuiltinToken::StrReverse);
        self.register("space", BuiltinToken::Space);
        self.register("string", BuiltinToken::String_);
        self.register("asc", BuiltinToken::Asc);
        self.register("chr", BuiltinToken::Chr);
        self.register("ascw", BuiltinToken::AscW);
        self.register("chrw", BuiltinToken::ChrW);

        // 日期时间函数
        self.register("now", BuiltinToken::Now);
        self.register("date", BuiltinToken::Date);
        self.register("time", BuiltinToken::Time);
        self.register("year", BuiltinToken::Year);
        self.register("month", BuiltinToken::Month);
        self.register("day", BuiltinToken::Day);
        self.register("hour", BuiltinToken::Hour);
        self.register("minute", BuiltinToken::Minute);
        self.register("second", BuiltinToken::Second);
        self.register("weekday", BuiltinToken::WeekDay);
        self.register("dateadd", BuiltinToken::DateAdd);
        self.register("datediff", BuiltinToken::DateDiff);
        self.register("datepart", BuiltinToken::DatePart);
        self.register("dateserial", BuiltinToken::DateSerial);
        self.register("datevalue", BuiltinToken::DateValue);
        self.register("timeserial", BuiltinToken::TimeSerial);
        self.register("timevalue", BuiltinToken::TimeValue);
        self.register("formatdatetime", BuiltinToken::FormatDateTime);
        self.register("monthname", BuiltinToken::MonthName);
        self.register("weekdayname", BuiltinToken::WeekDayName);

        // 数组函数
        self.register("array", BuiltinToken::Array);
        self.register("ubound", BuiltinToken::UBound);
        self.register("lbound", BuiltinToken::LBound);
        self.register("filter", BuiltinToken::Filter);
        self.register("isarray", BuiltinToken::IsArray);

        // 检验函数
        self.register("isnumeric", BuiltinToken::IsNumeric);
        self.register("isdate", BuiltinToken::IsDate);
        self.register("isempty", BuiltinToken::IsEmpty);
        self.register("isnull", BuiltinToken::IsNull);
        self.register("isobject", BuiltinToken::IsObject);
        self.register("isnothing", BuiltinToken::IsNothing);
        self.register("typename", BuiltinToken::TypeName);
        self.register("vartype", BuiltinToken::VarType);

        // 交互函数
        self.register("msgbox", BuiltinToken::MsgBox);
        self.register("inputbox", BuiltinToken::InputBox);

        // 其他函数
        self.register("createobject", BuiltinToken::CreateObject);
        self.register("getobject", BuiltinToken::GetObject);
        self.register("eval", BuiltinToken::Eval);
        self.register("execute", BuiltinToken::Execute);
        self.register("rgb", BuiltinToken::RGB);
    }

    /// 注册单个 Token
    fn register(&mut self, name: &str, token: BuiltinToken) {
        self.map.insert(name.to_lowercase(), token);
    }

    /// 查找函数名对应的 Token
    pub fn lookup(&self, name: &str) -> Option<BuiltinToken> {
        self.map.get(&name.to_lowercase()).copied()
    }

    /// 检查是否为内置函数
    pub fn is_builtin(&self, name: &str) -> bool {
        self.map.contains_key(&name.to_lowercase())
    }
}

impl Default for TokenRegistry {
    fn default() -> Self {
        Self::new()
    }
}

/// 内置函数执行器
pub struct BuiltinExecutor;

impl BuiltinExecutor {
    /// 执行内置函数
    pub fn execute(
        token: BuiltinToken,
        args: &[Value],
    ) -> Result<Value, RuntimeError> {
        match token {
            // ========== 数学函数 ==========
            BuiltinToken::Abs => Self::math_unary(args, |n| n.abs()),
            BuiltinToken::Sqr => Self::math_unary(args, |n| n.sqrt()),
            BuiltinToken::Sin => Self::math_unary(args, |n| n.sin()),
            BuiltinToken::Cos => Self::math_unary(args, |n| n.cos()),
            BuiltinToken::Tan => Self::math_unary(args, |n| n.tan()),
            BuiltinToken::Atn => Self::math_unary(args, |n| n.atan()),
            BuiltinToken::Log => Self::math_unary(args, |n| n.ln()),
            BuiltinToken::Exp => Self::math_unary(args, |n| n.exp()),
            BuiltinToken::Int => Self::math_unary(args, |n| n.floor()),
            BuiltinToken::Fix => Self::math_unary(args, |n| n.trunc()),
            BuiltinToken::Sgn => Self::math_unary(args, |n| {
                if n > 0.0 { 1.0 } else if n < 0.0 { -1.0 } else { 0.0 }
            }),
            BuiltinToken::Round => {
                if args.is_empty() {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                let n = ValueConversion::to_number(&args[0]);
                let decimals = if args.len() > 1 {
                    ValueConversion::to_number(&args[1]) as i32
                } else {
                    0
                };
                let multiplier = 10_f64.powi(decimals);
                Ok(Value::Number((n * multiplier).round() / multiplier))
            }
            BuiltinToken::Rnd => {
                // 返回 0-1 的随机数
                Ok(Value::Number(rand::random::<f64>()))
            }

            // ========== 类型转换函数 ==========
            BuiltinToken::CStr => {
                if args.is_empty() {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                Ok(Value::String(ValueConversion::to_string(&args[0])))
            }
            BuiltinToken::CInt | BuiltinToken::CByte | BuiltinToken::CBool => {
                if args.is_empty() {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                Ok(Value::Number(ValueConversion::to_number(&args[0]) as i32 as f64))
            }
            BuiltinToken::CLng | BuiltinToken::CSng | BuiltinToken::CDbl => {
                if args.is_empty() {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                Ok(Value::Number(ValueConversion::to_number(&args[0])))
            }
            BuiltinToken::CDate => {
                // TODO: 实现日期转换
                if args.is_empty() {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                Ok(args[0].clone())
            }

            // ========== 字符串函数 ==========
            BuiltinToken::Len => {
                if args.is_empty() {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                let s = ValueConversion::to_string(&args[0]);
                Ok(Value::Number(s.len() as f64))
            }
            BuiltinToken::Trim => Self::string_unary(args, |s| s.trim().to_string()),
            BuiltinToken::LTrim => Self::string_unary(args, |s| s.trim_start().to_string()),
            BuiltinToken::RTrim => Self::string_unary(args, |s| s.trim_end().to_string()),
            BuiltinToken::UCase => Self::string_unary(args, |s| s.to_uppercase()),
            BuiltinToken::LCase => Self::string_unary(args, |s| s.to_lowercase()),
            BuiltinToken::Left => {
                if args.len() < 2 {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                let s = ValueConversion::to_string(&args[0]);
                let n = ValueConversion::to_number(&args[1]) as usize;
                let result = s.chars().take(n).collect::<String>();
                Ok(Value::String(result))
            }
            BuiltinToken::Right => {
                if args.len() < 2 {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                let s = ValueConversion::to_string(&args[0]);
                let n = ValueConversion::to_number(&args[1]) as usize;
                let result = s.chars().rev().take(n).collect::<String>()
                    .chars().rev().collect::<String>();
                Ok(Value::String(result))
            }
            BuiltinToken::Mid => {
                if args.len() < 2 {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                let s = ValueConversion::to_string(&args[0]);
                let start = (ValueConversion::to_number(&args[1]) as usize).saturating_sub(1);
                let length = if args.len() >= 3 {
                    ValueConversion::to_number(&args[2]) as usize
                } else {
                    s.len()
                };
                let result = s.chars().skip(start).take(length).collect::<String>();
                Ok(Value::String(result))
            }
            BuiltinToken::Asc => {
                if args.is_empty() {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                let s = ValueConversion::to_string(&args[0]);
                let code = s.chars().next().map(|c| c as u8 as f64).unwrap_or(0.0);
                Ok(Value::Number(code))
            }
            BuiltinToken::Chr => {
                if args.is_empty() {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                let n = ValueConversion::to_number(&args[0]) as u32;
                Ok(Value::String(
                    char::from_u32(n).unwrap_or('\0').to_string(),
                ))
            }
            BuiltinToken::InStr => {
                // InStr([start,] string1, string2 [, compare])
                if args.len() < 2 {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                let (string1, string2) = if args.len() >= 3 {
                    (ValueConversion::to_string(&args[1]), ValueConversion::to_string(&args[2]))
                } else {
                    (ValueConversion::to_string(&args[0]), ValueConversion::to_string(&args[1]))
                };
                let pos = string1.to_lowercase().find(&string2.to_lowercase())
                    .map(|i| i + 1)
                    .unwrap_or(0) as f64;
                Ok(Value::Number(pos))
            }
            BuiltinToken::Replace => {
                if args.len() < 3 {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                let s = ValueConversion::to_string(&args[0]);
                let find = ValueConversion::to_string(&args[1]);
                let replace = ValueConversion::to_string(&args[2]);
                let result = s.replace(&find, &replace);
                Ok(Value::String(result))
            }
            BuiltinToken::Split => {
                if args.is_empty() {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                let s = ValueConversion::to_string(&args[0]);
                let delimiter = if args.len() > 1 {
                    ValueConversion::to_string(&args[1])
                } else {
                    " ".to_string()
                };
                let parts: Vec<Value> = s.split(&delimiter)
                    .map(|p| Value::String(p.to_string()))
                    .collect();
                Ok(Value::Array(parts))
            }
            BuiltinToken::Join => {
                if args.is_empty() {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                match &args[0] {
                    Value::Array(arr) => {
                        let delimiter = if args.len() > 1 {
                            ValueConversion::to_string(&args[1])
                        } else {
                            " ".to_string()
                        };
                        let strings: Vec<String> = arr.iter()
                            .map(|v| ValueConversion::to_string(v))
                            .collect();
                        Ok(Value::String(strings.join(&delimiter)))
                    }
                    _ => Err(RuntimeError::TypeMismatch),
                }
            }
            BuiltinToken::Space => {
                if args.is_empty() {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                let n = ValueConversion::to_number(&args[0]) as usize;
                Ok(Value::String(" ".repeat(n)))
            }
            BuiltinToken::StrReverse => {
                if args.is_empty() {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                let s = ValueConversion::to_string(&args[0]);
                Ok(Value::String(s.chars().rev().collect()))
            }

            // ========== 数组函数 ==========
            BuiltinToken::UBound => {
                if args.is_empty() {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                match &args[0] {
                    Value::Array(arr) => {
                        let dimension = if args.len() > 1 {
                            ValueConversion::to_number(&args[1]) as usize
                        } else {
                            1
                        };
                        // VBScript 数组下标从 0 开始，但 UBound 返回最大索引
                        // 这里简化处理，假设是一维数组
                        Ok(Value::Number((arr.len().saturating_sub(1)) as f64))
                    }
                    _ => Err(RuntimeError::TypeMismatch),
                }
            }
            BuiltinToken::LBound => {
                // VBScript 数组默认下界是 0
                Ok(Value::Number(0.0))
            }
            BuiltinToken::Filter => {
                if args.len() < 2 {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                match &args[0] {
                    Value::Array(arr) => {
                        let criteria = ValueConversion::to_string(&args[1]);
                        let include = if args.len() > 2 {
                            ValueConversion::to_bool(&args[2])
                        } else {
                            true
                        };
                        let filtered: Vec<Value> = arr.iter()
                            .filter(|v| {
                                let s = ValueConversion::to_string(v);
                                if include {
                                    s.contains(&criteria)
                                } else {
                                    !s.contains(&criteria)
                                }
                            })
                            .cloned()
                            .collect();
                        Ok(Value::Array(filtered))
                    }
                    _ => Err(RuntimeError::TypeMismatch),
                }
            }
            BuiltinToken::IsArray => {
                if args.is_empty() {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                Ok(Value::Boolean(matches!(args[0], Value::Array(_))))
            }

            // ========== 日期时间函数 ==========
            BuiltinToken::Now => {
                use std::time::{SystemTime, UNIX_EPOCH};
                let now = SystemTime::now()
                    .duration_since(UNIX_EPOCH)
                    .unwrap_or_default()
                    .as_secs() as f64;
                Ok(Value::Number(now))
            }
            BuiltinToken::Date => {
                // 简化实现，返回当前日期的时间戳
                use std::time::{SystemTime, UNIX_EPOCH};
                let now = SystemTime::now()
                    .duration_since(UNIX_EPOCH)
                    .unwrap_or_default()
                    .as_secs() as f64;
                Ok(Value::Number(now))
            }
            BuiltinToken::Time => {
                use std::time::{SystemTime, UNIX_EPOCH};
                let now = SystemTime::now()
                    .duration_since(UNIX_EPOCH)
                    .unwrap_or_default()
                    .as_secs() as f64;
                Ok(Value::Number(now))
            }
            BuiltinToken::Year => {
                if args.is_empty() {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                // 简化实现，从时间戳提取年份
                let timestamp = ValueConversion::to_number(&args[0]);
                let year = 1970.0 + (timestamp / 31536000.0);
                Ok(Value::Number(year.floor()))
            }
            BuiltinToken::Month | BuiltinToken::Day | BuiltinToken::Hour | BuiltinToken::Minute | BuiltinToken::Second | BuiltinToken::WeekDay => {
                // 简化实现
                if args.is_empty() {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                let timestamp = ValueConversion::to_number(&args[0]);
                Ok(Value::Number(timestamp % 100.0))
            }

            // ========== 检验函数 ==========
            BuiltinToken::IsNumeric => {
                if args.is_empty() {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                let is_num = match &args[0] {
                    Value::Number(_) => true,
                    Value::Boolean(_) => true,
                    Value::String(s) => s.parse::<f64>().is_ok(),
                    Value::Empty => true,
                    Value::Null => false,
                    Value::Nothing => false,
                    Value::Array(_) => false,
                    Value::Object(_) => false,
                };
                Ok(Value::Boolean(is_num))
            }
            BuiltinToken::IsEmpty => {
                if args.is_empty() {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                Ok(Value::Boolean(matches!(args[0], Value::Empty)))
            }
            BuiltinToken::IsNull => {
                if args.is_empty() {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                Ok(Value::Boolean(matches!(args[0], Value::Null)))
            }
            BuiltinToken::IsObject => {
                if args.is_empty() {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                Ok(Value::Boolean(matches!(args[0], Value::Object(_))))
            }
            BuiltinToken::IsDate => {
                // TODO: 实现日期检测
                if args.is_empty() {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                Ok(Value::Boolean(false))
            }
            BuiltinToken::VarType => {
                if args.is_empty() {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                let vt = match &args[0] {
                    Value::Empty => 0,
                    Value::Null => 1,
                    Value::Number(_) => 2,
                    Value::String(_) => 8,
                    Value::Boolean(_) => 11,
                    Value::Array(_) => 8192,
                    Value::Object(_) => 9,
                    Value::Nothing => 1,
                };
                Ok(Value::Number(vt as f64))
            }
            BuiltinToken::TypeName => {
                if args.is_empty() {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                let tn = match &args[0] {
                    Value::Empty => "Empty",
                    Value::Null => "Null",
                    Value::Number(_) => "Double",
                    Value::String(_) => "String",
                    Value::Boolean(_) => "Boolean",
                    Value::Array(_) => "Variant()",
                    Value::Object(_) => "Object",
                    Value::Nothing => "Nothing",
                };
                Ok(Value::String(tn.to_string()))
            }

            // 未实现的函数返回错误
            _ => Err(RuntimeError::Generic(format!(
                "Function not yet implemented: {:?}",
                token
            ))),
        }
    }

    /// 一元数学函数辅助方法
    fn math_unary<F>(args: &[Value], f: F) -> Result<Value, RuntimeError>
    where
        F: FnOnce(f64) -> f64,
    {
        if args.is_empty() {
            return Err(RuntimeError::ArgumentCountMismatch);
        }
        let n = ValueConversion::to_number(&args[0]);
        Ok(Value::Number(f(n)))
    }

    /// 一元字符串函数辅助方法
    fn string_unary<F>(args: &[Value], f: F) -> Result<Value, RuntimeError>
    where
        F: FnOnce(&str) -> String,
    {
        if args.is_empty() {
            return Err(RuntimeError::ArgumentCountMismatch);
        }
        let s = ValueConversion::to_string(&args[0]);
        Ok(Value::String(f(&s)))
    }
}

/// 便捷函数：通过函数名直接执行内置函数
pub fn execute_builtin(name: &str, args: &[Value]) -> Option<Result<Value, RuntimeError>> {
    let registry = TokenRegistry::new();
    registry.lookup(name).map(|token| {
        BuiltinExecutor::execute(token, args)
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_token_registry() {
        let registry = TokenRegistry::new();

        // 测试查找
        assert_eq!(registry.lookup("abs"), Some(BuiltinToken::Abs));
        assert_eq!(registry.lookup("ABS"), Some(BuiltinToken::Abs)); // 不区分大小写
        assert_eq!(registry.lookup("len"), Some(BuiltinToken::Len));

        // 测试未注册的函数
        assert_eq!(registry.lookup("unknown"), None);
    }

    #[test]
    fn test_math_functions() {
        // Abs
        let result = BuiltinExecutor::execute(BuiltinToken::Abs, &[Value::Number(-5.0)]).unwrap();
        assert_eq!(result, Value::Number(5.0));

        // Sqr
        let result = BuiltinExecutor::execute(BuiltinToken::Sqr, &[Value::Number(16.0)]).unwrap();
        assert_eq!(result, Value::Number(4.0));

        // Int
        let result = BuiltinExecutor::execute(BuiltinToken::Int, &[Value::Number(3.7)]).unwrap();
        assert_eq!(result, Value::Number(3.0));
    }

    #[test]
    fn test_string_functions() {
        // Len
        let result = BuiltinExecutor::execute(BuiltinToken::Len, &[Value::String("hello".to_string())]).unwrap();
        assert_eq!(result, Value::Number(5.0));

        // UCase
        let result = BuiltinExecutor::execute(BuiltinToken::UCase, &[Value::String("hello".to_string())]).unwrap();
        assert_eq!(result, Value::String("HELLO".to_string()));

        // Trim
        let result = BuiltinExecutor::execute(BuiltinToken::Trim, &[Value::String("  hello  ".to_string())]).unwrap();
        assert_eq!(result, Value::String("hello".to_string()));
    }

    #[test]
    fn test_type_conversion() {
        // CInt
        let result = BuiltinExecutor::execute(BuiltinToken::CInt, &[Value::String("42".to_string())]).unwrap();
        assert_eq!(result, Value::Number(42.0));

        // CStr
        let result = BuiltinExecutor::execute(BuiltinToken::CStr, &[Value::Number(42.0)]).unwrap();
        assert_eq!(result, Value::String("42".to_string()));
    }

    #[test]
    fn test_inspection_functions() {
        // IsNumeric
        let result = BuiltinExecutor::execute(BuiltinToken::IsNumeric, &[Value::Number(42.0)]).unwrap();
        assert_eq!(result, Value::Boolean(true));

        let result = BuiltinExecutor::execute(BuiltinToken::IsNumeric, &[Value::String("abc".to_string())]).unwrap();
        assert_eq!(result, Value::Boolean(false));

        // IsEmpty
        let result = BuiltinExecutor::execute(BuiltinToken::IsEmpty, &[Value::Empty]).unwrap();
        assert_eq!(result, Value::Boolean(true));

        // TypeName
        let result = BuiltinExecutor::execute(BuiltinToken::TypeName, &[Value::Number(42.0)]).unwrap();
        assert_eq!(result, Value::String("Double".to_string()));
    }
}
