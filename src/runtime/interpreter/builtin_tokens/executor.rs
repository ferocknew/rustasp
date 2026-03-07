
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
