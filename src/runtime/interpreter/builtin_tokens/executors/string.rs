//! 字符串函数执行器

use crate::runtime::{RuntimeError, Value, ValueConversion};
use super::super::token::BuiltinToken;

pub fn execute(token: BuiltinToken, args: &[Value]) -> Result<Option<Value>, RuntimeError> {
    let result = match token {
        BuiltinToken::Len => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            Value::Number(ValueConversion::to_string(&args[0]).len() as f64)
        }
        BuiltinToken::Trim => string_unary(args, |s| s.trim().to_string())?,
        BuiltinToken::LTrim => string_unary(args, |s| s.trim_start().to_string())?,
        BuiltinToken::RTrim => string_unary(args, |s| s.trim_end().to_string())?,
        BuiltinToken::UCase => string_unary(args, |s| s.to_uppercase())?,
        BuiltinToken::LCase => string_unary(args, |s| s.to_lowercase())?,
        BuiltinToken::Left => {
            if args.len() < 2 {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let s = ValueConversion::to_string(&args[0]);
            let n = ValueConversion::to_number(&args[1]) as usize;
            Value::String(s.chars().take(n).collect::<String>())
        }
        BuiltinToken::Right => {
            if args.len() < 2 {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let s = ValueConversion::to_string(&args[0]);
            let n = ValueConversion::to_number(&args[1]) as usize;
            Value::String(s.chars().rev().take(n).collect::<String>().chars().rev().collect::<String>())
        }
        BuiltinToken::Mid => {
            if args.len() < 2 {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let s = ValueConversion::to_string(&args[0]);
            let start = ValueConversion::to_number(&args[1]) as usize;
            let length = args.get(2).map(|v| ValueConversion::to_number(v) as usize).unwrap_or(s.len());
            Value::String(s.chars().skip(start.saturating_sub(1)).take(length).collect::<String>())
        }
        BuiltinToken::Asc => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let s = ValueConversion::to_string(&args[0]);
            Value::Number(s.chars().next().map(|c| c as u8 as f64).unwrap_or(0.0))
        }
        BuiltinToken::Chr => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let n = ValueConversion::to_number(&args[0]) as u32;
            Value::String(char::from_u32(n).unwrap_or('\0').to_string())
        }
        BuiltinToken::InStr => {
            if args.len() < 2 {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let (s1, s2) = if args.len() >= 3 {
                (ValueConversion::to_string(&args[1]), ValueConversion::to_string(&args[2]))
            } else {
                (ValueConversion::to_string(&args[0]), ValueConversion::to_string(&args[1]))
            };
            Value::Number(s1.to_lowercase().find(&s2.to_lowercase()).map(|i| (i + 1) as f64).unwrap_or(0.0))
        }
        BuiltinToken::Replace => {
            if args.len() < 3 {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let s = ValueConversion::to_string(&args[0]);
            let find = ValueConversion::to_string(&args[1]);
            let replace = ValueConversion::to_string(&args[2]);
            Value::String(s.replace(&find, &replace))
        }
        BuiltinToken::Split => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let s = ValueConversion::to_string(&args[0]);
            let delimiter = args.get(1).map(|v| ValueConversion::to_string(v)).unwrap_or(" ".to_string());
            Value::Array(s.split(&delimiter).map(|p| Value::String(p.to_string())).collect())
        }
        BuiltinToken::Join => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            match &args[0] {
                Value::Array(arr) => {
                    let delimiter = args.get(1).map(|v| ValueConversion::to_string(v)).unwrap_or(" ".to_string());
                    Value::String(arr.iter().map(|v| ValueConversion::to_string(v)).collect::<Vec<_>>().join(&delimiter))
                }
                _ => return Err(RuntimeError::TypeMismatch("Expected array".to_string())),
            }
        }
        BuiltinToken::Space => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            Value::String(" ".repeat(ValueConversion::to_number(&args[0]) as usize))
        }
        BuiltinToken::StrReverse => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            Value::String(ValueConversion::to_string(&args[0]).chars().rev().collect::<String>())
        }
        _ => return Ok(None),
    };
    Ok(Some(result))
}

fn string_unary<F>(args: &[Value], f: F) -> Result<Value, RuntimeError>
where
    F: FnOnce(&str) -> String,
{
    if args.is_empty() {
        return Err(RuntimeError::ArgumentCountMismatch);
    }
    Ok(Value::String(f(&ValueConversion::to_string(&args[0]))))
}
