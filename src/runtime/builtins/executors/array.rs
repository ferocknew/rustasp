//! 数组函数执行器

use crate::runtime::{RuntimeError, Value, ValueConversion};
use super::super::token::BuiltinToken;

pub fn execute(token: BuiltinToken, args: &[Value]) -> Result<Option<Value>, RuntimeError> {
    let result = match token {
        BuiltinToken::UBound => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            match &args[0] {
                Value::Array(arr) => Value::Number((arr.len().saturating_sub(1)) as f64),
                _ => return Err(RuntimeError::TypeMismatch("Expected array".to_string())),
            }
        }
        BuiltinToken::LBound => Value::Number(0.0),
        BuiltinToken::Array => {
            // Array 函数：将参数转换为数组
            Value::Array(args.to_vec())
        }
        BuiltinToken::Filter => {
            if args.len() < 2 {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            match &args[0] {
                Value::Array(arr) => {
                    let criteria = ValueConversion::to_string(&args[1]);
                    let include = args.get(2).map(|v| ValueConversion::to_bool(v)).unwrap_or(true);
                    Value::Array(arr.iter()
                        .filter(|v| {
                            let s = ValueConversion::to_string(&**v);
                            if include { s.contains(&criteria) } else { !s.contains(&criteria) }
                        })
                        .cloned()
                        .collect())
                }
                _ => return Err(RuntimeError::TypeMismatch("Expected array".to_string())),
            }
        }
        BuiltinToken::IsArray => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            Value::Boolean(matches!(args[0], Value::Array(_)))
        }
        _ => return Ok(None),
    };
    Ok(Some(result))
}
