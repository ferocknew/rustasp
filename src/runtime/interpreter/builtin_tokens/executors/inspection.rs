//! 检验函数执行器

use crate::runtime::{RuntimeError, Value};
use super::super::token::BuiltinToken;

pub fn execute(token: BuiltinToken, args: &[Value]) -> Result<Option<Value>, RuntimeError> {
    let result = match token {
        BuiltinToken::IsNumeric => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            Value::Boolean(is_numeric(&args[0]))
        }
        BuiltinToken::IsEmpty => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            Value::Boolean(matches!(args[0], Value::Empty))
        }
        BuiltinToken::IsNull => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            Value::Boolean(matches!(args[0], Value::Null))
        }
        BuiltinToken::IsObject => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            Value::Boolean(matches!(args[0], Value::Object(_)))
        }
        BuiltinToken::IsDate => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            Value::Boolean(false)
        }
        BuiltinToken::VarType => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            Value::Number(var_type(&args[0]) as f64)
        }
        BuiltinToken::TypeName => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            Value::String(type_name(&args[0]).to_string())
        }
        _ => return Ok(None),
    };
    Ok(Some(result))
}

fn is_numeric(v: &Value) -> bool {
    matches!(v, Value::Number(_) | Value::Boolean(_) | Value::Empty)
        || matches!(v, Value::String(s) if s.parse::<f64>().is_ok())
}

fn var_type(v: &Value) -> u16 {
    match v {
        Value::Empty => 0,
        Value::Null => 1,
        Value::Number(_) => 2,
        Value::String(_) => 8,
        Value::Boolean(_) => 11,
        Value::Array(_) => 8192,
        Value::Object(_) => 9,
        Value::Nothing => 1,
    }
}

fn type_name(v: &Value) -> &'static str {
    match v {
        Value::Empty => "Empty",
        Value::Null => "Null",
        Value::Number(_) => "Double",
        Value::String(_) => "String",
        Value::Boolean(_) => "Boolean",
        Value::Array(_) => "Variant()",
        Value::Object(_) => "Object",
        Value::Nothing => "Nothing",
    }
}
