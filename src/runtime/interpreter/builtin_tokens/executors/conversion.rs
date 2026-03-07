//! 类型转换函数执行器

use crate::runtime::{RuntimeError, Value, ValueConversion};
use super::super::token::BuiltinToken;

pub fn execute(token: BuiltinToken, args: &[Value]) -> Result<Option<Value>, RuntimeError> {
    let result = match token {
        BuiltinToken::CStr => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            Value::String(ValueConversion::to_string(&args[0]))
        }
        BuiltinToken::CInt | BuiltinToken::CByte | BuiltinToken::CBool => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            Value::Number(ValueConversion::to_number(&args[0]) as i32 as f64)
        }
        BuiltinToken::CLng | BuiltinToken::CSng | BuiltinToken::CDbl | BuiltinToken::CCur => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            Value::Number(ValueConversion::to_number(&args[0]))
        }
        BuiltinToken::CDate => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            args[0].clone()
        }
        _ => return Ok(None),
    };
    Ok(Some(result))
}
