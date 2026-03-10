//! 类型转换函数执行器

use super::super::token::BuiltinToken;
use crate::runtime::{RuntimeError, Value, ValueConversion};

pub fn execute(token: BuiltinToken, args: &[Value]) -> Result<Option<Value>, RuntimeError> {
    let result = match token {
        BuiltinToken::CStr => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            Value::String(ValueConversion::to_string(&args[0]))
        }
        BuiltinToken::CInt => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let num = ValueConversion::to_number(&args[0]);
            let rounded = num.round() as i32 as f64;
            Value::Number(rounded)
        }
        BuiltinToken::CLng => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let num = ValueConversion::to_number(&args[0]);
            let rounded = num.round() as i64 as f64;
            Value::Number(rounded)
        }
        BuiltinToken::CByte | BuiltinToken::CBool => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            Value::Number(ValueConversion::to_number(&args[0]) as i32 as f64)
        }
        BuiltinToken::CSng | BuiltinToken::CDbl | BuiltinToken::CCur => {
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
        BuiltinToken::Hex => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let n = ValueConversion::to_number(&args[0]) as i64;
            Value::String(format!("{:X}", n))
        }
        BuiltinToken::Oct => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let n = ValueConversion::to_number(&args[0]) as i64;
            Value::String(format!("{:o}", n))
        }
        _ => return Ok(None),
    };
    Ok(Some(result))
}
