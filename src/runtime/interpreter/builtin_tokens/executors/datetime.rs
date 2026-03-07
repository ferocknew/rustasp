//! 日期时间函数执行器

use crate::runtime::{RuntimeError, Value, ValueConversion};
use super::super::token::BuiltinToken;

pub fn execute(token: BuiltinToken, args: &[Value]) -> Result<Option<Value>, RuntimeError> {
    let result = match token {
        BuiltinToken::Now | BuiltinToken::Date | BuiltinToken::Time => {
            use std::time::{SystemTime, UNIX_EPOCH};
            let now = SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .unwrap_or_default()
                .as_secs() as f64;
            Value::Number(now)
        }
        BuiltinToken::Year => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let timestamp = ValueConversion::to_number(&args[0]);
            Value::Number((1970.0 + timestamp / 31536000.0).floor())
        }
        BuiltinToken::Month | BuiltinToken::Day | BuiltinToken::Hour | 
        BuiltinToken::Minute | BuiltinToken::Second | BuiltinToken::WeekDay => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let timestamp = ValueConversion::to_number(&args[0]);
            Value::Number(timestamp % 100.0)
        }
        _ => return Ok(None),
    };
    Ok(Some(result))
}
