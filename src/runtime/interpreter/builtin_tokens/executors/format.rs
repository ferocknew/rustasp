//! 格式化函数执行器
//!
//! 处理 FormatCurrency、FormatNumber、FormatPercent、RGB 等格式化函数

use crate::runtime::{RuntimeError, Value, ValueConversion};
use super::super::token::BuiltinToken;

pub fn execute(token: BuiltinToken, args: &[Value]) -> Result<Option<Value>, RuntimeError> {
    let result = match token {
        BuiltinToken::RGB => {
            // RGB(red, green, blue) - 返回 RGB 颜色值
            if args.len() < 3 {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let red = ValueConversion::to_number(&args[0]) as u32;
            let green = ValueConversion::to_number(&args[1]) as u32;
            let blue = ValueConversion::to_number(&args[2]) as u32;

            // RGB 值: &HBBGGRR (蓝蓝绿绿红红)
            let rgb_value = (red & 0xFF) | ((green & 0xFF) << 8) | ((blue & 0xFF) << 16);
            Value::Number(rgb_value as f64)
        }
        BuiltinToken::FormatCurrency => {
            // FormatCurrency(expression[, numdigits[, leadingdigit[, parenthesis[, groupdigit]]]])
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let value = ValueConversion::to_number(&args[0]);
            let num_digits = if args.len() >= 2 {
                ValueConversion::to_number(&args[1]) as usize
            } else {
                2
            };
            Value::String(format!("{:.precision$}", value, precision = num_digits))
        }
        BuiltinToken::FormatNumber => {
            // FormatNumber(expression[, numdigits[, leadingdigit[, parenthesis[, groupdigit]]]])
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let value = ValueConversion::to_number(&args[0]);
            let num_digits = if args.len() >= 2 {
                ValueConversion::to_number(&args[1]) as usize
            } else {
                2
            };
            Value::String(format!("{:.precision$}", value, precision = num_digits))
        }
        BuiltinToken::FormatPercent => {
            // FormatPercent(expression[, numdigits[, leadingdigit[, parenthesis[, groupdigit]]]])
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let value = ValueConversion::to_number(&args[0]) * 100.0;
            let num_digits = if args.len() >= 2 {
                ValueConversion::to_number(&args[1]) as usize
            } else {
                2
            };
            Value::String(format!("{:.precision$}%", value, precision = num_digits))
        }
        _ => return Ok(None),
    };
    Ok(Some(result))
}
