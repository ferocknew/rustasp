//! 数学函数执行器

use crate::runtime::{RuntimeError, Value, ValueConversion};
use super::super::token::BuiltinToken;

pub fn execute(token: BuiltinToken, args: &[Value]) -> Result<Option<Value>, RuntimeError> {
    let result = match token {
        BuiltinToken::Abs => math_unary(args, |n| n.abs())?,
        BuiltinToken::Sqr => math_unary(args, |n| n.sqrt())?,
        BuiltinToken::Sin => math_unary(args, |n| n.sin())?,
        BuiltinToken::Cos => math_unary(args, |n| n.cos())?,
        BuiltinToken::Tan => math_unary(args, |n| n.tan())?,
        BuiltinToken::Atn => math_unary(args, |n| n.atan())?,
        BuiltinToken::Log => math_unary(args, |n| n.ln())?,
        BuiltinToken::Exp => math_unary(args, |n| n.exp())?,
        BuiltinToken::Int => math_unary(args, |n| n.floor())?,
        BuiltinToken::Fix => math_unary(args, |n| n.trunc())?,
        BuiltinToken::Sgn => math_unary(args, |n| {
            if n > 0.0 { 1.0 } else if n < 0.0 { -1.0 } else { 0.0 }
        })?,
        BuiltinToken::Round => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let n = ValueConversion::to_number(&args[0]);
            let decimals = args.get(1).map(|v| ValueConversion::to_number(v) as i32).unwrap_or(0);
            let multiplier = 10_f64.powi(decimals);
            Value::Number((n * multiplier).round() / multiplier)
        }
        BuiltinToken::Rnd => Value::Number(rand::random::<f64>()),
        _ => return Ok(None),
    };
    Ok(Some(result))
}

fn math_unary<F>(args: &[Value], f: F) -> Result<Value, RuntimeError>
where
    F: FnOnce(f64) -> f64,
{
    if args.is_empty() {
        return Err(RuntimeError::ArgumentCountMismatch);
    }
    Ok(Value::Number(f(ValueConversion::to_number(&args[0]))))
}
