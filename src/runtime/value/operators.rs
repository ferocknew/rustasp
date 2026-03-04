//! 值运算操作

use crate::ast::BinaryOp;
use super::{Value, ValueConversion};
use crate::runtime::error::RuntimeError;

/// 值运算 trait
pub trait ValueOps {
    /// 二元运算
    fn binary_op(&self, op: BinaryOp, right: &Value) -> Result<Value, RuntimeError>;
}

impl ValueOps for Value {
    fn binary_op(&self, op: BinaryOp, right: &Value) -> Result<Value, RuntimeError> {
        match op {
            BinaryOp::Add => {
                match (self, right) {
                    (Value::Number(a), Value::Number(b)) => Ok(Value::Number(a + b)),
                    (Value::String(a), Value::String(b)) => Ok(Value::String(format!("{}{}", a, b))),
                    (Value::String(a), Value::Number(b)) => Ok(Value::String(format!("{}{}", a, b))),
                    (Value::Number(a), Value::String(b)) => Ok(Value::String(format!("{}{}", a, b))),
                    _ => Ok(Value::String(format!("{}{}", self.to_string(), right.to_string()))),
                }
            }
            BinaryOp::Sub => {
                let a = self.to_number();
                let b = right.to_number();
                Ok(Value::Number(a - b))
            }
            BinaryOp::Mul => {
                let a = self.to_number();
                let b = right.to_number();
                Ok(Value::Number(a * b))
            }
            BinaryOp::Div => {
                let a = self.to_number();
                let b = right.to_number();
                if b == 0.0 {
                    Err(RuntimeError::DivisionByZero)
                } else {
                    Ok(Value::Number(a / b))
                }
            }
            BinaryOp::IntDiv => {
                let a = self.to_number();
                let b = right.to_number();
                if b == 0.0 {
                    Err(RuntimeError::DivisionByZero)
                } else {
                    Ok(Value::Number((a / b).trunc()))
                }
            }
            BinaryOp::Mod => {
                let a = self.to_number();
                let b = right.to_number();
                if b == 0.0 {
                    Err(RuntimeError::DivisionByZero)
                } else {
                    Ok(Value::Number(a % b))
                }
            }
            BinaryOp::Pow => {
                let a = self.to_number();
                let b = right.to_number();
                Ok(Value::Number(a.powf(b)))
            }
            BinaryOp::Concat => {
                Ok(Value::String(format!("{}{}", self.to_string(), right.to_string())))
            }
            _ => Err(RuntimeError::Generic(format!("Unsupported operator: {:?}", op))),
        }
    }
}
