//! 值比较操作

use crate::ast::BinaryOp;
use super::{Value, ValueConversion};

/// 值比较 trait
pub trait ValueCompare {
    /// 比较
    fn compare(&self, op: BinaryOp, right: &Value) -> Value;
}

impl ValueCompare for Value {
    fn compare(&self, op: BinaryOp, right: &Value) -> Value {
        match op {
            BinaryOp::Eq => {
                Value::Boolean(self.equals(right))
            }
            BinaryOp::Ne => {
                Value::Boolean(!self.equals(right))
            }
            BinaryOp::Lt => {
                match (self, right) {
                    (Value::Number(a), Value::Number(b)) => Value::Boolean(a < b),
                    (Value::String(a), Value::String(b)) => Value::Boolean(a < b),
                    _ => Value::Boolean(self.to_number() < right.to_number()),
                }
            }
            BinaryOp::Le => {
                match (self, right) {
                    (Value::Number(a), Value::Number(b)) => Value::Boolean(a <= b),
                    (Value::String(a), Value::String(b)) => Value::Boolean(a <= b),
                    _ => Value::Boolean(self.to_number() <= right.to_number()),
                }
            }
            BinaryOp::Gt => {
                match (self, right) {
                    (Value::Number(a), Value::Number(b)) => Value::Boolean(a > b),
                    (Value::String(a), Value::String(b)) => Value::Boolean(a > b),
                    _ => Value::Boolean(self.to_number() > right.to_number()),
                }
            }
            BinaryOp::Ge => {
                match (self, right) {
                    (Value::Number(a), Value::Number(b)) => Value::Boolean(a >= b),
                    (Value::String(a), Value::String(b)) => Value::Boolean(a >= b),
                    _ => Value::Boolean(self.to_number() >= right.to_number()),
                }
            }
            BinaryOp::And => {
                Value::Boolean(self.is_truthy() && right.is_truthy())
            }
            BinaryOp::Or => {
                Value::Boolean(self.is_truthy() || right.is_truthy())
            }
            BinaryOp::Xor => {
                Value::Boolean(self.is_truthy() != right.is_truthy())
            }
            BinaryOp::Is => {
                Value::Boolean(matches!((self, right),
                    (Value::Nothing, Value::Nothing) |
                    (Value::Null, Value::Null) |
                    (Value::Empty, Value::Empty)
                ))
            }
            _ => Value::Boolean(false),
        }
    }
}

impl Value {
    /// 值相等比较
    fn equals(&self, other: &Value) -> bool {
        match (self, other) {
            (Value::Empty, Value::Empty) => true,
            (Value::Null, Value::Null) => true,
            (Value::Nothing, Value::Nothing) => true,
            (Value::Boolean(a), Value::Boolean(b)) => a == b,
            (Value::Number(a), Value::Number(b)) => a == b,
            (Value::String(a), Value::String(b)) => a == b,
            (Value::Array(a), Value::Array(b)) => a == b,
            (Value::Object(a), Value::Object(b)) => a == b,
            // 弱类型比较
            (Value::Number(a), Value::String(b)) => {
                if let Ok(n) = b.parse::<f64>() {
                    *a == n
                } else {
                    false
                }
            }
            (Value::String(a), Value::Number(b)) => {
                if let Ok(n) = a.parse::<f64>() {
                    n == *b
                } else {
                    false
                }
            }
            _ => false,
        }
    }
}
