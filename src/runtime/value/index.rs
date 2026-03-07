//! 值索引操作

use super::Value;
use crate::runtime::error::RuntimeError;

/// 值索引 trait
pub trait ValueIndex {
    /// 索引访问（ASP 索引从 1 开始）
    fn index(&self, index: &Value) -> Result<Value, RuntimeError>;
}

impl ValueIndex for Value {
    fn index(&self, index: &Value) -> Result<Value, RuntimeError> {
        match self {
            Value::Array(arr) => {
                if let Value::Number(i) = index {
                    let i = *i as usize;
                    // ASP 索引从 1 开始
                    if i >= 1 && i <= arr.len() {
                        return Ok(arr[i - 1].clone());
                    }
                }
                Ok(Value::Empty)
            }
            Value::String(s) => {
                // ASP 中字符串的索引访问：对于单值，(1) 返回字符串本身
                if let Value::Number(i) = index {
                    if *i == 1.0 {
                        return Ok(Value::String(s.clone()));
                    }
                }
                Ok(Value::Empty)
            }
            Value::Object(obj) => {
                // 使用 BuiltinObject trait 的 index 方法
                match obj.index(index) {
                    Ok(value) => Ok(value),
                    Err(_) => Ok(Value::Empty),
                }
            }
            _ => Ok(Value::Empty),
        }
    }
}
