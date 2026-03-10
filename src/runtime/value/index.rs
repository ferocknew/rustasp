//! 值索引操作

use super::Value;
use crate::runtime::error::RuntimeError;

/// 值索引 trait
pub trait ValueIndex {
    /// 索引访问（0-based）
    fn index(&self, index: &Value) -> Result<Value, RuntimeError>;
}

impl ValueIndex for Value {
    fn index(&self, index: &Value) -> Result<Value, RuntimeError> {
        match self {
            Value::Array(arr) => {
                if let Value::Number(i) = index {
                    let i = *i as usize;
                    // 使用 flat_index 计算索引（一维数组）
                    let locked_arr = arr
                        .lock()
                        .map_err(|_| RuntimeError::Generic("Failed to lock array".to_string()))?;

                    match locked_arr.flat_index(&[i]) {
                        Some(flat_idx) => Ok(locked_arr.data[flat_idx].clone()),
                        None => Ok(Value::Empty),
                    }
                } else {
                    Ok(Value::Empty)
                }
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
                match obj
                    .lock()
                    .map_err(|_| RuntimeError::Generic("Failed to lock object".to_string()))?
                    .index(index)
                {
                    Ok(value) => Ok(value),
                    Err(_) => Ok(Value::Empty),
                }
            }
            _ => Ok(Value::Empty),
        }
    }
}
