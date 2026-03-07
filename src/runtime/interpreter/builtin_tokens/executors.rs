//! 内置函数执行器

use crate::runtime::{RuntimeError, Value};
use super::token::BuiltinToken;

mod math;
mod string;
mod conversion;
mod array;
mod datetime;
mod inspection;

/// 内置函数执行器
pub struct BuiltinExecutor;

impl BuiltinExecutor {
    /// 执行内置函数
    pub fn execute(token: BuiltinToken, args: &[Value]) -> Result<Value, RuntimeError> {
        // 数学函数
        if let Some(result) = math::execute(token, args)? {
            return Ok(result);
        }

        // 类型转换函数
        if let Some(result) = conversion::execute(token, args)? {
            return Ok(result);
        }

        // 字符串函数
        if let Some(result) = string::execute(token, args)? {
            return Ok(result);
        }

        // 数组函数
        if let Some(result) = array::execute(token, args)? {
            return Ok(result);
        }

        // 日期时间函数
        if let Some(result) = datetime::execute(token, args)? {
            return Ok(result);
        }

        // 检验函数
        if let Some(result) = inspection::execute(token, args)? {
            return Ok(result);
        }

        Err(RuntimeError::Generic(format!(
            "Function not yet implemented: {:?}",
            token
        )))
    }
}
