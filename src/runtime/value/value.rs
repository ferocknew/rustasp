//! VBScript 值类型定义

use std::collections::HashMap;

/// VBScript 值类型
#[derive(Debug, Clone, PartialEq)]
pub enum Value {
    /// Empty - 未初始化
    Empty,
    /// Null - 无效数据
    Null,
    /// Nothing - 对象引用为空
    Nothing,
    /// 布尔值
    Boolean(bool),
    /// 数字
    Number(f64),
    /// 字符串
    String(String),
    /// 数组
    Array(Vec<Value>),
    /// 字典/对象
    Object(HashMap<String, Value>),
}

impl Default for Value {
    fn default() -> Self {
        Value::Empty
    }
}
