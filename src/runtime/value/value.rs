//! VBScript 值类型定义

use std::collections::HashMap;

use super::super::{BuiltinObject, objects::Dictionary};

/// VBScript 值类型
#[derive(Debug, Clone)]
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
    /// 对象（包括内置对象和字典）
    Object(Box<dyn BuiltinObject>),
}

impl PartialEq for Value {
    fn eq(&self, other: &Self) -> bool {
        match (self, other) {
            (Value::Empty, Value::Empty) => true,
            (Value::Null, Value::Null) => true,
            (Value::Nothing, Value::Nothing) => true,
            (Value::Boolean(a), Value::Boolean(b)) => a == b,
            (Value::Number(a), Value::Number(b)) => a == b,
            (Value::String(a), Value::String(b)) => a == b,
            (Value::Array(a), Value::Array(b)) => a == b,
            // 对象比较：只比较是否同一类型，不比较内容
            (Value::Object(_), Value::Object(_)) => false,
            _ => false,
        }
    }
}

impl Value {
    /// 创建字典对象
    pub fn new_dictionary() -> Self {
        Value::Object(Box::new(Dictionary::new()))
    }

    /// 从 HashMap 创建字典对象
    pub fn from_hashmap(map: HashMap<String, Value>) -> Self {
        Value::Object(Box::new(Dictionary::from_hashmap(map)))
    }

    /// 尝试获取字典对象的可变引用
    pub fn as_dictionary_mut(&mut self) -> Option<&mut Dictionary> {
        match self {
            Value::Object(obj) => obj.as_any_mut().downcast_mut::<Dictionary>(),
            _ => None,
        }
    }

    /// 尝试获取字典对象的引用
    pub fn as_dictionary(&self) -> Option<&Dictionary> {
        match self {
            Value::Object(obj) => obj.as_any().downcast_ref::<Dictionary>(),
            _ => None,
        }
    }

    /// 检查是否是字典类型
    pub fn is_dictionary(&self) -> bool {
        self.as_dictionary().is_some()
    }
}

impl Default for Value {
    fn default() -> Self {
        Value::Empty
    }
}
