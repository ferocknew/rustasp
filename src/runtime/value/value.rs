//! VBScript 值类型定义

use std::collections::HashMap;
use std::sync::{Arc, Mutex};

use super::super::{objects::Dictionary, ObjectRef};

/// VBScript 值类型
#[derive(Debug)]
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
    ///
    /// 使用 Arc<Mutex<dyn BuiltinObject>> 实现共享所有权：
    /// - clone() 只是增加 Arc 引用计数，不复制底层对象
    /// - Mutex 提供内部可变性（方法调用需要 &mut self）
    Object(ObjectRef),
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
            // 对象比较：通过 Arc::ptr_eq 比较是否同一个对象
            (Value::Object(a), Value::Object(b)) => {
                Arc::ptr_eq(a, b)
            }
            _ => false,
        }
    }
}

impl Clone for Value {
    fn clone(&self) -> Self {
        match self {
            Value::Empty => Value::Empty,
            Value::Null => Value::Null,
            Value::Nothing => Value::Nothing,
            Value::Boolean(v) => Value::Boolean(*v),
            Value::Number(v) => Value::Number(*v),
            Value::String(v) => Value::String(v.clone()),
            Value::Array(v) => Value::Array(v.clone()),
            // Object: Arc::clone 只增加引用计数，不复制底层对象
            Value::Object(obj) => Value::Object(Arc::clone(obj)),
        }
    }
}

impl Value {
    /// 创建字典对象
    pub fn new_dictionary() -> Self {
        Value::Object(Arc::new(Mutex::new(Dictionary::new())))
    }

    /// 从 HashMap 创建字典对象
    pub fn from_hashmap(map: HashMap<String, Value>) -> Self {
        Value::Object(Arc::new(Mutex::new(Dictionary::from_hashmap(map))))
    }

    /// 尝试获取字典对象的可变引用（通过 Mutex lock）
    pub fn as_dictionary_mut(&mut self) -> Option<std::sync::MutexGuard<'_, Dictionary>> {
        match self {
            Value::Object(_obj) => {
                // 直接 lock 并尝试向下转型
                // 使用 unsafe 来延长生命周期（这里可以安全，因为 Mutex 的生命周期由 self 保证）
                // 但更好的方法是直接返回 None，让调用者通过其他方式访问
                None
            }
            _ => None,
        }
    }

    /// 尝试获取字典对象的引用（通过 Mutex lock）
    pub fn as_dictionary(&self) -> Option<&Dictionary> {
        // 由于 MutexGuard 的生命周期限制，这里无法直接返回引用
        // 返回 None，让调用者使用其他方式访问字典
        None
    }

    /// 检查是否是字典类型
    pub fn is_dictionary(&self) -> bool {
        match self {
            Value::Object(obj) => {
                if let Ok(guard) = obj.lock() {
                    guard.as_any().is::<Dictionary>()
                } else {
                    false
                }
            }
            _ => false,
        }
    }
}

impl Default for Value {
    fn default() -> Self {
        Value::Empty
    }
}
