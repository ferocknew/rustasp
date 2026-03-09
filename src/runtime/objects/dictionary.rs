//! Dictionary 对象 - 用于用户自定义字典/哈希表

use crate::runtime::{BuiltinObject, RuntimeError, Value, ValueConversion};
use std::collections::HashMap;
use std::sync::{Arc, Mutex};

/// Dictionary 对象
#[derive(Debug, Clone)]
pub struct Dictionary {
    /// 数据存储
    data: HashMap<String, Value>,
}

impl Dictionary {
    /// 创建新 Dictionary
    pub fn new() -> Self {
        Dictionary {
            data: HashMap::new(),
        }
    }

    /// 从 HashMap 创建
    pub fn from_hashmap(data: HashMap<String, Value>) -> Self {
        Dictionary { data }
    }

    /// 获取值
    pub fn get(&self, key: &str) -> Option<&Value> {
        self.data.get(&key.to_lowercase())
    }

    /// 设置值
    pub fn set(&mut self, key: String, value: Value) {
        self.data.insert(key.to_lowercase(), value);
    }

    /// 删除值
    pub fn remove(&mut self, key: &str) -> Option<Value> {
        self.data.remove(&key.to_lowercase())
    }

    /// 获取所有键
    pub fn keys(&self) -> Vec<String> {
        self.data.keys().cloned().collect()
    }

    /// 获取所有值
    pub fn values(&self) -> Vec<Value> {
        self.data.values().cloned().collect()
    }

    /// 清空
    pub fn clear(&mut self) {
        self.data.clear();
    }

    /// 获取数量
    pub fn count(&self) -> usize {
        self.data.len()
    }

    /// 检查键是否存在
    pub fn exists(&self, key: &str) -> bool {
        self.data.contains_key(&key.to_lowercase())
    }

    /// 获取内部 HashMap 的引用
    pub fn as_hashmap(&self) -> &HashMap<String, Value> {
        &self.data
    }

    /// 获取内部 HashMap 的可变引用
    pub fn as_hashmap_mut(&mut self) -> &mut HashMap<String, Value> {
        &mut self.data
    }
}

impl Default for Dictionary {
    fn default() -> Self {
        Self::new()
    }
}

impl BuiltinObject for Dictionary {
    fn as_any(&self) -> &dyn std::any::Any {
        self
    }

    fn as_any_mut(&mut self) -> &mut dyn std::any::Any {
        self
    }

    fn get_property(&self, name: &str) -> Result<Value, RuntimeError> {
        match name.to_lowercase().as_str() {
            "count" => Ok(Value::Number(self.data.len() as f64)),
            "keys" => {
                let keys: Vec<Value> = self.data.keys().map(|k| Value::String(k.clone())).collect();
                Ok(Value::Array(Arc::new(Mutex::new(crate::runtime::VbsArray::from_vec(keys)))))
            }
            "items" => {
                Ok(Value::Array(Arc::new(Mutex::new(crate::runtime::VbsArray::from_vec(self.data.values().cloned().collect())))))
            }
            "key" | "item" => {
                // 这些属性需要参数，返回 Empty
                Ok(Value::Empty)
            }
            // 尝试作为键访问
            _ => {
                Ok(self.data.get(&name.to_lowercase()).cloned().unwrap_or(Value::Empty))
            }
        }
    }

    fn set_property(&mut self, name: &str, value: Value) -> Result<(), RuntimeError> {
        // 设置属性就是设置键值对
        self.data.insert(name.to_lowercase(), value);
        Ok(())
    }

    fn call_method(&mut self, name: &str, args: Vec<Value>) -> Result<Value, RuntimeError> {
        match name.to_lowercase().as_str() {
            "add" | "set" => {
                if args.len() >= 2 {
                    let key = ValueConversion::to_string(&args[0]);
                    let value = args[1].clone();
                    self.data.insert(key.to_lowercase(), value);
                }
                Ok(Value::Empty)
            }
            "remove" => {
                if !args.is_empty() {
                    let key = ValueConversion::to_string(&args[0]);
                    self.data.remove(&key.to_lowercase());
                }
                Ok(Value::Empty)
            }
            "removeall" => {
                self.data.clear();
                Ok(Value::Empty)
            }
            "exists" => {
                if args.is_empty() {
                    return Ok(Value::Boolean(false));
                }
                let key = ValueConversion::to_string(&args[0]);
                Ok(Value::Boolean(self.data.contains_key(&key.to_lowercase())))
            }
            "keys" => {
                let keys: Vec<Value> = self.data.keys().map(|k| Value::String(k.clone())).collect();
                Ok(Value::Array(Arc::new(Mutex::new(crate::runtime::VbsArray::from_vec(keys)))))
            }
            "items" => {
                Ok(Value::Array(Arc::new(Mutex::new(crate::runtime::VbsArray::from_vec(self.data.values().cloned().collect())))))
            }
            "item" => {
                if args.is_empty() {
                    return Ok(Value::Empty);
                }
                let key = ValueConversion::to_string(&args[0]);
                Ok(self.data.get(&key.to_lowercase()).cloned().unwrap_or(Value::Empty))
            }
            _ => Err(RuntimeError::MethodNotFound(name.to_string())),
        }
    }

    fn index(&self, key: &Value) -> Result<Value, RuntimeError> {
        let key_str = ValueConversion::to_string(key).to_lowercase();
        Ok(self.data.get(&key_str).cloned().unwrap_or(Value::Empty))
    }
}
