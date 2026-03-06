//! Session 对象

use crate::runtime::{RuntimeError, Value, ValueConversion};
use std::collections::HashMap;
use std::sync::{Arc, Mutex};

/// Session 存储
pub type SessionStore = Arc<Mutex<HashMap<String, Value>>>;

/// Session 对象
#[derive(Clone)]
pub struct Session {
    /// Session ID
    session_id: String,
    /// 超时时间（分钟）
    timeout: u32,
    /// 数据存储
    data: SessionStore,
}

impl Session {
    /// 创建新 Session
    pub fn new(session_id: String) -> Self {
        Session {
            session_id,
            timeout: 20,
            data: Arc::new(Mutex::new(HashMap::new())),
        }
    }

    /// 获取 Session ID
    pub fn session_id(&self) -> &str {
        &self.session_id
    }

    /// 获取超时时间
    pub fn timeout(&self) -> u32 {
        self.timeout
    }

    /// 设置超时时间
    pub fn set_timeout(&mut self, timeout: u32) {
        self.timeout = timeout;
    }

    /// 放弃 Session
    pub fn abandon(&mut self) {
        if let Ok(mut data) = self.data.lock() {
            data.clear();
        }
    }

    /// 获取值
    pub fn get(&self, key: &str) -> Option<Value> {
        if let Ok(data) = self.data.lock() {
            data.get(&key.to_lowercase()).cloned()
        } else {
            None
        }
    }

    /// 设置值
    pub fn set(&mut self, key: String, value: Value) {
        if let Ok(mut data) = self.data.lock() {
            data.insert(key.to_lowercase(), value);
        }
    }
}

impl crate::runtime::BuiltinObject for Session {
    fn get_property(&self, name: &str) -> Result<Value, RuntimeError> {
        match name.to_lowercase().as_str() {
            "sessionid" => Ok(Value::String(self.session_id.clone())),
            "timeout" => Ok(Value::Number(self.timeout as f64)),
            _ => {
                // 尝试获取 Session 值
                self.get(name)
                    .ok_or_else(|| RuntimeError::PropertyNotFound(name.to_string()))
            }
        }
    }

    fn set_property(&mut self, name: &str, value: Value) -> Result<(), RuntimeError> {
        match name.to_lowercase().as_str() {
            "timeout" => {
                self.timeout = value.to_number() as u32;
                Ok(())
            }
            _ => {
                self.set(name.to_string(), value);
                Ok(())
            }
        }
    }

    fn call_method(&mut self, name: &str, args: Vec<Value>) -> Result<Value, RuntimeError> {
        match name.to_lowercase().as_str() {
            "abandon" => {
                self.abandon();
                Ok(Value::Empty)
            }
            "contents" => {
                if args.is_empty() {
                    return Ok(Value::Empty);
                }
                let key = ValueConversion::to_string(&args[0]).to_lowercase();
                self.get(&key)
                    .ok_or_else(|| RuntimeError::PropertyNotFound(key))
            }
            _ => Err(RuntimeError::MethodNotFound(name.to_string())),
        }
    }
}
