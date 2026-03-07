//! Session 对象

use crate::runtime::{RuntimeError, Value, ValueConversion};
use std::collections::HashMap;
use std::sync::{Arc, Mutex};

/// Session 数据（用于序列化/持久化）
#[derive(Debug, Clone, serde::Serialize, serde::Deserialize)]
pub struct SessionData {
    pub session_id: String,
    pub timeout: u32,
    pub created_at: u64,
    pub last_accessed: u64,
    pub data: HashMap<String, serde_json::Value>,
}

impl SessionData {
    pub fn new(session_id: String, timeout: u32) -> Self {
        let now = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .as_secs();

        SessionData {
            session_id,
            timeout,
            created_at: now,
            last_accessed: now,
            data: HashMap::new(),
        }
    }

    pub fn is_expired(&self, now: u64) -> bool {
        let timeout_seconds = self.timeout as u64 * 60;
        now > self.last_accessed + timeout_seconds
    }
}

/// Session 存储（内存版本）
pub type SessionStore = Arc<Mutex<HashMap<String, Value>>>;

/// Session 对象
#[derive(Debug, Clone)]
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

    /// 从 SessionData 创建 Session（用于恢复）
    pub fn from_session_data(session_data: SessionData) -> Self {
        // 将 serde_json::Value 转换回 Value
        let mut data = HashMap::new();
        for (key, json_value) in session_data.data {
            let value = match json_value {
                serde_json::Value::String(s) => Value::String(s),
                serde_json::Value::Number(n) => {
                    if n.is_i64() {
                        Value::Number(n.as_i64().unwrap_or(0) as f64)
                    } else if n.is_u64() {
                        Value::Number(n.as_u64().unwrap_or(0) as f64)
                    } else {
                        Value::Number(n.as_f64().unwrap_or(0.0))
                    }
                }
                serde_json::Value::Bool(b) => Value::Boolean(b),
                serde_json::Value::Null => Value::Null,
                _ => Value::Empty, // 不支持的类型
            };
            data.insert(key, value);
        }

        Session {
            session_id: session_data.session_id,
            timeout: session_data.timeout,
            data: Arc::new(Mutex::new(data)),
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

    /// 获取所有数据
    pub fn get_all_data(&self) -> HashMap<String, Value> {
        if let Ok(data) = self.data.lock() {
            data.clone()
        } else {
            HashMap::new()
        }
    }

    /// 转换为 SessionData（用于序列化）
    pub fn to_session_data(&self) -> Result<SessionData, String> {
        let now = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .as_secs();

        // 将 HashMap<String, Value> 转换为 HashMap<String, serde_json::Value>
        let mut data = HashMap::new();
        if let Ok(session_data) = self.data.lock() {
            for (key, value) in session_data.iter() {
                // 简化处理：只支持基本类型
                let json_value = match value {
                    Value::String(s) => serde_json::Value::String(s.clone()),
                    Value::Number(n) => {
                        // 检查是否是整数
                        if n.fract() == 0.0 && *n >= i64::MIN as f64 && *n <= i64::MAX as f64 {
                            serde_json::Value::Number(serde_json::Number::from(*n as i64))
                        } else {
                            // 浮点数作为字符串存储
                            serde_json::Value::String(n.to_string())
                        }
                    }
                    Value::Boolean(b) => serde_json::Value::Bool(*b),
                    Value::Empty => serde_json::Value::Null,
                    Value::Null => serde_json::Value::Null,
                    _ => serde_json::Value::Null, // 数组和对象暂不支持
                };
                data.insert(key.clone(), json_value);
            }
        }

        Ok(SessionData {
            session_id: self.session_id.clone(),
            timeout: self.timeout,
            created_at: now - 100, // 假设创建于 100 秒前
            last_accessed: now,
            data,
        })
    }
}

impl crate::runtime::BuiltinObject for Session {
    fn clone_box(&self) -> Box<dyn crate::runtime::BuiltinObject> {
        Box::new(self.clone())
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }

    fn as_any_mut(&mut self) -> &mut dyn std::any::Any {
        self
    }

    fn get_property(&self, name: &str) -> Result<Value, RuntimeError> {
        match name.to_lowercase().as_str() {
            "sessionid" => Ok(Value::String(self.session_id.clone())),
            "timeout" => Ok(Value::Number(self.timeout as f64)),
            "codepage" => Ok(Value::Empty), // CodePage 属性，返回 Empty
            // Note: Contents 属性的特殊处理在 eval_property 中实现
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
            "codepage" => {
                // CodePage 属性，保存但不做实际转换
                // 因为 Rust 的字符串处理默认就是 UTF-8
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
            "remove" => {
                if args.is_empty() {
                    return Ok(Value::Empty);
                }
                let key = ValueConversion::to_string(&args[0]).to_lowercase();
                if let Ok(mut data) = self.data.lock() {
                    data.remove(&key);
                }
                Ok(Value::Empty)
            }
            "removeall" => {
                if let Ok(mut data) = self.data.lock() {
                    // 只移除用户数据，保留 sessionid 和 timeout
                    let keys_to_remove: Vec<String> = data.keys()
                        .filter(|k| !k.starts_with("__") && *k != "sessionid" && *k != "timeout")
                        .cloned()
                        .collect();
                    for key in keys_to_remove {
                        data.remove(&key);
                    }
                }
                Ok(Value::Empty)
            }
            _ => Err(RuntimeError::MethodNotFound(name.to_string())),
        }
    }

    fn index(&self, key: &Value) -> Result<Value, RuntimeError> {
        let key_str = ValueConversion::to_string(key).to_lowercase();
        self.get(&key_str)
            .ok_or_else(|| RuntimeError::PropertyNotFound(key_str))
    }
}
