//! Server 对象

use crate::runtime::{RuntimeError, Value, ValueConversion};

/// Server 对象
#[derive(Debug, Clone)]
pub struct Server {
    /// 脚本超时时间（秒）
    script_timeout: u32,
    /// 根路径
    root_path: String,
}

impl Server {
    /// 创建新 Server
    pub fn new() -> Self {
        Server {
            script_timeout: 90,
            root_path: ".".to_string(),
        }
    }

    /// 设置根路径
    pub fn set_root_path(&mut self, path: String) {
        self.root_path = path;
    }

    /// MapPath 方法
    pub fn map_path(&self, path: &str) -> String {
        if path.starts_with('/') || path.starts_with('\\') {
            format!("{}{}", self.root_path, path)
        } else {
            format!("{}/{}", self.root_path, path)
        }
    }

    /// URLEncode 方法
    pub fn url_encode(&self, s: &str) -> String {
        urlencoding::encode(s).to_string()
    }

    /// HTMLEncode 方法
    pub fn html_encode(&self, s: &str) -> String {
        html_escape::encode_text(s).to_string()
    }
}

impl crate::runtime::BuiltinObject for Server {
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
            "scripttimeout" => Ok(Value::Number(self.script_timeout as f64)),
            _ => Err(RuntimeError::PropertyNotFound(name.to_string())),
        }
    }

    fn set_property(&mut self, name: &str, value: Value) -> Result<(), RuntimeError> {
        match name.to_lowercase().as_str() {
            "scripttimeout" => {
                self.script_timeout = value.to_number() as u32;
                Ok(())
            }
            _ => Err(RuntimeError::PropertyNotFound(name.to_string())),
        }
    }

    fn call_method(&mut self, name: &str, args: Vec<Value>) -> Result<Value, RuntimeError> {
        match name.to_lowercase().as_str() {
            "mappath" => {
                if args.len() != 1 {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                let path = ValueConversion::to_string(&args[0]);
                Ok(Value::String(self.map_path(&path)))
            }
            "urlencode" => {
                if args.len() != 1 {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                let s = ValueConversion::to_string(&args[0]);
                Ok(Value::String(self.url_encode(&s)))
            }
            "htmlencode" => {
                if args.len() != 1 {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                let s = ValueConversion::to_string(&args[0]);
                Ok(Value::String(self.html_encode(&s)))
            }
            _ => Err(RuntimeError::MethodNotFound(name.to_string())),
        }
    }
}

impl Default for Server {
    fn default() -> Self {
        Self::new()
    }
}
