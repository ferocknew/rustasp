//! Response 对象

use crate::runtime::{RuntimeError, Value, ValueConversion};
use std::collections::HashMap;

/// Response 对象
pub struct Response {
    /// 输出缓冲区
    buffer: String,
    /// 内容类型
    content_type: String,
    /// 状态码
    status: u16,
    /// 响应头
    headers: HashMap<String, String>,
    /// 是否缓冲
    is_buffering: bool,
}

impl Response {
    /// 创建新 Response
    pub fn new() -> Self {
        Response {
            buffer: String::new(),
            content_type: "text/html".to_string(),
            status: 200,
            headers: HashMap::new(),
            is_buffering: true,
        }
    }

    /// 获取输出
    pub fn get_output(&self) -> &str {
        &self.buffer
    }

    /// 获取内容类型
    pub fn get_content_type(&self) -> &str {
        &self.content_type
    }

    /// 获取状态码
    pub fn get_status(&self) -> u16 {
        self.status
    }

    /// 获取响应头
    pub fn get_headers(&self) -> &HashMap<String, String> {
        &self.headers
    }

    /// 清空缓冲区
    pub fn clear(&mut self) {
        self.buffer.clear();
    }

    /// 结束响应
    pub fn end(&mut self) {
        self.is_buffering = false;
    }

    /// Write 方法
    pub fn write(&mut self, text: &str) {
        self.buffer.push_str(text);
    }

    /// WriteLn 方法
    pub fn write_ln(&mut self, text: &str) {
        self.buffer.push_str(text);
        self.buffer.push('\n');
    }

    /// Redirect 方法
    pub fn redirect(&mut self, url: &str) {
        self.status = 302;
        self.headers.insert("Location".to_string(), url.to_string());
    }

    /// AddHeader 方法
    pub fn add_header(&mut self, name: &str, value: &str) {
        self.headers.insert(name.to_string(), value.to_string());
    }
}

impl crate::runtime::BuiltinObject for Response {
    fn get_property(&self, name: &str) -> Result<Value, RuntimeError> {
        match name.to_lowercase().as_str() {
            "buffer" => Ok(Value::Boolean(self.is_buffering)),
            "contenttype" => Ok(Value::String(self.content_type.clone())),
            "status" => Ok(Value::Number(self.status as f64)),
            _ => Err(RuntimeError::PropertyNotFound(name.to_string())),
        }
    }

    fn set_property(&mut self, name: &str, value: Value) -> Result<(), RuntimeError> {
        match name.to_lowercase().as_str() {
            "buffer" => {
                self.is_buffering = value.to_bool();
                Ok(())
            }
            "contenttype" => {
                self.content_type = value.to_string();
                Ok(())
            }
            "status" => {
                self.status = value.to_number() as u16;
                Ok(())
            }
            _ => Err(RuntimeError::PropertyNotFound(name.to_string())),
        }
    }

    fn call_method(&mut self, name: &str, args: Vec<Value>) -> Result<Value, RuntimeError> {
        match name.to_lowercase().as_str() {
            "write" => {
                if args.len() != 1 {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                self.write(&args[0].to_string());
                Ok(Value::Empty)
            }
            "writeln" => {
                if args.len() != 1 {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                self.write_ln(&args[0].to_string());
                Ok(Value::Empty)
            }
            "redirect" => {
                if args.len() != 1 {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                self.redirect(&args[0].to_string());
                Ok(Value::Empty)
            }
            "clear" => {
                self.clear();
                Ok(Value::Empty)
            }
            "end" => {
                self.end();
                Ok(Value::Empty)
            }
            "addheader" => {
                if args.len() != 2 {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                self.add_header(&args[0].to_string(), &args[1].to_string());
                Ok(Value::Empty)
            }
            _ => Err(RuntimeError::MethodNotFound(name.to_string())),
        }
    }
}

impl Default for Response {
    fn default() -> Self {
        Self::new()
    }
}
