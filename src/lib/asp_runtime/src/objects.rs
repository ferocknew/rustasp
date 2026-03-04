//! ASP 内置对象实现

use std::collections::HashMap;
use std::sync::{Arc, Mutex};
use vbscript_core::{BuiltinObject, RuntimeError, Value};

/// Response 对象
pub struct Response {
    buffer: String,
    content_type: String,
    status: u16,
    headers: HashMap<String, String>,
    cookies: HashMap<String, String>,
    is_buffering: bool,
}

impl Response {
    pub fn new() -> Self {
        Response {
            buffer: String::new(),
            content_type: "text/html".to_string(),
            status: 200,
            headers: HashMap::new(),
            cookies: HashMap::new(),
            is_buffering: true,
        }
    }

    /// 获取响应内容
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

    /// 获取 headers
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
}

impl BuiltinObject for Response {
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
                if let Value::Boolean(b) = value {
                    self.is_buffering = b;
                    Ok(())
                } else {
                    Err(RuntimeError::TypeMismatch("buffer".to_string()))
                }
            }
            "contenttype" => {
                if let Value::String(s) = value {
                    self.content_type = s;
                    Ok(())
                } else {
                    Err(RuntimeError::TypeMismatch("contentType".to_string()))
                }
            }
            "status" => {
                if let Value::Number(n) = value {
                    self.status = n as u16;
                    Ok(())
                } else {
                    Err(RuntimeError::TypeMismatch("status".to_string()))
                }
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
                let text = match &args[0] {
                    Value::String(s) => s.clone(),
                    other => other.to_string(),
                };
                self.buffer.push_str(&text);
                Ok(Value::Empty)
            }
            "writeLn" | "writeln" => {
                if args.len() != 1 {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                let text = match &args[0] {
                    Value::String(s) => s.clone(),
                    other => other.to_string(),
                };
                self.buffer.push_str(&text);
                self.buffer.push('\n');
                Ok(Value::Empty)
            }
            "redirect" => {
                if args.len() != 1 {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                if let Value::String(url) = &args[0] {
                    self.status = 302;
                    self.headers.insert("Location".to_string(), url.clone());
                    Ok(Value::Empty)
                } else {
                    Err(RuntimeError::TypeMismatch("redirect".to_string()))
                }
            }
            "clear" => {
                self.clear();
                Ok(Value::Empty)
            }
            "end" => {
                self.end();
                Ok(Value::Empty)
            }
            "addHeader" | "addheader" => {
                if args.len() != 2 {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                let key = match &args[0] {
                    Value::String(s) => s.clone(),
                    other => other.to_string(),
                };
                let value = match &args[1] {
                    Value::String(s) => s.clone(),
                    other => other.to_string(),
                };
                self.headers.insert(key, value);
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

/// Request 对象
pub struct Request {
    query_string: HashMap<String, String>,
    form: HashMap<String, String>,
    cookies: HashMap<String, String>,
    server_variables: HashMap<String, String>,
}

impl Request {
    pub fn new() -> Self {
        Request {
            query_string: HashMap::new(),
            form: HashMap::new(),
            cookies: HashMap::new(),
            server_variables: HashMap::new(),
        }
    }

    /// 设置查询字符串参数
    pub fn set_query_string(&mut self, key: String, value: String) {
        self.query_string.insert(key, value);
    }

    /// 设置表单数据
    pub fn set_form(&mut self, key: String, value: String) {
        self.form.insert(key, value);
    }

    /// 设置 Cookie
    pub fn set_cookie(&mut self, key: String, value: String) {
        self.cookies.insert(key, value);
    }

    /// 设置服务器变量
    pub fn set_server_variable(&mut self, key: String, value: String) {
        self.server_variables.insert(key, value);
    }
}

impl BuiltinObject for Request {
    fn get_property(&self, name: &str) -> Result<Value, RuntimeError> {
        match name.to_lowercase().as_str() {
            "querystring" => {
                // 返回一个对象，支持 Request.QueryString("key")
                Ok(Value::Object(HashMap::new()))
            }
            "form" => Ok(Value::Object(HashMap::new())),
            "cookies" => Ok(Value::Object(HashMap::new())),
            "servervariables" => Ok(Value::Object(HashMap::new())),
            _ => Err(RuntimeError::PropertyNotFound(name.to_string())),
        }
    }

    fn set_property(&mut self, _name: &str, _value: Value) -> Result<(), RuntimeError> {
        Err(RuntimeError::Generic("Request object is read-only".to_string()))
    }

    fn call_method(&mut self, name: &str, args: Vec<Value>) -> Result<Value, RuntimeError> {
        match name.to_lowercase().as_str() {
            "querystring" | "form" | "cookies" | "servervariables" => {
                if args.is_empty() {
                    return Ok(Value::Empty);
                }
                let key = match &args[0] {
                    Value::String(s) => s.to_lowercase(),
                    other => other.to_string().to_lowercase(),
                };
                let map = match name.to_lowercase().as_str() {
                    "querystring" => &self.query_string,
                    "form" => &self.form,
                    "cookies" => &self.cookies,
                    "servervariables" => &self.server_variables,
                    _ => return Err(RuntimeError::MethodNotFound(name.to_string())),
                };
                Ok(map.get(&key)
                    .map(|s| Value::String(s.clone()))
                    .unwrap_or(Value::Empty))
            }
            _ => Err(RuntimeError::MethodNotFound(name.to_string())),
        }
    }
}

impl Default for Request {
    fn default() -> Self {
        Self::new()
    }
}

/// Server 对象
pub struct Server {
    script_timeout: u32,
}

impl Server {
    pub fn new() -> Self {
        Server { script_timeout: 90 }
    }
}

impl BuiltinObject for Server {
    fn get_property(&self, name: &str) -> Result<Value, RuntimeError> {
        match name.to_lowercase().as_str() {
            "scripttimeout" => Ok(Value::Number(self.script_timeout as f64)),
            _ => Err(RuntimeError::PropertyNotFound(name.to_string())),
        }
    }

    fn set_property(&mut self, name: &str, value: Value) -> Result<(), RuntimeError> {
        match name.to_lowercase().as_str() {
            "scripttimeout" => {
                if let Value::Number(n) = value {
                    self.script_timeout = n as u32;
                    Ok(())
                } else {
                    Err(RuntimeError::TypeMismatch("scriptTimeout".to_string()))
                }
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
                let path = match &args[0] {
                    Value::String(s) => s.clone(),
                    other => other.to_string(),
                };
                // 简化实现：直接返回路径
                Ok(Value::String(path))
            }
            "urlencode" | "urlencode" => {
                if args.len() != 1 {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                let s = match &args[0] {
                    Value::String(s) => s.clone(),
                    other => other.to_string(),
                };
                Ok(Value::String(urlencoding::encode(&s).to_string()))
            }
            "htmlencode" | "htmlencode" => {
                if args.len() != 1 {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                let s = match &args[0] {
                    Value::String(s) => s.clone(),
                    other => other.to_string(),
                };
                Ok(Value::String(html_escape::encode_text(&s).to_string()))
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

/// Session 对象（使用 Arc<Mutex> 支持多线程）
pub type Session = Arc<Mutex<SessionInner>>;

/// Session 内部实现
pub struct SessionInner {
    pub session_id: String,
    pub timeout: u32,
    pub data: HashMap<String, Value>,
}

impl SessionInner {
    pub fn new(session_id: String) -> Self {
        SessionInner {
            session_id,
            timeout: 20,
            data: HashMap::new(),
        }
    }
}

/// Application 对象（使用 Arc<Mutex> 支持多线程）
pub type Application = Arc<Mutex<ApplicationInner>>;

/// Application 内部实现
pub struct ApplicationInner {
    pub data: HashMap<String, Value>,
}

impl ApplicationInner {
    pub fn new() -> Self {
        ApplicationInner {
            data: HashMap::new(),
        }
    }
}

impl Default for ApplicationInner {
    fn default() -> Self {
        Self::new()
    }
}
