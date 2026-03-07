//! Request 对象

use crate::runtime::{RuntimeError, Value, ValueConversion};
use std::collections::HashMap;

/// Request 对象
pub struct Request {
    /// 查询字符串参数
    query_string: HashMap<String, String>,
    /// 表单数据
    form: HashMap<String, String>,
    /// Cookies
    cookies: HashMap<String, String>,
    /// 服务器变量
    server_variables: HashMap<String, String>,
    /// 请求正文（用于 BinaryRead）
    raw_body: Vec<u8>,
    /// 客户端证书信息
    client_certificate: HashMap<String, String>,
}

impl Request {
    /// 创建新 Request
    pub fn new() -> Self {
        Request {
            query_string: HashMap::new(),
            form: HashMap::new(),
            cookies: HashMap::new(),
            server_variables: HashMap::new(),
            raw_body: Vec::new(),
            client_certificate: HashMap::new(),
        }
    }

    /// 设置原始请求体
    pub fn set_raw_body(&mut self, body: Vec<u8>) {
        self.raw_body = body;
    }

    /// 获取 TotalBytes（请求体字节总数）
    pub fn total_bytes(&self) -> i32 {
        self.raw_body.len() as i32
    }

    /// BinaryRead：读取请求体的二进制数据
    pub fn binary_read(&self, bytes: i32) -> Result<Vec<u8>, RuntimeError> {
        let bytes_to_read = bytes.max(0) as usize;
        if bytes_to_read == 0 {
            return Ok(Vec::new());
        }

        let available = self.raw_body.len();
        let to_read = bytes_to_read.min(available);

        Ok(self.raw_body[..to_read].to_vec())
    }

    /// 设置客户端证书字段
    pub fn set_client_certificate(&mut self, key: String, value: String) {
        self.client_certificate.insert(key.to_lowercase(), value);
    }

    /// 获取客户端证书信息
    pub fn get_client_certificate(&self, field: &str) -> Option<&String> {
        self.client_certificate.get(&field.to_lowercase())
    }

    /// 设置查询字符串参数
    pub fn set_query_string(&mut self, key: String, value: String) {
        self.query_string.insert(key.to_lowercase(), value);
    }

    /// 设置表单数据
    pub fn set_form(&mut self, key: String, value: String) {
        self.form.insert(key.to_lowercase(), value);
    }

    /// 设置 Cookie
    pub fn set_cookie(&mut self, key: String, value: String) {
        self.cookies.insert(key.to_lowercase(), value);
    }

    /// 设置服务器变量
    pub fn set_server_variable(&mut self, key: String, value: String) {
        self.server_variables.insert(key.to_lowercase(), value);
    }

    /// 获取 QueryString
    pub fn query_string(&self, key: &str) -> Option<&String> {
        self.query_string.get(&key.to_lowercase())
    }

    /// 获取 Form
    pub fn form(&self, key: &str) -> Option<&String> {
        self.form.get(&key.to_lowercase())
    }
}

impl crate::runtime::BuiltinObject for Request {
    fn get_property(&self, name: &str) -> Result<Value, RuntimeError> {
        match name.to_lowercase().as_str() {
            "querystring" => Ok(Value::Object(HashMap::new())),
            "form" => Ok(Value::Object(HashMap::new())),
            "cookies" => Ok(Value::Object(HashMap::new())),
            "servervariables" => Ok(Value::Object(HashMap::new())),
            "clientcertificate" => Ok(Value::Object(HashMap::new())),
            "totalbytes" => Ok(Value::Number(self.total_bytes() as f64)),
            _ => Err(RuntimeError::PropertyNotFound(name.to_string())),
        }
    }

    fn set_property(&mut self, _name: &str, _value: Value) -> Result<(), RuntimeError> {
        Err(RuntimeError::Generic(
            "Request object is read-only".to_string(),
        ))
    }

    fn call_method(&mut self, name: &str, args: Vec<Value>) -> Result<Value, RuntimeError> {
        match name.to_lowercase().as_str() {
            "binaryread" => {
                // BinaryRead 方法：读取请求体的二进制数据
                let bytes_to_read = if args.is_empty() {
                    0
                } else {
                    ValueConversion::to_number(&args[0]) as i32
                };
                match self.binary_read(bytes_to_read) {
                    Ok(data) => {
                        // 将字节数组转换为 Value::Array
                        let array_values: Vec<Value> = data
                            .into_iter()
                            .map(|b| Value::Number(b as f64))
                            .collect();
                        Ok(Value::Array(array_values))
                    }
                    Err(e) => Err(e),
                }
            }
            "querystring" => {
                if args.is_empty() {
                    return Ok(Value::Empty);
                }
                let key = ValueConversion::to_string(&args[0]).to_lowercase();
                Ok(self
                    .query_string
                    .get(&key)
                    .map(|s| Value::String(s.clone()))
                    .unwrap_or(Value::Empty))
            }
            "form" => {
                if args.is_empty() {
                    return Ok(Value::Empty);
                }
                let key = ValueConversion::to_string(&args[0]).to_lowercase();
                Ok(self
                    .form
                    .get(&key)
                    .map(|s| Value::String(s.clone()))
                    .unwrap_or(Value::Empty))
            }
            "cookies" => {
                if args.is_empty() {
                    return Ok(Value::Empty);
                }
                let key = ValueConversion::to_string(&args[0]).to_lowercase();
                Ok(self
                    .cookies
                    .get(&key)
                    .map(|s| Value::String(s.clone()))
                    .unwrap_or(Value::Empty))
            }
            "servervariables" => {
                if args.is_empty() {
                    return Ok(Value::Empty);
                }
                let key = ValueConversion::to_string(&args[0]).to_lowercase();
                Ok(self
                    .server_variables
                    .get(&key)
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
