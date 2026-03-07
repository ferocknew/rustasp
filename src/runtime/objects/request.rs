//! Request 对象

use crate::runtime::{RuntimeError, Value, ValueConversion};
use std::collections::HashMap;

/// Request 对象
#[derive(Debug, Clone)]
pub struct Request {
    /// 查询字符串参数（支持多值）
    query_string: HashMap<String, Vec<String>>,
    /// 表单数据（支持多值）
    form: HashMap<String, Vec<String>>,
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

    /// 设置查询字符串参数（单个值）
    pub fn set_query_string(&mut self, key: String, value: String) {
        self.query_string.insert(key.to_lowercase(), vec![value]);
    }

    /// 设置表单数据（单个值）
    pub fn set_form(&mut self, key: String, value: String) {
        self.form.insert(key.to_lowercase(), vec![value]);
    }

    /// 设置 Cookie
    pub fn set_cookie(&mut self, key: String, value: String) {
        self.cookies.insert(key.to_lowercase(), value);
    }

    /// 设置服务器变量
    pub fn set_server_variable(&mut self, key: String, value: String) {
        self.server_variables.insert(key.to_lowercase(), value);
    }

    /// 设置查询字符串参数（多个值）
    pub fn set_query_string_multiple(&mut self, key: String, values: Vec<String>) {
        self.query_string.insert(key.to_lowercase(), values);
    }

    /// 设置表单数据（多个值）
    pub fn set_form_multiple(&mut self, key: String, values: Vec<String>) {
        self.form.insert(key.to_lowercase(), values);
    }

    /// 获取 QueryString 第一个值
    pub fn query_string(&self, key: &str) -> Option<&String> {
        self.query_string.get(&key.to_lowercase()).and_then(|v| v.first())
    }

    /// 获取 Form 第一个值
    pub fn form(&self, key: &str) -> Option<&String> {
        self.form.get(&key.to_lowercase()).and_then(|v| v.first())
    }

    /// 获取 QueryString 所有值
    pub fn query_string_all(&self, key: &str) -> Option<&Vec<String>> {
        self.query_string.get(&key.to_lowercase())
    }

    /// 获取 Form 所有值
    pub fn form_all(&self, key: &str) -> Option<&Vec<String>> {
        self.form.get(&key.to_lowercase())
    }

    /// 获取所有值（先 QueryString 后 Form）
    pub fn get_all(&self, key: &str) -> Option<&Vec<String>> {
        let key_lower = key.to_lowercase();
        if let Some(values) = self.query_string.get(&key_lower) {
            return Some(values);
        }
        self.form.get(&key_lower)
    }
}

impl crate::runtime::BuiltinObject for Request {
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
            "querystring" => Ok(Value::new_dictionary()),
            "form" => Ok(Value::new_dictionary()),
            "cookies" => Ok(Value::new_dictionary()),
            "servervariables" => Ok(Value::new_dictionary()),
            "clientcertificate" => Ok(Value::new_dictionary()),
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
                // BinaryReader 方法：读取请求体的二进制数据
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
                match self.query_string.get(&key) {
                    Some(values) if values.len() == 1 => Ok(Value::String(values[0].clone())),
                    Some(values) => {
                        let arr: Vec<Value> = values.iter().map(|s| Value::String(s.clone())).collect();
                        Ok(Value::Array(arr))
                    }
                    None => Ok(Value::Empty),
                }
            }
            "form" => {
                if args.is_empty() {
                    return Ok(Value::Empty);
                }
                let key = ValueConversion::to_string(&args[0]).to_lowercase();
                match self.form.get(&key) {
                    Some(values) if values.len() == 1 => Ok(Value::String(values[0].clone())),
                    Some(values) => {
                        let arr: Vec<Value> = values.iter().map(|s| Value::String(s.clone())).collect();
                        Ok(Value::Array(arr))
                    }
                    None => Ok(Value::Empty),
                }
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

    fn index(&self, key: &Value) -> Result<Value, RuntimeError> {
        let key_str = ValueConversion::to_string(key).to_lowercase();
        // 优先查询 querystring，然后是 form
        if let Some(values) = self.query_string.get(&key_str) {
            if values.len() == 1 {
                return Ok(Value::String(values[0].clone()));
            }
            let arr: Vec<Value> = values.iter().map(|s| Value::String(s.clone())).collect();
            return Ok(Value::Array(arr));
        }
        if let Some(values) = self.form.get(&key_str) {
            if values.len() == 1 {
                return Ok(Value::String(values[0].clone()));
            }
            let arr: Vec<Value> = values.iter().map(|s| Value::String(s.clone())).collect();
            return Ok(Value::Array(arr));
        }
        Ok(Value::Empty)
    }
}

impl Default for Request {
    fn default() -> Self {
        Self::new()
    }
}
