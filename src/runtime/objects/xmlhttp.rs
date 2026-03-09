//! MSXML2.XMLHTTP 对象 - HTTP 请求

use crate::runtime::{BuiltinObject, RuntimeError, Value, ValueConversion};

/// XMLHTTP 对象
#[derive(Debug, Clone)]
pub struct XmlHttp {
    /// 请求方法
    method: String,
    /// 请求 URL
    url: String,
    /// 是否异步
    is_async: bool,
    /// 请求头
    headers: Vec<(String, String)>,
    /// 响应状态
    status: u16,
    /// 响应状态文本
    status_text: String,
    /// 响应文本
    response_text: String,
    /// 响应体
    response_body: Vec<u8>,
    /// 是否已发送
    sent: bool,
}

impl XmlHttp {
    /// 创建新 XMLHTTP
    pub fn new() -> Self {
        XmlHttp {
            method: String::new(),
            url: String::new(),
            is_async: false,
            headers: Vec::new(),
            status: 0,
            status_text: String::new(),
            response_text: String::new(),
            response_body: Vec::new(),
            sent: false,
        }
    }

    /// 打开连接
    pub fn open(&mut self, method: &str, url: &str, is_async: Option<bool>) -> Result<(), RuntimeError> {
        self.method = method.to_uppercase();
        self.url = url.to_string();
        self.is_async = is_async.unwrap_or(false);
        self.sent = false;
        Ok(())
    }

    /// 设置请求头
    pub fn set_request_header(&mut self, name: &str, value: &str) -> Result<(), RuntimeError> {
        self.headers.push((name.to_string(), value.to_string()));
        Ok(())
    }

    /// 发送请求
    pub fn send(&mut self, _body: Option<&str>) -> Result<(), RuntimeError> {
        // 简化实现：只支持基本的功能演示
        // 实际实现需要使用 HTTP 客户端库

        // 模拟响应
        self.status = 200;
        self.status_text = "OK".to_string();
        self.response_text = format!("{{\"status\": \"ok\", \"url\": \"{}\"}}", self.url);
        self.response_body = self.response_text.as_bytes().to_vec();
        self.sent = true;

        Ok(())
    }

    /// 中止请求
    pub fn abort(&mut self) {
        self.method.clear();
        self.url.clear();
        self.status = 0;
        self.status_text.clear();
        self.response_text.clear();
        self.response_body.clear();
        self.sent = false;
    }

    /// 获取响应头
    pub fn get_response_header(&self, _name: &str) -> Option<String> {
        // 简化实现
        None
    }

    /// 获取所有响应头
    pub fn get_all_response_headers(&self) -> String {
        // 简化实现
        String::new()
    }
}

impl Default for XmlHttp {
    fn default() -> Self {
        Self::new()
    }
}

impl BuiltinObject for XmlHttp {
    fn as_any(&self) -> &dyn std::any::Any {
        self
    }

    fn as_any_mut(&mut self) -> &mut dyn std::any::Any {
        self
    }

    fn get_property(&self, name: &str) -> Result<Value, RuntimeError> {
        match name.to_lowercase().as_str() {
            "readystate" => Ok(Value::Number(if self.sent { 4.0 } else { 0.0 })),
            "status" => Ok(Value::Number(self.status as f64)),
            "statustext" => Ok(Value::String(self.status_text.clone())),
            "responsetext" => Ok(Value::String(self.response_text.clone())),
            "responsebody" => Ok(Value::String(format!("{:?}", self.response_body))),
            "responsexml" => Ok(Value::Empty), // 暂不支持 XML 解析
            _ => Err(RuntimeError::PropertyNotFound(name.to_string())),
        }
    }

    fn set_property(&mut self, _name: &str, _value: Value) -> Result<(), RuntimeError> {
        Err(RuntimeError::PropertyNotFound(_name.to_string()))
    }

    fn call_method(&mut self, name: &str, args: Vec<Value>) -> Result<Value, RuntimeError> {
        match name.to_lowercase().as_str() {
            "open" => {
                if args.len() < 2 {
                    return Err(RuntimeError::Generic("Open 方法需要至少 2 个参数".to_string()));
                }
                let method = ValueConversion::to_string(&args[0]);
                let url = ValueConversion::to_string(&args[1]);
                let is_async_flag = if args.len() > 2 {
                    Some(args[2].to_bool())
                } else {
                    None
                };
                self.open(&method, &url, is_async_flag)?;
                Ok(Value::Empty)
            }
            "setrequestheader" => {
                if args.len() < 2 {
                    return Err(RuntimeError::Generic("SetRequestHeader 方法需要 2 个参数".to_string()));
                }
                let name = ValueConversion::to_string(&args[0]);
                let value = ValueConversion::to_string(&args[1]);
                self.set_request_header(&name, &value)?;
                Ok(Value::Empty)
            }
            "send" => {
                let body = if !args.is_empty() {
                    Some(ValueConversion::to_string(&args[0]))
                } else {
                    None
                };
                self.send(body.as_deref())?;
                Ok(Value::Empty)
            }
            "abort" => {
                self.abort();
                Ok(Value::Empty)
            }
            "getresponseheader" => {
                if args.is_empty() {
                    return Ok(Value::Empty);
                }
                let name = ValueConversion::to_string(&args[0]);
                Ok(self.get_response_header(&name).map(Value::String).unwrap_or(Value::Empty))
            }
            "getallresponseheaders" => {
                Ok(Value::String(self.get_all_response_headers()))
            }
            _ => Err(RuntimeError::MethodNotFound(name.to_string())),
        }
    }
}
