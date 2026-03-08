//! Response 对象

use crate::runtime::{RuntimeError, Value, ValueConversion};
use std::collections::HashMap;
use std::time::{SystemTime, UNIX_EPOCH};

/// Response 对象
#[derive(Debug, Clone)]
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
    /// 字符集
    charset: Option<String>,
    /// 代码页
    codepage: Option<i32>,
    /// 缓存控制
    cache_control: Option<String>,
    /// 过期时间（分钟）
    expires: Option<i32>,
    /// 过期绝对时间
    expires_absolute: Option<SystemTime>,
    /// PICS 标签
    pics: Option<String>,
    /// Cookies 集合
    cookies: HashMap<String, String>,
    /// 日志追加内容
    append_to_log: Vec<String>,
    /// 是否已结束（Response.End 调用后设置为 true）
    is_ended: bool,
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
            charset: None,
            codepage: None,
            cache_control: None,
            expires: None,
            expires_absolute: None,
            pics: None,
            cookies: HashMap::new(),
            append_to_log: Vec::new(),
            is_ended: false,
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
        self.is_ended = true;
    }

    /// 检查是否已结束
    pub fn is_ended(&self) -> bool {
        self.is_ended
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

    /// AppendToLog 方法
    pub fn append_to_log(&mut self, text: &str) {
        self.append_to_log.push(text.to_string());
    }

    /// BinaryWrite 方法
    pub fn binary_write(&mut self, data: &[u8]) {
        // 对于文本输出，转换为字符串
        // 在实际应用中可能需要特殊处理二进制数据
        if let Ok(text) = std::str::from_utf8(data) {
            self.buffer.push_str(text);
        }
    }

    /// Flush 方法
    pub fn flush(&mut self) {
        // 立即发送已缓冲的输出
        // 在实际实现中，这里可能需要触发实际的网络发送
        self.is_buffering = false;
    }

    /// 设置 Cookie
    pub fn set_cookie(&mut self, name: &str, value: &str) {
        self.cookies.insert(name.to_string(), value.to_string());
    }

    /// 获取 Cookie
    pub fn get_cookie(&self, name: &str) -> Option<&String> {
        self.cookies.get(name)
    }

    /// 设置 CacheControl
    pub fn set_cache_control(&mut self, value: &str) {
        self.cache_control = Some(value.to_string());
    }

    /// 获取 CacheControl
    pub fn get_cache_control(&self) -> Option<&String> {
        self.cache_control.as_ref()
    }

    /// 设置 Charset
    pub fn set_charset(&mut self, value: &str) {
        self.charset = Some(value.to_string());
    }

    /// 获取 Charset
    pub fn get_charset(&self) -> Option<&String> {
        self.charset.as_ref()
    }

    /// 设置 Expires（分钟）
    pub fn set_expires(&mut self, minutes: i32) {
        self.expires = Some(minutes);
        // 同时计算 expires_absolute
        let duration = std::time::Duration::from_secs(minutes.max(0) as u64 * 60);
        self.expires_absolute = SystemTime::now().checked_add(duration);
    }

    /// 获取 Expires
    pub fn get_expires(&self) -> Option<i32> {
        self.expires
    }

    ///获取 ExpiresAbsolute（Unix 时间戳秒数）
    pub fn get_expires_absolute(&self) -> Option<f64> {
        self.expires_absolute.map(|t| {
            t.duration_since(UNIX_EPOCH)
                .unwrap_or_default()
                .as_secs_f64()
        })
    }

    /// 检查客户端是否已断开连接
    pub fn is_client_connected(&self) -> bool {
        // 简化实现：假设客户端始终连接
        // 实际应用中可以检测 TCP 连接状态
        true
    }

    /// 设置 PICS 标签
    pub fn set_pics(&mut self, value: &str) {
        self.pics = Some(value.to_string());
    }

    /// 获取 PICS 标签
    pub fn get_pics(&self) -> Option<&String> {
        self.pics.as_ref()
    }
}

impl crate::runtime::BuiltinObject for Response {
    fn as_any(&self) -> &dyn std::any::Any {
        self
    }

    fn as_any_mut(&mut self) -> &mut dyn std::any::Any {
        self
    }

    fn get_property(&self, name: &str) -> Result<Value, RuntimeError> {
        match name.to_lowercase().as_str() {
            // 基本属性
            "buffer" => Ok(Value::Boolean(self.is_buffering)),
            "contenttype" => Ok(Value::String(self.content_type.clone())),
            "status" => Ok(Value::Number(self.status as f64)),
            // 新增属性
            "charset" => Ok(self.get_charset().map(|s| Value::String(s.clone())).unwrap_or(Value::Empty)),
            "codepage" => Ok(self.codepage.map(|n| Value::Number(n as f64)).unwrap_or(Value::Empty)),
            "cachecontrol" => Ok(self.get_cache_control().map(|s| Value::String(s.clone())).unwrap_or(Value::Empty)),
            "expires" => Ok(self.get_expires().map(|n| Value::Number(n as f64)).unwrap_or(Value::Empty)),
            "expiresabsolute" => Ok(self.get_expires_absolute().map(|n| Value::Number(n)).unwrap_or(Value::Empty)),
            "pics" => Ok(self.get_pics().map(|s| Value::String(s.clone())).unwrap_or(Value::Empty)),
            "isclientconnected" => Ok(Value::Boolean(self.is_client_connected())),
            // Cookies 集合返回一个标记对象
            "cookies" => {
                let mut cookies_obj = HashMap::new();
                cookies_obj.insert("__response_cookies__".to_string(), Value::Boolean(true));
                Ok(Value::from_hashmap(cookies_obj))
            }
            _ => Err(RuntimeError::PropertyNotFound(name.to_string())),
        }
    }

    fn set_property(&mut self, name: &str, value: Value) -> Result<(), RuntimeError> {
        match name.to_lowercase().as_str() {
            // 基本属性
            "buffer" => {
                self.is_buffering = ValueConversion::to_bool(&value);
                Ok(())
            }
            "contenttype" => {
                self.content_type = ValueConversion::to_string(&value);
                Ok(())
            }
            "status" => {
                let status_str = ValueConversion::to_string(&value);
                // 解析状态码，支持 "404 Not Found" 或纯数字 "404" 格式
                if let Some(code_str) = status_str.split_whitespace().next() {
                    if let Ok(code) = code_str.parse::<u16>() {
                        self.status = code;
                    }
                }
                Ok(())
            }
            // 新增属性
            "charset" => {
                self.set_charset(&ValueConversion::to_string(&value));
                Ok(())
            }
            "codepage" => {
                // 保存代码页设置,但在当前实现中不做实际转换
                // 因为 Rust 的字符串处理默认就是 UTF-8
                self.codepage = Some(ValueConversion::to_number(&value) as i32);
                Ok(())
            }
            "cachecontrol" => {
                self.set_cache_control(&ValueConversion::to_string(&value));
                Ok(())
            }
            "expires" => {
                self.set_expires(ValueConversion::to_number(&value) as i32);
                Ok(())
            }
            "pics" => {
                self.set_pics(&ValueConversion::to_string(&value));
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
                self.write(&ValueConversion::to_string(&args[0]));
                Ok(Value::Empty)
            }
            "writeln" => {
                if args.len() != 1 {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                self.write_ln(&ValueConversion::to_string(&args[0]));
                Ok(Value::Empty)
            }
            "redirect" => {
                if args.len() != 1 {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                self.redirect(&ValueConversion::to_string(&args[0]));
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
                self.add_header(
                    &ValueConversion::to_string(&args[0]),
                    &ValueConversion::to_string(&args[1]),
                );
                Ok(Value::Empty)
            }
            // 新增方法
            "appendtolog" => {
                if args.len() != 1 {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                self.append_to_log(&ValueConversion::to_string(&args[0]));
                Ok(Value::Empty)
            }
            "binarywrite" => {
                if args.len() != 1 {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                // 尝试将参数作为字节数组处理
                match &args[0] {
                    Value::Array(ref arr) => {
                        let locked_arr = arr.lock().unwrap();
                        let bytes: Vec<u8> = locked_arr.data.iter()
                            .map(|v| ValueConversion::to_number(v) as u8)
                            .collect();
                        self.binary_write(&bytes);
                    }
                    Value::String(s) => {
                        self.binary_write(s.as_bytes());
                    }
                    _ => {
                        // 尝试转换
                        let text = ValueConversion::to_string(&args[0]);
                        self.binary_write(text.as_bytes());
                    }
                }
                Ok(Value::Empty)
            }
            "flush" => {
                self.flush();
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
