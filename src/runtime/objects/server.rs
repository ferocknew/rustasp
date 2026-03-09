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

    /// MapPath 方法 - 将虚拟路径转换为物理路径
    ///
    /// # Examples
    /// ```
    /// // root_path = "/var/www"
    /// server.map_path("/images/a.png") // => "/var/www/images/a.png"
    /// server.map_path("test.asp")      // => "/var/www/test.asp"
    /// ```
    pub fn map_path(&self, path: &str) -> String {
        let path = path.trim();

        // 如果是绝对路径或空路径，直接返回根路径
        if path.is_empty() {
            return self.root_path.clone();
        }

        // 处理绝对路径（以 / 或 \ 开头）
        if path.starts_with('/') || path.starts_with('\\') {
            // 移除开头的斜杠后拼接
            let clean_path = path.trim_start_matches('/').trim_start_matches('\\');
            if clean_path.is_empty() {
                return self.root_path.clone();
            }
            format!("{}/{}", self.root_path, clean_path)
        } else {
            // 相对路径，直接拼接
            format!("{}/{}", self.root_path, path)
        }
    }

    /// URLEncode 方法 - URL 编码
    ///
    /// # Examples
    /// ```
    /// server.url_encode("hello world") // => "hello%20world"
    /// server.url_encode("a=b&c=d")     // => "a%3Db%26c%3Dd"
    /// ```
    pub fn url_encode(&self, s: &str) -> String {
        urlencoding::encode(s).to_string()
    }

    /// HTMLEncode 方法 - HTML 转义
    ///
    /// # Examples
    /// ```
    /// server.html_encode("<b>")    // => "&lt;b&gt;"
    /// server.html_encode("a&b")    // => "a&amp;b"
    /// ```
    pub fn html_encode(&self, s: &str) -> String {
        html_escape::encode_text(s).to_string()
    }

    /// 获取脚本超时时间
    pub fn script_timeout(&self) -> u32 {
        self.script_timeout
    }

    /// 设置脚本超时时间
    pub fn set_script_timeout(&mut self, timeout: u32) {
        self.script_timeout = timeout;
    }
}

impl crate::runtime::BuiltinObject for Server {
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
            "createobject" => {
                // Server.CreateObject - 仅支持白名单中的对象
                if args.is_empty() {
                    return Err(RuntimeError::Generic(
                        "Server.CreateObject 需要 ProgID 参数".to_string()
                    ));
                }
                let prog_id = crate::utils::normalize_identifier(&ValueConversion::to_string(&args[0]));

                // 检查白名单
                const WHITELIST: &[&str] = &[
                    "scripting.dictionary",
                    "scripting.filesystemobject",
                    "msxml2.xmlhttp",
                ];

                let is_whitelisted = WHITELIST.iter().any(|&allowed| {
                    prog_id.eq_ignore_ascii_case(allowed)
                });

                if !is_whitelisted {
                    return Err(RuntimeError::CreateObjectFailed(format!(
                        " '{}' 不在白名单中。出于安全考虑，只允许创建以下对象: {}",
                        ValueConversion::to_string(&args[0]),
                        WHITELIST.join(", ")
                    )));
                }

                // 创建对象
                use crate::runtime::objects::{Dictionary, FileSystemObject, XmlHttp};
                use std::sync::{Arc, Mutex};

                match prog_id.as_str() {
                    "dictionary" | "scripting.dictionary" => {
                        Ok(Value::Object(Arc::new(Mutex::new(Dictionary::new()))))
                    }
                    "filesystemobject" | "scripting.filesystemobject" => {
                        Ok(Value::Object(Arc::new(Mutex::new(FileSystemObject::new()))))
                    }
                    "xmlhttp" | "msxml2.xmlhttp" | "microsoft.xmlhttp" => {
                        Ok(Value::Object(Arc::new(Mutex::new(XmlHttp::new()))))
                    }
                    _ => Err(RuntimeError::CreateObjectFailed(format!(
                        "无法创建对象 '{}'",
                        ValueConversion::to_string(&args[0])
                    )))
                }
            }
            "execute" => {
                // Server.Execute - 暂不支持
                Err(RuntimeError::Generic(
                    "Server.Execute is not yet implemented.".to_string()
                ))
            }
            "transfer" => {
                // Server.Transfer - 暂不支持
                Err(RuntimeError::Generic(
                    "Server.Transfer is not yet implemented.".to_string()
                ))
            }
            "getlasterror" => {
                // Server.GetLastError - 暂不支持
                Err(RuntimeError::Generic(
                    "Server.GetLastError is not yet implemented.".to_string()
                ))
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
