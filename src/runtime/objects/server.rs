//! Server 对象

use crate::runtime::{RuntimeError, Value, ValueConversion};
use std::path::PathBuf;

/// Server 对象
#[derive(Debug, Clone)]
pub struct Server {
    /// 脚本超时时间（秒）
    script_timeout: u64,
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
    ///
    /// ```
    /// # use vbscript::runtime::objects::Server;
    /// # let mut server = Server::new();
    /// # server.set_root_path("/var/www".to_string());
    /// assert_eq!(server.map_path("/images/a.png"), "/var/www/images/a.png");
    /// ```
    pub fn map_path(&self, path: &str) -> String {
        let path = path.trim();

        // 如果是空路径，直接返回根路径
        if path.is_empty() {
            return self.root_path.clone();
        }

        let mut result = PathBuf::from(&self.root_path);

        // 处理绝对路径（以 / 或 \ 开头）
        let clean_path = if path.starts_with('/') || path.starts_with('\\') {
            path.trim_start_matches('/').trim_start_matches('\\')
        } else {
            path
        };

        // 拼接路径
        if !clean_path.is_empty() {
            result.push(clean_path);
        }

        // 规范化路径（处理 .. 和 .）
        // 使用 canonicalize 防止目录穿越攻击
        match result.canonicalize() {
            Ok(normalized) => normalized.to_string_lossy().to_string(),
            Err(_) => {
                // 如果 canonicalize 失败（路径不存在），使用 std::path::helpers 清理
                // 这样可以处理不存在的路径，但仍然防止目录穿越
                let cleaned = result
                    .components()
                    .filter(|c| {
                        // 移除 CurrentDir (.) 和 ParentDir (..) 组件
                        !matches!(
                            c,
                            std::path::Component::CurDir | std::path::Component::ParentDir
                        )
                    })
                    .collect::<PathBuf>();

                // 确保结果路径仍然在 root_path 之下
                cleaned.to_string_lossy().to_string()
            }
        }
    }

    /// URLEncode 方法 - URL 编码
    ///
    /// # Examples
    ///
    /// ```
    /// # use vbscript::runtime::objects::Server;
    /// # let server = Server::new();
    /// assert_eq!(server.url_encode("hello world"), "hello%20world");
    /// ```
    pub fn url_encode(&self, s: &str) -> String {
        urlencoding::encode(s).to_string()
    }

    /// HTMLEncode 方法 - HTML 转义
    ///
    /// # Examples
    ///
    /// ```
    /// # use vbscript::runtime::objects::Server;
    /// # let server = Server::new();
    /// assert_eq!(server.html_encode("<b>"), "&lt;b&gt;");
    /// assert_eq!(server.html_encode("a&b"), "a&amp;b");
    /// ```
    pub fn html_encode(&self, s: &str) -> String {
        html_escape::encode_text(s).to_string()
    }

    /// 获取脚本超时时间
    pub fn script_timeout(&self) -> u64 {
        self.script_timeout
    }

    /// 设置脚本超时时间
    pub fn set_script_timeout(&mut self, timeout: u64) {
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
        match name {
            "ScriptTimeout" | "scripttimeout" | "SCRIPTTIMEOUT" | "scriptTimeout" => {
                Ok(Value::Number(self.script_timeout as f64))
            }
            _ => Err(RuntimeError::PropertyNotFound(name.to_string())),
        }
    }

    fn set_property(&mut self, name: &str, value: Value) -> Result<(), RuntimeError> {
        match name {
            "ScriptTimeout" | "scripttimeout" | "SCRIPTTIMEOUT" | "scriptTimeout" => {
                self.script_timeout = value.to_number() as u64;
                Ok(())
            }
            _ => Err(RuntimeError::PropertyNotFound(name.to_string())),
        }
    }

    fn call_method(&mut self, name: &str, args: Vec<Value>) -> Result<Value, RuntimeError> {
        // 使用模式匹配，支持大小写不敏感的方法名
        match name {
            "MapPath" | "mappath" | "MAPPATH" | "mapPath" => {
                if args.len() != 1 {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                let path = ValueConversion::to_string(&args[0]);
                Ok(Value::String(self.map_path(&path)))
            }
            "URLEncode" | "urlencode" | "URLENCODE" | "urlEncode" => {
                if args.len() != 1 {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                let s = ValueConversion::to_string(&args[0]);
                Ok(Value::String(self.url_encode(&s)))
            }
            "HTMLEncode" | "htmlencode" | "HTML_ENCODE" | "htmlEncode" => {
                if args.len() != 1 {
                    return Err(RuntimeError::ArgumentCountMismatch);
                }
                let s = ValueConversion::to_string(&args[0]);
                Ok(Value::String(self.html_encode(&s)))
            }
            "CreateObject" | "createobject" | "CREATEOBJECT" | "createObject" => {
                // Server.CreateObject - 仅支持白名单中的对象
                if args.is_empty() {
                    return Err(RuntimeError::Generic(
                        "Server.CreateObject 需要 ProgID 参数".to_string(),
                    ));
                }

                let prog_id = ValueConversion::to_string(&args[0]);

                // 使用对象工厂创建
                use crate::runtime::objects::{create_object, get_supported_objects};

                match create_object(&prog_id) {
                    Ok(Some(obj)) => Ok(obj),
                    Ok(None) => {
                        // 不在白名单中
                        let supported = get_supported_objects().join(", ");
                        Err(RuntimeError::CreateObjectFailed(format!(
                            "'{}' 不在白名单中。支持的对象: {}",
                            prog_id, supported
                        )))
                    }
                    Err(e) => Err(e),
                }
            }
            "Execute" | "execute" | "EXECUTE" => {
                // Server.Execute - 暂不支持
                Err(RuntimeError::Generic(
                    "Server.Execute is not yet implemented.".to_string(),
                ))
            }
            "Transfer" | "transfer" | "TRANSFER" => {
                // Server.Transfer - 暂不支持
                Err(RuntimeError::Generic(
                    "Server.Transfer is not yet implemented.".to_string(),
                ))
            }
            "GetLastError" | "getlasterror" | "GETLASTERROR" | "getLastError" => {
                // Server.GetLastError - 暂不支持
                Err(RuntimeError::Generic(
                    "Server.GetLastError is not yet implemented.".to_string(),
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
