//! 对象工厂 - 用于 CreateObject 白名单管理

use crate::runtime::objects::{Dictionary, FileSystemObject, XmlHttp};
use crate::runtime::{RuntimeError, Value};
use std::sync::{Arc, Mutex};

/// 对象工厂函数类型
type ObjectFactory = fn() -> Value;

/// 创建 Dictionary 对象
fn create_dictionary() -> Value {
    Value::Object(Arc::new(Mutex::new(Dictionary::new())))
}

/// 创建 FileSystemObject 对象
fn create_filesystemobject() -> Value {
    Value::Object(Arc::new(Mutex::new(FileSystemObject::new())))
}

/// 创建 XmlHttp 对象
fn create_xmlhttp() -> Value {
    Value::Object(Arc::new(Mutex::new(XmlHttp::new())))
}

/// 根据ProgID创建对象
/// 返回 Ok(Some(value)) 表示创建成功
/// 返回 Ok(None) 表示 ProgID 不在白名单中
/// 返回 Err 表示创建失败
pub fn create_object(prog_id: &str) -> Result<Option<Value>, RuntimeError> {
    let prog_id_lower = prog_id.to_lowercase();

    // 使用 match 匹配 ProgID
    let factory: Option<ObjectFactory> = match prog_id_lower.as_str() {
        // Scripting.Dictionary 对象
        "scripting.dictionary" | "dictionary" => Some(create_dictionary),

        // Scripting.FileSystemObject 对象
        "scripting.filesystemobject" | "filesystemobject" => Some(create_filesystemobject),

        // MSXML2.XMLHTTP 对象
        "msxml2.xmlhttp" | "microsoft.xmlhttp" | "xmlhttp" => Some(create_xmlhttp),

        // 不在白名单中
        _ => None,
    };

    match factory {
        Some(f) => Ok(Some(f())),
        None => Ok(None),
    }
}

/// 获取支持的对象列表（用于错误提示）
pub fn get_supported_objects() -> Vec<&'static str> {
    vec![
        "scripting.dictionary",
        "scripting.filesystemobject",
        "msxml2.xmlhttp",
    ]
}

/// 检查 ProgID 是否在白名单中
pub fn is_whitelisted(prog_id: &str) -> bool {
    let prog_id_lower = prog_id.to_lowercase();
    matches!(
        prog_id_lower.as_str(),
        "scripting.dictionary"
            | "dictionary"
            | "scripting.filesystemobject"
            | "filesystemobject"
            | "msxml2.xmlhttp"
            | "microsoft.xmlhttp"
            | "xmlhttp"
    )
}
