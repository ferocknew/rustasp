//! New 表达式求值（类实例化）

use crate::runtime::{RuntimeError, Value};
use crate::runtime::objects::{Dictionary, FileSystemObject, XmlHttp};
use std::sync::{Arc, Mutex};

use super::super::Interpreter;

/// CreateObject 白名单
/// 只允许创建安全的、经过测试的对象
const CREATE_OBJECT_WHITELIST: &[&str] = &[
    "scripting.dictionary",
    "scripting.filesystemobject",
    "msxml2.xmlhttp",
];

impl Interpreter {
    /// 执行 New 表达式 - 创建类实例
    /// 支持内置对象和用户定义的类
    pub fn eval_new(&mut self, class_name: &str) -> Result<Value, RuntimeError> {
        let normalized_name = crate::utils::normalize_identifier(class_name);

        // 1. 先检查是否是内置对象
        if let Some(builtin) = Self::create_builtin_object(&normalized_name) {
            return Ok(builtin);
        }

        // 2. 从缓存中获取预编译的 VbsClass
        if let Some(vbs_class) = self.context.classes.get(&normalized_name) {
            // 直接使用缓存的类创建实例（避免重复构建）
            let instance = vbs_class.new_instance();
            return Ok(instance.to_value());
        }

        Err(RuntimeError::Generic(format!(
            "Class '{}' not found",
            class_name
        )))
    }

    /// 创建内置对象实例
    /// 返回 None 表示不是内置对象，需要查找用户定义的类
    fn create_builtin_object(class_name: &str) -> Option<Value> {
        // 检查白名单
        let is_whitelisted = CREATE_OBJECT_WHITELIST.iter().any(|&allowed| {
            class_name.eq_ignore_ascii_case(allowed)
        });

        if !is_whitelisted {
            // 不在白名单中的对象返回错误
            return None;
        }

        match class_name {
            // Scripting.Dictionary 对象
            "dictionary" | "scripting.dictionary" => {
                Some(Value::Object(Arc::new(Mutex::new(Dictionary::new()))))
            }
            // Scripting.FileSystemObject 对象
            "filesystemobject" | "scripting.filesystemobject" => {
                Some(Value::Object(Arc::new(Mutex::new(FileSystemObject::new()))))
            }
            // MSXML2.XMLHTTP 对象
            "xmlhttp" | "msxml2.xmlhttp" | "microsoft.xmlhttp" => {
                Some(Value::Object(Arc::new(Mutex::new(XmlHttp::new()))))
            }
            _ => None,
        }
    }

    /// Server.CreateObject 方法 - 创建 COM 对象（仅支持白名单）
    pub fn server_create_object(&mut self, prog_id: &str) -> Result<Value, RuntimeError> {
        let normalized_prog_id = crate::utils::normalize_identifier(prog_id);

        // 检查白名单
        let is_whitelisted = CREATE_OBJECT_WHITELIST.iter().any(|&allowed| {
            normalized_prog_id.eq_ignore_ascii_case(allowed)
        });

        if !is_whitelisted {
            return Err(RuntimeError::Generic(format!(
                "Server.CreateObject: '{}' 不在白名单中。出于安全考虑，只允许创建以下对象: {}",
                prog_id,
                CREATE_OBJECT_WHITELIST.join(", ")
            )));
        }

        // 创建内置对象
        if let Some(obj) = Self::create_builtin_object(&normalized_prog_id) {
            Ok(obj)
        } else {
            Err(RuntimeError::Generic(format!(
                "Server.CreateObject: 无法创建对象 '{}'",
                prog_id
            )))
        }
    }
}
