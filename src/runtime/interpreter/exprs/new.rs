//! New 表达式求值（类实例化）

use crate::runtime::{RuntimeError, Value};
use crate::runtime::objects::{create_object, get_supported_objects};

use super::super::Interpreter;

impl Interpreter {
    /// 执行 New 表达式 - 创建类实例
    /// 支持内置对象和用户定义的类
    pub fn eval_new(&mut self, class_name: &str) -> Result<Value, RuntimeError> {
        // 1. 先检查是否是内置对象
        if let Some(builtin) = Self::create_builtin_object(class_name) {
            return Ok(builtin);
        }

        // 2. 从缓存中获取预编译的 VbsClass
        let normalized_name = crate::utils::normalize_identifier(class_name);
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
    fn create_builtin_object(prog_id: &str) -> Option<Value> {
        // 使用对象工厂创建
        match create_object(prog_id) {
            Ok(Some(obj)) => Some(obj),
            _ => None,
        }
    }

    /// Server.CreateObject 方法 - 创建 COM 对象（仅支持白名单）
    pub fn server_create_object(&mut self, prog_id: &str) -> Result<Value, RuntimeError> {
        match create_object(prog_id) {
            Ok(Some(obj)) => Ok(obj),
            Ok(None) => {
                // 不在白名单中
                let supported = get_supported_objects().join(", ");
                Err(RuntimeError::Generic(format!(
                    "Server.CreateObject: '{}' 不在白名单中。支持的对象: {}",
                    prog_id,
                    supported
                )))
            }
            Err(e) => Err(e),
        }
    }
}
