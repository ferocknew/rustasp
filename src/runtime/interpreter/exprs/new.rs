//! New 表达式求值（类实例化）

use crate::runtime::objects::{create_object, get_supported_objects};
use crate::runtime::{RuntimeError, Value};

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
            let instance_value = instance.to_value();

            // 克隆构造函数体（需要在借用 vbs_class 之前完成）
            let init_body = vbs_class
                .get_method("class_initialize")
                .map(|m| m.body.clone());

            // 调用构造函数 Class_Initialize（如果存在）
            if let Some(body) = init_body {
                for stmt in &body {
                    // 需要创建一个临时作用域，让 'Me' 指向当前实例
                    self.context.push_scope();

                    // 设置 Me 变量指向当前实例
                    self.context
                        .set_var("Me".to_string(), instance_value.clone());

                    // 执行构造函数体（暂时忽略错误，因为构造函数不应该有返回值）
                    let _ = self.eval_stmt(stmt);

                    // 恢复作用域
                    self.context.pop_scope();
                }
            }

            return Ok(instance_value);
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
                    prog_id, supported
                )))
            }
            Err(e) => Err(e),
        }
    }
}
