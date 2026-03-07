//! 属性访问表达式求值

use crate::ast::Expr;
use crate::runtime::{BuiltinObject, RuntimeError, Value};

use super::super::Interpreter;

impl Interpreter {
    /// 处理属性访问表达式
    pub fn eval_property(&mut self, object: &Expr, property: &str) -> Result<Value, RuntimeError> {
        let property_lower = property.to_lowercase();

        // 统一 trait dispatch：eval object 并访问属性
        let obj_value = self.eval_expr(object)?;

        // 特殊处理：.Count 属性（适用于多种类型）
        if property_lower == "count" {
            return Ok(Self::get_count_value(&obj_value));
        }

        // 处理对象的属性访问
        if let Value::Object(obj) = obj_value {
            return obj.get_property(&property_lower);
        }

        Err(RuntimeError::PropertyNotFound(property.to_string()))
    }

    /// 获取 Count 属性值
    fn get_count_value(value: &Value) -> Value {
        match value {
            Value::Array(arr) => Value::Number(arr.len() as f64),
            Value::Object(obj) => obj.get_property("count").unwrap_or(Value::Number(0.0)),
            Value::String(_) | Value::Number(_) | Value::Boolean(_)
            | Value::Empty | Value::Null | Value::Nothing => Value::Number(1.0),
        }
    }
}
