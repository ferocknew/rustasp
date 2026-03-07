//! 属性访问表达式求值

use crate::ast::Expr;
use crate::runtime::{BuiltinObject, RuntimeError, Value};

use super::super::Interpreter;

impl Interpreter {
    /// 处理属性访问表达式
    pub fn eval_property(&mut self, object: &Expr, property: &str) -> Result<Value, RuntimeError> {
        let property_lower = property.to_lowercase();

        // 特殊处理：Response 对象（存储在 context 中，不在变量表中）
        if let Expr::Variable(name) = object {
            if crate::utils::normalize_identifier(name) == "response" {
                return self.eval_response_property(&property_lower);
            }
        }

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

    /// 处理 Response 对象的属性访问
    fn eval_response_property(&mut self, property: &str) -> Result<Value, RuntimeError> {
        let response = self.context.response();
        
        // 先尝试通过 trait 获取属性
        if let Ok(val) = response.get_property(property) {
            return Ok(val);
        }
        
        // 特殊处理：Response.Clear/End 作为属性访问（无参数调用）
        match property {
            "clear" => {
                self.context.response_mut().clear();
                Ok(Value::Empty)
            }
            "end" => {
                self.context.response_mut().end();
                self.context.should_exit = true;
                Ok(Value::Empty)
            }
            _ => Ok(Value::Empty),
        }
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
