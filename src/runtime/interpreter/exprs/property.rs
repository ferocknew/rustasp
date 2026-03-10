//! 属性访问表达式求值

use crate::ast::Expr;
use crate::runtime::{RuntimeError, Value, VbsInstance};

use super::super::Interpreter;

impl Interpreter {
    /// 处理属性访问表达式
    pub fn eval_property(&mut self, object: &Expr, property: &str) -> Result<Value, RuntimeError> {
        let property_lower = property.to_lowercase();

        // 特殊处理：Err 对象属性访问
        if let Expr::Variable(name) = object {
            if name.to_lowercase() == "err" {
                return self.eval_err_property(&property_lower);
            }
        }

        // 特殊处理：Response 对象的无参数方法（End, Clear）
        if let Expr::Variable(name) = object {
            if name.to_lowercase() == "response" {
                match property_lower.as_str() {
                    "end" => {
                        // 调用 Response.End 方法
                        return self.eval_response_end();
                    }
                    "clear" => {
                        // 调用 Response.Clear 方法
                        return self.eval_response_clear();
                    }
                    _ => {}
                }
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
            // 尝试获取属性
            let prop_result = obj
                .lock()
                .map_err(|_| RuntimeError::Generic("Failed to lock object".to_string()))?
                .get_property(&property_lower);

            match prop_result {
                Ok(value) => return Ok(value),
                Err(_) => {
                    // 属性未找到，检查是否是类实例的无参数方法
                    let locked = obj
                        .lock()
                        .map_err(|_| RuntimeError::Generic("Failed to lock object".to_string()))?;

                    if let Some(instance) = locked.as_any().downcast_ref::<VbsInstance>() {
                        // 获取类名
                        let class_name = instance.class_name.clone();
                        let normalized_class = crate::utils::normalize_identifier(&class_name);

                        // 释放锁后再获取类定义
                        drop(locked);

                        // 从上下文获取类定义并查找方法
                        if let Some(vbs_class) = self.context.classes.get(&normalized_class) {
                            // 查找方法（大小写不敏感）
                            let method_decl = vbs_class
                                .methods
                                .iter()
                                .find(|(name, _)| {
                                    crate::utils::normalize_identifier(name) == property_lower
                                })
                                .map(|(_, decl)| decl.clone());

                            if let Some(method_decl) = method_decl {
                                // 找到方法，执行无参数调用
                                let instance_fields = {
                                    let locked = obj.lock().map_err(|_| {
                                        RuntimeError::Generic("Failed to lock object".to_string())
                                    })?;
                                    if let Some(inst) =
                                        locked.as_any().downcast_ref::<VbsInstance>()
                                    {
                                        inst.fields.clone()
                                    } else {
                                        return Err(RuntimeError::Generic(
                                            "Invalid instance".to_string(),
                                        ));
                                    }
                                };

                                // 调用 method.rs 中定义的 execute_class_method
                                return self.execute_class_method(
                                    &method_decl,
                                    &[],
                                    instance_fields,
                                    obj.clone(),
                                );
                            }
                        }
                    }
                    return prop_result;
                }
            }
        }

        Err(RuntimeError::PropertyNotFound(property.to_string()))
    }

    /// 处理 Err 对象的属性访问
    fn eval_err_property(&mut self, property: &str) -> Result<Value, RuntimeError> {
        match property {
            "number" => Ok(Value::Number(self.context.err.get_number() as f64)),
            "description" => Ok(Value::String(
                self.context.err.get_description().to_string(),
            )),
            "clear" => {
                // VBScript 允许不带括号调用方法
                self.context.err.clear();
                Ok(Value::Empty)
            }
            _ => Err(RuntimeError::PropertyNotFound(format!("Err.{}", property))),
        }
    }

    /// 处理 Response.End 方法调用
    fn eval_response_end(&mut self) -> Result<Value, RuntimeError> {
        // 获取 Response 对象并调用 End 方法
        if let Some(value) = self.context.get_var("Response") {
            if let Value::Object(obj) = value {
                let mut locked = obj.lock().map_err(|_| {
                    RuntimeError::Generic("Failed to lock Response object".to_string())
                })?;
                locked.call_method("end", vec![])?;
            }
        }
        // 设置退出标志
        self.context.set_should_exit();
        Ok(Value::Empty)
    }

    /// 处理 Response.Clear 方法调用
    fn eval_response_clear(&mut self) -> Result<Value, RuntimeError> {
        // 获取 Response 对象并调用 Clear 方法
        if let Some(value) = self.context.get_var("Response") {
            if let Value::Object(obj) = value {
                let mut locked = obj.lock().map_err(|_| {
                    RuntimeError::Generic("Failed to lock Response object".to_string())
                })?;
                locked.call_method("clear", vec![])?;
            }
        }
        Ok(Value::Empty)
    }

    /// 获取 Count 属性值
    fn get_count_value(value: &Value) -> Value {
        match value {
            Value::Array(arr) => {
                let locked_arr = arr.lock().unwrap();
                Value::Number(locked_arr.len() as f64)
            }
            Value::Object(obj) => obj
                .lock()
                .ok()
                .and_then(|o| o.get_property("count").ok())
                .unwrap_or(Value::Number(0.0)),
            Value::String(_)
            | Value::Number(_)
            | Value::Boolean(_)
            | Value::Empty
            | Value::Null
            | Value::Nothing => Value::Number(1.0),
        }
    }
}
