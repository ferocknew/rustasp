//! 赋值语句执行模块
//!
//! 处理 Assignment、Set 及索引/属性赋值

use crate::ast::Expr;
use crate::runtime::{vb_error, ErrorMode, RuntimeError, Value};
use std::sync::{Arc, Mutex};

use super::Interpreter;

/// 从 RuntimeError 中提取错误信息（简化版本）
fn extract_error_info(error: &RuntimeError) -> (i32, String) {
    match error {
        RuntimeError::DivisionByZero => {
            (vb_error::DIVISION_BY_ZERO, "Division by zero".to_string())
        }
        RuntimeError::TypeMismatch(msg) => {
            (vb_error::TYPE_MISMATCH, format!("Type mismatch: {}", msg))
        }
        RuntimeError::ObjectRequired => (vb_error::OBJECT_REQUIRED, "Object required".to_string()),
        RuntimeError::UndefinedFunction(name) => (
            vb_error::UNDEFINED_FUNCTION,
            format!("Undefined function: {}", name),
        ),
        RuntimeError::IndexOutOfBounds(_) => (
            vb_error::SUBSCRIPT_OUT_OF_RANGE,
            "Subscript out of range".to_string(),
        ),
        RuntimeError::CreateObjectFailed(msg) => (
            vb_error::CANT_CREATE_OBJECT,
            format!("Server.CreateObject: {}", msg),
        ),
        _ => (0, format!("{:?}", error)),
    }
}

/// 赋值语句执行器
impl Interpreter {
    /// 执行赋值语句
    pub fn eval_assignment(&mut self, target: &Expr, value: &Expr) -> Result<Value, RuntimeError> {
        // 获取当前错误模式
        let error_mode = self.context.current_scope().get_error_mode();

        // 尝试计算值
        let val_result = self.eval_expr(value);

        // 处理结果
        match val_result {
            Ok(val) => match target {
                Expr::Variable(name) => {
                    self.context.set_var(name.clone(), val);
                    Ok(Value::Empty)
                }
                Expr::Index { object, indices } => self.eval_index_assignment(object, indices, val),
                Expr::Property { object, property } => {
                    self.eval_property_assignment(object, property, val)
                }
                _ => Err(RuntimeError::InvalidAssignment),
            },
            Err(e) => {
                match target {
                    Expr::Variable(name) => {
                        match error_mode {
                            ErrorMode::Stop => Err(e),
                            ErrorMode::ResumeNext => {
                                // 设置变量为 Empty
                                self.context.set_var(name.clone(), Value::Empty);
                                // 返回错误，由外层处理
                                Err(e)
                            }
                        }
                    }
                    _ => Err(e),
                }
            }
        }
    }

    /// 执行 Set 语句
    pub fn eval_set(&mut self, target: &Expr, value: &Expr) -> Result<Value, RuntimeError> {
        // 获取当前错误模式
        let error_mode = self.context.current_scope().get_error_mode();

        // 尝试计算值
        let val_result = self.eval_expr(value);

        // 处理结果
        match val_result {
            Ok(val) => {
                // Set 语句要求右边必须是对象、Nothing 或 Empty
                // 如果是 Empty，视为错误（Object required）
                if matches!(val, Value::Empty) {
                    match target {
                        Expr::Variable(name) => {
                            match error_mode {
                                ErrorMode::Stop => return Err(RuntimeError::ObjectRequired),
                                ErrorMode::ResumeNext => {
                                    // 设置变量为 Nothing (Null)
                                    self.context.set_var(name.clone(), Value::Null);
                                    return Ok(Value::Empty);
                                }
                            }
                        }
                        _ => return Err(RuntimeError::ObjectRequired),
                    }
                }

                match target {
                    Expr::Variable(name) => {
                        self.context.set_var(name.clone(), val);
                        Ok(Value::Empty)
                    }
                    Expr::Index { object, indices } => {
                        self.eval_index_assignment(object, indices, val)
                    }
                    Expr::Property { object, property } => {
                        self.eval_property_assignment(object, property, val)
                    }
                    _ => Err(RuntimeError::InvalidAssignment),
                }
            }
            Err(e) => {
                match target {
                    Expr::Variable(name) => {
                        match error_mode {
                            ErrorMode::Stop => Err(e),
                            ErrorMode::ResumeNext => {
                                // 记录错误到 Err 对象
                                let (number, description) = extract_error_info(&e);
                                self.context.err.set(number, description);
                                // 设置变量为 Nothing (Null)
                                self.context.set_var(name.clone(), Value::Null);
                                // 返回成功，避免外层再次处理错误
                                Ok(Value::Empty)
                            }
                        }
                    }
                    _ => Err(e),
                }
            }
        }
    }

    /// 执行索引赋值（如 arr(i) = value 或 dict("key") = value）
    /// 支持多层对象属性访问后的索引赋值，如 Easp.Lang("key") = value
    fn eval_index_assignment(
        &mut self,
        object: &Expr,
        indices: &[Expr],
        val: Value,
    ) -> Result<Value, RuntimeError> {
        // 求值所有索引
        let index_vals: Result<Vec<Value>, RuntimeError> =
            indices.iter().map(|idx| self.eval_expr(idx)).collect();
        let index_vals = index_vals?;

        match object {
            Expr::Variable(name) => {
                // 先尝试从变量表获取对象
                if let Some(obj_val) = self.context.get_var(name) {
                    // 如果是对象，使用 trait 的 set_index 方法
                    if let Value::Object(ref obj) = obj_val {
                        if indices.len() == 1 {
                            let result = obj
                                .lock()
                                .map_err(|_| {
                                    RuntimeError::Generic("Failed to lock object".to_string())
                                })?
                                .set_index(&index_vals[0], val.clone());
                            return result.map(|_| Value::Empty);
                        }
                    }
                }

                // 处理数组索引赋值
                let idx: Result<Vec<usize>, RuntimeError> = index_vals
                    .iter()
                    .map(|v| match v {
                        Value::Number(n) => Ok(*n as usize),
                        _ => Err(RuntimeError::Generic(format!(
                            "Array index must be a number, got {:?}",
                            v
                        ))),
                    })
                    .collect();

                let idx = idx?;

                match self.context.get_var(name) {
                    Some(Value::Array(ref arr)) => {
                        let mut locked_arr = arr.lock().map_err(|_| {
                            RuntimeError::Generic("Failed to lock array".to_string())
                        })?;

                        // 使用 flat_index 计算索引
                        match locked_arr.flat_index(&idx) {
                            Some(flat_idx) => {
                                locked_arr.data[flat_idx] = val;
                                Ok(Value::Empty)
                            }
                            None => Err(RuntimeError::Generic(format!(
                                "Array index out of bounds: {:?}",
                                idx
                            ))),
                        }
                    }
                    Some(Value::Empty) => {
                        // 动态创建数组
                        use crate::runtime::VbsArray;
                        let mut vbs_arr = VbsArray::new(vec![idx[0] + 1]);
                        if idx.len() == 1 && idx[0] < vbs_arr.data.len() {
                            vbs_arr.data[idx[0]] = val;
                        }
                        self.context
                            .set_var(name.clone(), Value::Array(Arc::new(Mutex::new(vbs_arr))));
                        Ok(Value::Empty)
                    }
                    _ => {
                        // 变量不是数组，创建新数组
                        use crate::runtime::VbsArray;
                        let mut vbs_arr = VbsArray::new(vec![idx[0] + 1]);
                        if idx.len() == 1 && idx[0] < vbs_arr.data.len() {
                            vbs_arr.data[idx[0]] = val;
                        }
                        self.context
                            .set_var(name.clone(), Value::Array(Arc::new(Mutex::new(vbs_arr))));
                        Ok(Value::Empty)
                    }
                }
            }
            Expr::Property {
                object: prop_obj,
                property: _,
            } => {
                // 处理属性访问后的索引赋值，如 Easp.Lang("key") = value
                // 先求值属性访问表达式，得到对象
                let obj_val = if let Expr::Variable(var_name) = prop_obj.as_ref() {
                    // 从变量表获取对象
                    self.context
                        .get_var(var_name)
                        .ok_or_else(|| RuntimeError::UndefinedVariable(var_name.clone()))?
                        .clone()
                } else {
                    // 递归求值
                    self.eval_expr(prop_obj)?
                };

                // 获取属性值（应该是一个对象）
                if let Value::Object(ref obj) = obj_val {
                    let mut obj_guard = obj
                        .lock()
                        .map_err(|_| RuntimeError::Generic("Failed to lock object".to_string()))?;

                    // 对对象进行索引赋值
                    if indices.len() == 1 {
                        obj_guard.set_index(&index_vals[0], val.clone())?;
                        return Ok(Value::Empty);
                    } else {
                        return Err(RuntimeError::Generic(
                            "Multi-dimensional index assignment not supported for objects"
                                .to_string(),
                        ));
                    }
                } else {
                    return Err(RuntimeError::ObjectRequired);
                }
            }
            _ => Err(RuntimeError::InvalidAssignment),
        }
    }

    /// 执行属性赋值（如 Response.Buffer = value）
    /// 统一使用 trait dispatch
    fn eval_property_assignment(
        &mut self,
        object: &Expr,
        property: &str,
        val: Value,
    ) -> Result<Value, RuntimeError> {
        // 获取对象表达式并求值
        match object {
            Expr::Variable(var_name) => {
                // 从变量表获取对象
                if let Some(obj_val) = self.context.get_var(var_name) {
                    if let Value::Object(ref obj) = obj_val {
                        // Arc<Mutex<dyn BuiltinObject>> 会直接修改共享对象，无需写回
                        let result = obj
                            .lock()
                            .map_err(|_| {
                                RuntimeError::Generic("Failed to lock object".to_string())
                            })?
                            .set_property(property, val);
                        return result.map(|_| Value::Empty);
                    }
                }
                Err(RuntimeError::PropertyNotFound(format!(
                    "{}.{}",
                    var_name, property
                )))
            }
            _ => {
                // 对于非变量的对象表达式（如 obj.prop.subprop），先求值
                let obj_val = self.eval_expr(object)?;
                if let Value::Object(ref obj) = obj_val {
                    let result = obj
                        .lock()
                        .map_err(|_| RuntimeError::Generic("Failed to lock object".to_string()))?
                        .set_property(property, val);
                    return result.map(|_| Value::Empty);
                }
                Err(RuntimeError::InvalidAssignment)
            }
        }
    }
}
