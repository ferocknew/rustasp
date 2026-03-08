//! 赋值语句执行模块
//!
//! 处理 Assignment、Set 及索引/属性赋值

use crate::ast::Expr;
use crate::runtime::{RuntimeError, Value};

use super::Interpreter;

/// 赋值语句执行器
impl Interpreter {
    /// 执行赋值语句
    pub fn eval_assignment(
        &mut self,
        target: &Expr,
        value: &Expr,
    ) -> Result<Value, RuntimeError> {
        let val = self.eval_expr(value)?;

        match target {
            Expr::Variable(name) => {
                self.context.set_var(name.clone(), val);
                Ok(Value::Empty)
            }
            Expr::Index { object, index } => self.eval_index_assignment(object, index, val),
            Expr::Property { object, property } => {
                self.eval_property_assignment(object, property, val)
            }
            _ => Err(RuntimeError::InvalidAssignment),
        }
    }

    /// 执行 Set 语句
    pub fn eval_set(&mut self, target: &Expr, value: &Expr) -> Result<Value, RuntimeError> {
        self.eval_assignment(target, value)
    }

    /// 执行索引赋值（如 arr(i) = value 或 Session("key") = value）
    fn eval_index_assignment(
        &mut self,
        object: &Expr,
        index: &Expr,
        val: Value,
    ) -> Result<Value, RuntimeError> {
        // 处理多维数组索引，例如 arr2D(0)(0) = value
        if let Expr::Index { .. } = object {
            return self.eval_nested_index_assignment(object, index, val);
        }

        let idx = self.eval_expr(index)?;

        match object {
            Expr::Variable(name) => {
                // 先尝试从变量表获取对象
                if let Some(obj_val) = self.context.get_var(name) {
                    // 如果是对象，使用 trait 的 set_index 方法
                    // Arc<Mutex<dyn BuiltinObject>> 会直接修改共享对象，无需写回
                    if let Value::Object(ref obj) = obj_val {
                        let result = obj.lock()
                            .map_err(|_| RuntimeError::Generic("Failed to lock object".to_string()))?
                            .set_index(&idx, val.clone());
                        return result.map(|_| Value::Empty);
                    }
                }

                // 处理数组索引
                if let Value::Number(i) = idx {
                    let i = i as usize;
                    match self.context.get_var(name).cloned() {
                        Some(Value::Array(mut arr)) => {
                            if i >= arr.len() {
                                arr.resize(i + 1, Value::Empty);
                            }
                            arr[i] = val;
                            self.context.set_var(name.clone(), Value::Array(arr));
                            Ok(Value::Empty)
                        }
                        Some(Value::Empty) => {
                            let mut arr = vec![Value::Empty; i + 1];
                            arr[i] = val;
                            self.context.set_var(name.clone(), Value::Array(arr));
                            Ok(Value::Empty)
                        }
                        _ => {
                            let mut arr = vec![Value::Empty; i + 1];
                            arr[i] = val;
                            self.context.set_var(name.clone(), Value::Array(arr));
                            Ok(Value::Empty)
                        }
                    }
                } else {
                    Err(RuntimeError::Generic(format!(
                        "Array index must be a number, got {:?}",
                        idx
                    )))
                }
            }
            _ => Err(RuntimeError::InvalidAssignment),
        }
    }

    /// 执行嵌套索引赋值（多维数组）
    fn eval_nested_index_assignment(
        &mut self,
        object: &Expr,
        index: &Expr,
        val: Value,
    ) -> Result<Value, RuntimeError> {
        let (var_name, mut indices) = self.flatten_index_expression(object, index)?;
        indices.reverse();

        // 获取数组
        let mut arr = match self.context.get_var(&var_name).cloned() {
            Some(Value::Array(arr)) => arr,
            Some(Value::Empty) => vec![],
            _ => {
                return Err(RuntimeError::Generic(format!(
                    "'{}' is not an array",
                    var_name
                )))
            }
        };

        // 计算扁平索引
        let flat_index = if indices.len() == 1 {
            match &indices[0] {
                Value::Number(i) => *i as usize,
                _ => {
                    return Err(RuntimeError::Generic(
                        "Array index must be a number".to_string(),
                    ))
                }
            }
        } else {
            let mut result: usize = 0;
            for idx in &indices {
                match idx {
                    Value::Number(n) => {
                        result = result * 3 + (*n as usize);
                    }
                    _ => {
                        return Err(RuntimeError::Generic(
                            "Array index must be a number".to_string(),
                        ))
                    }
                }
            }
            result
        };

        // 扩展数组并赋值
        if flat_index >= arr.len() {
            arr.resize(flat_index + 1, Value::Empty);
        }
        arr[flat_index] = val;
        self.context.set_var(var_name.clone(), Value::Array(arr));
        Ok(Value::Empty)
    }

    /// 展平索引表达式
    fn flatten_index_expression(
        &mut self,
        object: &Expr,
        index: &Expr,
    ) -> Result<(String, Vec<Value>), RuntimeError> {
        let mut indices = vec![];
        let mut current_expr = object;
        indices.push(self.eval_expr(index)?);

        loop {
            match current_expr {
                Expr::Index {
                    object: inner_object,
                    index: inner_index,
                } => {
                    indices.push(self.eval_expr(inner_index)?);
                    current_expr = inner_object;
                }
                Expr::Variable(name) => {
                    return Ok((name.clone(), indices));
                }
                _ => return Err(RuntimeError::InvalidAssignment),
            }
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
                        let result = obj.lock()
                            .map_err(|_| RuntimeError::Generic("Failed to lock object".to_string()))?
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
                    let result = obj.lock()
                        .map_err(|_| RuntimeError::Generic("Failed to lock object".to_string()))?
                        .set_property(property, val);
                    return result.map(|_| Value::Empty);
                }
                Err(RuntimeError::InvalidAssignment)
            }
        }
    }
}
