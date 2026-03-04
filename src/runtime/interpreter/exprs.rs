//! 表达式求值模块
//!
//! 处理各种 VBScript 表达式的求值逻辑

use crate::ast::{BinaryOp, Expr, UnaryOp};
use crate::runtime::{RuntimeError, Value, ValueCompare, ValueConversion, ValueOps};
use crate::utils::identifier_matches;

use super::Interpreter;

/// 表达式求值器
impl Interpreter {
    /// 求值表达式
    pub fn eval_expr(&mut self, expr: &Expr) -> Result<Value, RuntimeError> {
        match expr {
            Expr::Number(n) => Ok(Value::Number(*n)),
            Expr::String(s) => Ok(Value::String(s.clone())),
            Expr::Boolean(b) => Ok(Value::Boolean(*b)),
            Expr::Nothing => Ok(Value::Nothing),
            Expr::Empty => Ok(Value::Empty),
            Expr::Null => Ok(Value::Null),
            Expr::Variable(name) => self
                .context
                .get_var(name)
                .cloned()
                .ok_or_else(|| RuntimeError::UndefinedVariable(name.clone())),
            Expr::Binary { left, op, right } => {
                let left_val = self.eval_expr(left)?;
                let right_val = self.eval_expr(right)?;
                match op {
                    BinaryOp::Eq
                    | BinaryOp::Ne
                    | BinaryOp::Lt
                    | BinaryOp::Le
                    | BinaryOp::Gt
                    | BinaryOp::Ge
                    | BinaryOp::And
                    | BinaryOp::Or
                    | BinaryOp::Xor
                    | BinaryOp::Is => Ok(left_val.compare(*op, &right_val)),
                    _ => left_val.binary_op(*op, &right_val),
                }
            }
            Expr::Unary { op, operand } => {
                let val = self.eval_expr(operand)?;
                match op {
                    UnaryOp::Neg => Ok(Value::Number(-val.to_number())),
                    UnaryOp::Not => Ok(Value::Boolean(!val.is_truthy())),
                }
            }
            Expr::Call { name: _, args } => {
                let arg_values: Result<Vec<Value>, _> =
                    args.iter().map(|e| self.eval_expr(e)).collect();
                let _arg_values = arg_values?;
                // TODO: 实现函数调用
                Ok(Value::Empty)
            }
            Expr::Property { object, property } => {
                let _obj = self.eval_expr(object)?;
                // TODO: 实现属性访问
                Err(RuntimeError::PropertyNotFound(property.clone()))
            }
            Expr::Method { object, method, args } => {
                self.eval_method(object, method, args)
            }
            Expr::Array(elements) => {
                let values: Result<Vec<Value>, _> =
                    elements.iter().map(|e| self.eval_expr(e)).collect();
                Ok(Value::Array(values?))
            }
            Expr::Index { object, index } => self.eval_index(object, index),
            _ => Err(RuntimeError::Generic(format!(
                "Unimplemented expr: {:?}",
                expr
            ))),
        }
    }

    /// 处理索引表达式
    fn eval_index(&mut self, object: &Expr, index: &Expr) -> Result<Value, RuntimeError> {
        let index_val = self.eval_expr(index)?;
        let index_key = match &index_val {
            Value::String(s) => s.clone(),
            _ => ValueConversion::to_string(&index_val),
        };

        match object {
            // 特殊处理 Request 对象
            Expr::Variable(name) if identifier_matches(name, "request") => {
                // 从 request_data 中获取值
                match self.context.get_request_param(&index_key) {
                    Some(value) => Ok(Value::String(value.clone())),
                    None => Ok(Value::Empty),
                }
            }
            // 处理数组或对象/字典访问，或内置函数调用
            Expr::Variable(name) => {
                // 检查是否是内置函数调用
                if let Some(result) =
                    super::builtins::call_builtin_function(&name.to_lowercase(), &index_val)
                {
                    return result;
                }

                // 检查变量是否是数组或对象
                if let Some(value) = self.context.get_var(name).cloned() {
                    match value {
                        Value::Array(arr) => {
                            if let Value::Number(i) = index_val {
                                let i = i as usize;
                                if i < arr.len() {
                                    return Ok(arr[i].clone());
                                }
                            }
                        }
                        Value::Object(obj) => {
                            if let Some(v) = obj.get(&index_key) {
                                return Ok(v.clone());
                            }
                        }
                        _ => {}
                    }
                }
                Err(RuntimeError::InvalidIndex)
            }
            _ => Err(RuntimeError::InvalidIndex),
        }
    }

    /// 执行方法调用
    pub fn eval_method(
        &mut self,
        object: &Expr,
        method: &str,
        args: &[Expr],
    ) -> Result<Value, RuntimeError> {
        // 获取对象名称（如果是变量）
        let object_name = match object {
            Expr::Variable(name) => Some(name.to_lowercase()),
            _ => None,
        };

        let method_lower = method.to_lowercase();

        // 处理内建对象的方法
        match (object_name.as_deref(), method_lower.as_str()) {
            // Response.Write
            (Some("response"), "write") => {
                if !args.is_empty() {
                    let value = self.eval_expr(&args[0])?;
                    self.context.write(&ValueConversion::to_string(&value));
                }
                Ok(Value::Empty)
            }
            // Response.End
            (Some("response"), "end") => {
                // TODO: 实现响应结束
                Ok(Value::Empty)
            }
            // Response.Redirect
            (Some("response"), "redirect") => {
                // TODO: 实现重定向
                Ok(Value::Empty)
            }
            // Request.QueryString / Request.Form
            (Some("request"), "querystring") => {
                if !args.is_empty() {
                    let key = self.eval_expr(&args[0])?;
                    let key_str = ValueConversion::to_string(&key);
                    // 从上下文获取 QueryString
                    Ok(self.context.get_var(&key_str).cloned().unwrap_or(Value::Empty))
                } else {
                    Ok(Value::Empty)
                }
            }
            (Some("request"), "form") => {
                if !args.is_empty() {
                    let key = self.eval_expr(&args[0])?;
                    let key_str = ValueConversion::to_string(&key);
                    // 从上下文获取 Form 数据
                    Ok(self.context.get_var(&key_str).cloned().unwrap_or(Value::Empty))
                } else {
                    Ok(Value::Empty)
                }
            }
            // Server.CreateObject
            (Some("server"), "createobject") => {
                // 不支持 COM 对象创建
                Err(RuntimeError::Generic(
                    "COM object creation is not supported".to_string(),
                ))
            }
            // 其他方法调用
            _ => {
                // 尝试调用用户定义的方法
                let arg_values: Result<Vec<Value>, _> =
                    args.iter().map(|e| self.eval_expr(e)).collect();
                let _arg_values = arg_values?;
                // TODO: 实现用户定义方法调用
                Ok(Value::Empty)
            }
        }
    }
}
