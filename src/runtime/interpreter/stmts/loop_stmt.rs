//! 循环语句执行
//!
//! 处理 For, While, ForEach 和 Exit 语句

use crate::ast::{Expr, Stmt};
use crate::runtime::{RuntimeError, Value};
use super::Interpreter;

impl Interpreter {
    /// 执行 For 循环
    pub fn eval_for(
        &mut self,
        var: &str,
        start: &Expr,
        end: &Expr,
        step: Option<&Expr>,
        body: &[Stmt],
    ) -> Result<Value, RuntimeError> {
        let start_val = self.eval_expr(start)?.to_number();
        let end_val = self.eval_expr(end)?.to_number();
        let step_val = step
            .map(|e| self.eval_expr(e).map(|v: Value| v.to_number()))
            .transpose()?
            .unwrap_or(1.0);

        let mut i = start_val;
        let condition = if step_val > 0.0 {
            move |i: f64, end: f64| i <= end
        } else {
            move |i: f64, end: f64| i >= end
        };

        while condition(i, end_val) {
            self.context.define_var(var.to_string(), Value::Number(i));
            self.exec_block(body)?;
            i += step_val;
        }

        Ok(Value::Empty)
    }

    /// 执行 While 循环
    pub fn eval_while(&mut self, cond: &Expr, body: &[Stmt]) -> Result<Value, RuntimeError> {
        while self.eval_expr(cond)?.is_truthy() {
            self.exec_block(body)?;
        }
        Ok(Value::Empty)
    }

    /// 执行 For Each 循环
    pub fn eval_for_each(&mut self, var: &str, collection: &Expr, body: &[Stmt]) -> Result<Value, RuntimeError> {
        let collection_val = self.eval_expr(collection)?;

        let elements = match collection_val {
            Value::Array(arr) => arr,
            Value::Object(mut obj) => {
                // 尝试作为字典处理
                use crate::runtime::objects::Dictionary;
                if let Some(dict) = obj.as_any().downcast_ref::<Dictionary>() {
                    dict.values()
                } else {
                    // 对于其他对象，尝试调用 items 方法
                    match obj.call_method("items", vec![]) {
                        Ok(Value::Array(arr)) => arr,
                        _ => {
                            return Err(RuntimeError::Generic(
                                "For Each requires an iterable object".to_string()
                            ))
                        }
                    }
                }
            }
            Value::String(s) => s.chars().map(|c| Value::String(c.to_string())).collect(),
            _ => {
                return Err(RuntimeError::Generic(format!(
                    "For Each requires an array, object, or string, got {:?}",
                    collection_val
                )))
            }
        };

        for element in elements {
            self.context.define_var(var.to_string(), element);
            self.exec_block(body)?;
        }

        Ok(Value::Empty)
    }

    /// Exit For
    pub fn eval_exit_for(&mut self) -> Result<Value, RuntimeError> {
        // TODO: 实现循环退出
        Ok(Value::Empty)
    }

    /// Exit Function/Sub
    pub fn eval_exit(&mut self) -> Result<Value, RuntimeError> {
        self.context.should_exit = true;
        Ok(Value::Empty)
    }

    /// 执行语句块
    pub fn exec_block(&mut self, stmts: &[Stmt]) -> Result<Value, RuntimeError> {
        for stmt in stmts {
            self.eval_stmt(stmt)?;
        }
        Ok(Value::Empty)
    }
}
