//! 控制流语句执行模块
//!
//! 处理 If、For、While、ForEach、Select Case 等控制流语句

use crate::ast::{BinaryOp, CaseClause, Expr, IfBranch};
use crate::runtime::{RuntimeError, Value, ValueCompare, ValueConversion};

use super::Interpreter;

/// 控制流语句执行器
impl Interpreter {
    /// 执行 If 语句
    pub fn eval_if(
        &mut self,
        branches: &[IfBranch],
        else_block: &Option<Vec<crate::ast::Stmt>>,
    ) -> Result<Value, RuntimeError> {
        for branch in branches {
            let cond = self.eval_expr(&branch.cond)?;
            if cond.is_truthy() {
                return self.exec_block(&branch.body);
            }
        }
        if let Some(else_stmts) = else_block {
            self.exec_block(else_stmts)?;
        }
        Ok(Value::Empty)
    }

    /// 执行 For 循环
    pub fn eval_for(
        &mut self,
        var: &str,
        start: &Expr,
        end: &Expr,
        step: Option<&Expr>,
        body: &[crate::ast::Stmt],
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
    pub fn eval_while(
        &mut self,
        cond: &Expr,
        body: &[crate::ast::Stmt],
    ) -> Result<Value, RuntimeError> {
        while self.eval_expr(cond)?.is_truthy() {
            self.exec_block(body)?;
        }
        Ok(Value::Empty)
    }

    /// 执行 For Each 循环
    pub fn eval_for_each(
        &mut self,
        var: &str,
        collection: &Expr,
        body: &[crate::ast::Stmt],
    ) -> Result<Value, RuntimeError> {
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
                                "For Each requires an iterable object".to_string(),
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

    /// 执行 Select Case 语句
    pub fn eval_select(
        &mut self,
        expr: &Expr,
        cases: &[CaseClause],
        else_block: &Option<Vec<crate::ast::Stmt>>,
    ) -> Result<Value, RuntimeError> {
        let select_value = self.eval_expr(expr)?;

        for case in cases {
            if let Some(values) = &case.values {
                for value_expr in values {
                    let case_value = self.eval_expr(value_expr)?;
                    let result = select_value.compare(BinaryOp::Eq, &case_value);
                    if let Value::Boolean(true) = result {
                        return self.exec_block(&case.body);
                    }
                }
            }
        }

        if let Some(else_stmts) = else_block {
            self.exec_block(else_stmts)?;
        }

        Ok(Value::Empty)
    }

    /// 执行 Call 语句
    pub fn eval_call(
        &mut self,
        name: &str,
        args: &[Expr],
    ) -> Result<Value, RuntimeError> {
        let arg_values: Result<Vec<Value>, _> =
            args.iter().map(|e| self.eval_expr(e)).collect();
        let arg_values = arg_values?;

        let name_lower = crate::utils::normalize_identifier(name);
        if let Some(func) = self.context.functions.get(&name_lower).cloned() {
            self.context.push_scope();

            for (i, param_name) in func.params.iter().enumerate() {
                let value = if i < arg_values.len() {
                    arg_values[i].clone()
                } else {
                    Value::Empty
                };
                self.context.define_var(param_name.clone(), value);
            }

            let result = self.exec_block(&func.body);
            self.context.pop_scope();
            result?;
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
}
