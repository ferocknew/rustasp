//! 语句执行模块
//!
//! 处理各种 VBScript 语句的执行逻辑

use crate::ast::{BinaryOp, Expr, IfBranch, Param, Stmt};
use crate::runtime::{Function, RuntimeError, Value, ValueCompare, ValueConversion};
use crate::utils::normalize_identifier;

use super::Interpreter;

/// 语句执行器
impl Interpreter {
    /// 执行语句（调度）
    pub fn eval_stmt(&mut self, stmt: &Stmt) -> Result<Value, RuntimeError> {
        match stmt {
            Stmt::Dim { name, init, .. } => self.eval_dim(name, init.as_ref()),
            Stmt::Const { name, value } => self.eval_const(name, value),
            Stmt::Assignment { target, value } => self.eval_assignment(target, value),
            Stmt::Set { target, value } => self.eval_set(target, value),
            Stmt::If {
                branches,
                else_block,
            } => self.eval_if(branches, else_block),
            Stmt::For {
                var,
                start,
                end,
                step,
                body,
            } => self.eval_for(var, start, end, step.as_ref(), body),
            Stmt::While { cond, body } => self.eval_while(cond, body),
            Stmt::Select {
                expr,
                cases,
                else_block,
            } => self.eval_select(expr, cases, else_block),
            Stmt::Sub { name, params, body } => self.eval_sub(name, params, body),
            Stmt::Function { name, params, body } => self.eval_function(name, params, body),
            Stmt::Call { name, args } => self.eval_call(name, args),
            Stmt::ExitFor => self.eval_exit_for(),
            Stmt::ExitFunction => self.eval_exit_function(),
            Stmt::ExitSub => self.eval_exit_sub(),
            Stmt::OptionExplicit => {
                // Option Explicit: 要求所有变量必须先声明
                // 当前实现暂时忽略，不强制检查
                Ok(Value::Empty)
            }
            Stmt::Expr(expr) => self.eval_expr(expr),
            _ => Err(RuntimeError::Generic(format!("Unimplemented: {:?}", stmt))),
        }
    }

    /// 执行 Dim 语句
    fn eval_dim(&mut self, name: &str, init: Option<&Expr>) -> Result<Value, RuntimeError> {
        let value = if let Some(expr) = init {
            self.eval_expr(expr)?
        } else {
            Value::Empty
        };
        self.context.define_var(name.to_string(), value);
        Ok(Value::Empty)
    }

    /// 执行 Const 语句
    fn eval_const(&mut self, name: &str, value: &Expr) -> Result<Value, RuntimeError> {
        let val = self.eval_expr(value)?;
        self.context.define_var(name.to_string(), val);
        Ok(Value::Empty)
    }

    /// 执行赋值语句
    fn eval_assignment(&mut self, target: &Expr, value: &Expr) -> Result<Value, RuntimeError> {
        let val = self.eval_expr(value)?;
        match target {
            Expr::Variable(name) => {
                self.context.set_var(name.clone(), val);
                Ok(Value::Empty)
            }
            Expr::Index { object, index } => {
                let idx = self.eval_expr(index)?;
                match object.as_ref() {
                    Expr::Variable(name) => {
                        if let Some(Value::Array(arr)) = self.context.get_var(name).cloned() {
                            let mut arr = arr;
                            if let Value::Number(i) = idx {
                                let i = i as usize;
                                if i < arr.len() {
                                    arr[i] = val;
                                    self.context.set_var(name.clone(), Value::Array(arr));
                                    return Ok(Value::Empty);
                                }
                            }
                        }
                        Err(RuntimeError::InvalidAssignment)
                    }
                    _ => Err(RuntimeError::InvalidAssignment),
                }
            }
            _ => Err(RuntimeError::InvalidAssignment),
        }
    }

    /// 执行 Set 语句
    fn eval_set(&mut self, target: &Expr, value: &Expr) -> Result<Value, RuntimeError> {
        self.eval_assignment(target, value)
    }

    /// 执行 If 语句
    fn eval_if(
        &mut self,
        branches: &[IfBranch],
        else_block: &Option<Vec<Stmt>>,
    ) -> Result<Value, RuntimeError> {
        for branch in branches {
            let cond = self.eval_expr(&branch.cond)?;
            if cond.is_truthy() {
                for stmt in &branch.body {
                    self.eval_stmt(stmt)?;
                }
                return Ok(Value::Empty);
            }
        }
        if let Some(else_stmts) = else_block {
            for stmt in else_stmts {
                self.eval_stmt(stmt)?;
            }
        }
        Ok(Value::Empty)
    }

    /// 执行 For 循环
    fn eval_for(
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
            for stmt in body {
                self.eval_stmt(stmt)?;
            }
            i += step_val;
        }

        Ok(Value::Empty)
    }

    /// 执行 While 循环
    fn eval_while(&mut self, cond: &Expr, body: &[Stmt]) -> Result<Value, RuntimeError> {
        while self.eval_expr(cond)?.is_truthy() {
            for stmt in body {
                self.eval_stmt(stmt)?;
            }
        }
        Ok(Value::Empty)
    }

    /// 执行 Select Case 语句
    fn eval_select(
        &mut self,
        expr: &Expr,
        cases: &[crate::ast::CaseClause],
        else_block: &Option<Vec<Stmt>>,
    ) -> Result<Value, RuntimeError> {
        // 计算表达式的值
        let select_value = self.eval_expr(expr)?;

        // 遍历所有 Case 分支
        for case in cases {
            if let Some(values) = &case.values {
                // 检查是否有匹配的值
                for value_expr in values {
                    let case_value = self.eval_expr(value_expr)?;
                    // 使用 compare 方法进行相等比较
                    let result = select_value.compare(BinaryOp::Eq, &case_value);
                    if let Value::Boolean(true) = result {
                        // 执行匹配的 Case body
                        for stmt in &case.body {
                            self.eval_stmt(stmt)?;
                        }
                        return Ok(Value::Empty);
                    }
                }
            }
        }

        // 如果没有匹配的 Case，执行 Else 块
        if let Some(else_stmts) = else_block {
            for stmt in else_stmts {
                self.eval_stmt(stmt)?;
            }
        }

        Ok(Value::Empty)
    }

    /// 注册 Sub
    fn eval_sub(
        &mut self,
        name: &str,
        params: &[Param],
        body: &[Stmt],
    ) -> Result<Value, RuntimeError> {
        self.context.functions.insert(
            normalize_identifier(name),
            Function {
                name: name.to_string(),
                params: params.iter().map(|p| p.name.clone()).collect(),
                body: body.to_vec(),
            },
        );
        Ok(Value::Empty)
    }

    /// 注册 Function
    fn eval_function(
        &mut self,
        name: &str,
        params: &[Param],
        body: &[Stmt],
    ) -> Result<Value, RuntimeError> {
        self.context.functions.insert(
            normalize_identifier(name),
            Function {
                name: name.to_string(),
                params: params.iter().map(|p| p.name.clone()).collect(),
                body: body.to_vec(),
            },
        );
        Ok(Value::Empty)
    }

    /// 执行 Call 语句
    fn eval_call(&mut self, name: &str, args: &[Expr]) -> Result<Value, RuntimeError> {
        let _args: Result<Vec<Value>, _> = args.iter().map(|e| self.eval_expr(e)).collect();
        // 简化：直接调用函数
        let name_lower = normalize_identifier(name);
        if let Some(func) = self.context.functions.get(&name_lower).cloned() {
            self.context.push_scope();
            for stmt in &func.body {
                self.eval_stmt(stmt)?;
            }
            self.context.pop_scope();
        }
        Ok(Value::Empty)
    }

    /// Exit For
    fn eval_exit_for(&mut self) -> Result<Value, RuntimeError> {
        // TODO: 实现循环退出
        Ok(Value::Empty)
    }

    /// Exit Function
    fn eval_exit_function(&mut self) -> Result<Value, RuntimeError> {
        self.context.should_exit = true;
        Ok(Value::Empty)
    }

    /// Exit Sub
    fn eval_exit_sub(&mut self) -> Result<Value, RuntimeError> {
        self.context.should_exit = true;
        Ok(Value::Empty)
    }
}
