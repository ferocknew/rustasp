//! 解释器 - 调度语句执行

use super::value::{ValueCompare, ValueConversion, ValueOps};
use super::{Context, RuntimeError, Value};
use crate::ast::{BinaryOp, Expr, Program, Stmt, UnaryOp};

/// 解释器
pub struct Interpreter {
    context: Context,
}

impl Interpreter {
    /// 创建新解释器
    pub fn new() -> Self {
        Interpreter {
            context: Context::new(),
        }
    }

    /// 获取上下文
    pub fn context(&self) -> &Context {
        &self.context
    }

    /// 获取可变上下文
    pub fn context_mut(&mut self) -> &mut Context {
        &mut self.context
    }

    /// 执行程序
    pub fn execute(&mut self, program: &Program) -> Result<Value, RuntimeError> {
        let mut result = Value::Empty;
        for stmt in &program.statements {
            result = self.eval_stmt(stmt)?;
            if self.context.should_exit {
                break;
            }
        }
        Ok(result)
    }

    /// 执行语句（调度）
    fn eval_stmt(&mut self, stmt: &Stmt) -> Result<Value, RuntimeError> {
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
            Stmt::Sub { name, params, body } => self.eval_sub(name, params, body),
            Stmt::Function { name, params, body } => self.eval_function(name, params, body),
            Stmt::Call { name, args } => self.eval_call(name, args),
            Stmt::ExitFor => self.eval_exit_for(),
            Stmt::ExitFunction => self.eval_exit_function(),
            Stmt::ExitSub => self.eval_exit_sub(),
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
        branches: &[crate::ast::IfBranch],
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
            .map(|e| self.eval_expr(e).map(|v| v.to_number()))
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

    /// 注册 Sub
    fn eval_sub(
        &mut self,
        name: &str,
        params: &[crate::ast::Param],
        body: &[Stmt],
    ) -> Result<Value, RuntimeError> {
        self.context.functions.insert(
            name.to_lowercase(),
            super::Function {
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
        params: &[crate::ast::Param],
        body: &[Stmt],
    ) -> Result<Value, RuntimeError> {
        self.context.functions.insert(
            name.to_lowercase(),
            super::Function {
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
        let name_lower = name.to_lowercase();
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

    /// 求值表达式
    fn eval_expr(&mut self, expr: &Expr) -> Result<Value, RuntimeError> {
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
            Expr::Call { name, args } => {
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
            Expr::Array(elements) => {
                let values: Result<Vec<Value>, _> =
                    elements.iter().map(|e| self.eval_expr(e)).collect();
                Ok(Value::Array(values?))
            }
            _ => Err(RuntimeError::Generic(format!(
                "Unimplemented expr: {:?}",
                expr
            ))),
        }
    }
}

impl Default for Interpreter {
    fn default() -> Self {
        Self::new()
    }
}
