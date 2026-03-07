//! 语句执行模块
//!
//! 处理各种 VBScript 语句的执行逻辑

mod decl_stmt;
mod assign_stmt;
mod control_stmt;

use crate::ast::{Param, Stmt};
use crate::runtime::{Function, RuntimeError, Value};

use super::Interpreter;

/// 语句执行器
impl Interpreter {
    /// 执行语句（调度）
    pub fn eval_stmt(&mut self, stmt: &Stmt) -> Result<Value, RuntimeError> {
        match stmt {
            // 声明语句
            Stmt::Dim { name, init, is_array, sizes } => {
                self.eval_dim(name, init.as_ref(), *is_array, sizes)
            }
            Stmt::Const { name, value } => self.eval_const(name, value),
            Stmt::ReDim { name, sizes, preserve } => self.eval_redim(name, sizes, *preserve),

            // 赋值语句
            Stmt::Assignment { target, value } => self.eval_assignment(target, value),
            Stmt::Set { target, value } => self.eval_set(target, value),

            // 控制流语句
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
            Stmt::ForEach { var, collection, body } => self.eval_for_each(var, collection, body),
            Stmt::Select {
                expr,
                cases,
                else_block,
            } => self.eval_select(expr, cases, else_block),

            // 函数相关
            Stmt::Sub { name, params, body } | Stmt::Function { name, params, body } => {
                self.register_function(name, params, body)
            }
            Stmt::Call { name, args } => self.eval_call(name, args),
            Stmt::ExitFor => self.eval_exit_for(),
            Stmt::ExitFunction | Stmt::ExitSub => self.eval_exit(),

            // 其他语句
            Stmt::OptionExplicit => {
                // Option Explicit: 要求所有变量必须先声明
                // 当前实现暂时忽略，不强制检查
                Ok(Value::Empty)
            }
            Stmt::Expr(expr) => self.eval_expr(expr),
            _ => Err(RuntimeError::Generic(format!("Unimplemented: {:?}", stmt))),
        }
    }

    /// 执行语句块
    fn exec_block(&mut self, stmts: &[Stmt]) -> Result<Value, RuntimeError> {
        for stmt in stmts {
            self.eval_stmt(stmt)?;
        }
        Ok(Value::Empty)
    }

    /// 注册函数(Sub 或 Function)
    fn register_function(&mut self, name: &str, params: &[Param], body: &[Stmt]) -> Result<Value, RuntimeError> {
        self.context.functions.insert(
            crate::utils::normalize_identifier(name),
            Function {
                name: name.to_string(),
                params: params.iter().map(|p| p.name.clone()).collect(),
                body: body.to_vec(),
            },
        );
        Ok(Value::Empty)
    }
}
