//! 控制流语句执行
//!
//! 处理 If 和 Select Case 语句

use crate::ast::{BinaryOp, Expr, IfBranch, Stmt};
use crate::runtime::{RuntimeError, Value, ValueCompare};
use super::Interpreter;

impl Interpreter {
    /// 执行 If 语句
    pub fn eval_if(
        &mut self,
        branches: &[IfBranch],
        else_block: &Option<Vec<Stmt>>,
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

    /// 执行 Select Case 语句
    pub fn eval_select(
        &mut self,
        expr: &Expr,
        cases: &[crate::ast::CaseClause],
        else_block: &Option<Vec<Stmt>>,
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

    /// 执行语句块
    pub fn exec_block(&mut self, stmts: &[Stmt]) -> Result<Value, RuntimeError> {
        for stmt in stmts {
            self.eval_stmt(stmt)?;
        }
        Ok(Value::Empty)
    }
}
