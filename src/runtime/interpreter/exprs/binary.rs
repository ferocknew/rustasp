//! 二元运算表达式求值

use crate::ast::{BinaryOp, Expr};
use crate::runtime::{RuntimeError, Value, ValueCompare, ValueOps};

use super::super::Interpreter;

impl Interpreter {
    /// 执行二元运算
    pub fn eval_binary(
        &mut self,
        left: &Expr,
        op: BinaryOp,
        right: &Expr,
    ) -> Result<Value, RuntimeError> {
        let left_val = self.eval_expr(left)?;
        let right_val = self.eval_expr(right)?;

        match op {
            // 比较运算符
            BinaryOp::Eq
            | BinaryOp::Ne
            | BinaryOp::Lt
            | BinaryOp::Le
            | BinaryOp::Gt
            | BinaryOp::Ge
            | BinaryOp::And
            | BinaryOp::Or
            | BinaryOp::Xor
            | BinaryOp::Is => Ok(left_val.compare(op, &right_val)),
            // 其他运算符
            _ => left_val.binary_op(op, &right_val),
        }
    }
}
