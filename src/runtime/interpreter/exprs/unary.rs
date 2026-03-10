//! 一元运算表达式求值

use crate::ast::{Expr, UnaryOp};
use crate::runtime::{Value, ValueConversion};

use super::super::Interpreter;

impl Interpreter {
    /// 执行一元运算
    pub fn eval_unary(
        &mut self,
        op: UnaryOp,
        operand: &Expr,
    ) -> Result<Value, crate::runtime::RuntimeError> {
        let val = self.eval_expr(operand)?;

        match op {
            UnaryOp::Neg => Ok(Value::Number(-val.to_number())),
            UnaryOp::Not => Ok(Value::Boolean(!val.is_truthy())),
        }
    }
}
