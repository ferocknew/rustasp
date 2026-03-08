//! 表达式求值模块
//!
//! 处理各种 VBScript 表达式的求值逻辑

mod binary;
mod call;
mod index;
mod method;
mod new;
mod property;
mod unary;

use crate::ast::Expr;
use crate::runtime::{RuntimeError, Value};
use std::sync::{Arc, Mutex};

use super::Interpreter;

/// 表达式求值器
impl Interpreter {
    /// 求值表达式
    pub fn eval_expr(&mut self, expr: &Expr) -> Result<Value, RuntimeError> {
        match expr {
            // 字面量
            Expr::Number(n) => Ok(Value::Number(*n)),
            Expr::String(s) => Ok(Value::String(s.clone())),
            Expr::Boolean(b) => Ok(Value::Boolean(*b)),
            Expr::Date(date_str) => Ok(Value::String(date_str.clone())),
            Expr::Nothing => Ok(Value::Nothing),
            Expr::Empty => Ok(Value::Empty),
            Expr::Null => Ok(Value::Null),

            // 变量
            Expr::Variable(name) => self.eval_variable(name),

            // 运算符
            Expr::Binary { left, op, right } => self.eval_binary(left, *op, right),
            Expr::Unary { op, operand } => self.eval_unary(*op, operand),

            // 调用和访问
            Expr::Call { name, args } => self.eval_call_expr(name, args),
            Expr::Property { object, property } => self.eval_property(object, property),
            Expr::Method { object, method, args } => self.eval_method(object, method, args),
            Expr::Index { object, index } => self.eval_index(object, index),

            // 其他
            Expr::Array(elements) => self.eval_array(elements),
            Expr::New(class_name) => self.eval_new(class_name),
        }
    }

    /// 求值变量
    fn eval_variable(&mut self, name: &str) -> Result<Value, RuntimeError> {
        // 首先尝试从上下文中获取变量
        if let Some(value) = self.context.get_var(name).cloned() {
            return Ok(value);
        }

        // 如果没有找到变量，检查是否是内置函数（无参数调用）
        if let Some(result) = self.call_builtin(name, &[]) {
            return result;
        }

        // 检查是否是用户定义的函数（无参数调用）
        if self.context.get_function(name).is_some() {
            return self.eval_call_expr(name, &[]);
        }

        Ok(Value::Empty)
    }

    /// 求值数组字面量
    fn eval_array(&mut self, elements: &[Expr]) -> Result<Value, RuntimeError> {
        let values: Result<Vec<Value>, _> =
            elements.iter().map(|e| self.eval_expr(e)).collect();
        Ok(Value::Array(Arc::new(Mutex::new(values?))))
    }
}
