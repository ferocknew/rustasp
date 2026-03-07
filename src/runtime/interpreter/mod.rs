//! 解释器模块
//!
//! 将解释器拆分为多个子模块以提高可维护性

mod builtins;
mod builtin_tokens;
mod exprs;
mod stmts;

use crate::ast::Program;
use crate::runtime::{Context, RuntimeError, Value};

/// 解释器
pub struct Interpreter {
    pub(crate) context: Context,
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
}

impl Default for Interpreter {
    fn default() -> Self {
        Self::new()
    }
}
