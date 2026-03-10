//! 解释器模块
//!
//! 将解释器拆分为多个子模块以提高可维护性

pub mod exprs;
mod stmts;

use crate::ast::Program;
use crate::runtime::builtins::TokenRegistry;
use crate::runtime::{Context, RuntimeError, Value};

/// 解释器
pub struct Interpreter {
    pub(crate) context: Context,
    /// 内置函数注册表（缓存，避免每次调用都创建）
    pub(crate) builtin_registry: TokenRegistry,
}

impl Interpreter {
    /// 创建新解释器
    pub fn new() -> Self {
        Interpreter {
            context: Context::new(),
            builtin_registry: TokenRegistry::new(),
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
        // 预处理阶段：扫描并注册所有的类和函数定义
        self.preprocess_declarations(&program.statements)?;

        // 执行阶段：按顺序执行语句
        let mut result = Value::Empty;
        for stmt in &program.statements {
            // 跳过类和函数声明（已在预处理阶段注册）
            match stmt {
                crate::ast::Stmt::Class { .. }
                | crate::ast::Stmt::Function { .. }
                | crate::ast::Stmt::Sub { .. } => continue,
                _ => {
                    result = self.eval_stmt(stmt)?;
                    if self.context.should_exit {
                        break;
                    }
                }
            }
        }
        Ok(result)
    }

    /// 预处理阶段：扫描并注册所有的类和函数定义
    fn preprocess_declarations(
        &mut self,
        statements: &[crate::ast::Stmt],
    ) -> Result<Value, RuntimeError> {
        use crate::ast::Stmt;

        for stmt in statements {
            match stmt {
                Stmt::Class { name, members } => {
                    // 注册类定义
                    self.register_class(name, members)?;
                }
                Stmt::Function { name, params, body } | Stmt::Sub { name, params, body } => {
                    // 注册函数/过程定义
                    self.register_function(name, params, body)?;
                }
                _ => {
                    // 其他语句忽略，只处理声明
                }
            }
        }

        Ok(Value::Empty)
    }
}

impl Default for Interpreter {
    fn default() -> Self {
        Self::new()
    }
}
