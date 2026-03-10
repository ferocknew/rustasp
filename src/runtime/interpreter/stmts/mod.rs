//! 语句执行模块
//!
//! 处理各种 VBScript 语句的执行逻辑

mod decl_stmt;
mod assign_stmt;
mod control_stmt;

use crate::ast::{Param, Stmt};
use crate::runtime::{ClassDef, Function, RuntimeError, Value, VbsClass, ErrorMode, vb_error};
use std::rc::Rc;

use super::Interpreter;

/// 语句执行器
impl Interpreter {
    /// 执行语句（调度）- 支持错误处理
    pub fn eval_stmt(&mut self, stmt: &Stmt) -> Result<Value, RuntimeError> {
        // 获取当前错误模式
        let error_mode = self.context.current_scope().get_error_mode();

        // 执行语句
        let result = self.eval_stmt_inner(stmt);

        // 根据错误模式处理错误
        match result {
            Ok(value) => Ok(value),
            Err(e) => {
                match error_mode {
                    ErrorMode::Stop => Err(e),
                    ErrorMode::ResumeNext => {
                        // 记录错误到 Err 对象
                        let (number, description) = self.extract_error_info(&e);
                        self.context.err.set(number, description);
                        // 继续执行
                        Ok(Value::Empty)
                    }
                }
            }
        }
    }

    /// 执行语句（内部实现）- 不处理错误
    fn eval_stmt_inner(&mut self, stmt: &Stmt) -> Result<Value, RuntimeError> {
        match stmt {
            // 错误处理语句
            Stmt::OnErrorResumeNext => {
                self.context.current_scope_mut().set_error_mode(ErrorMode::ResumeNext);
                Ok(Value::Empty)
            }
            Stmt::OnErrorGoto0 => {
                self.context.current_scope_mut().set_error_mode(ErrorMode::Stop);
                Ok(Value::Empty)
            }

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

            // Do 循环
            Stmt::DoWhile { cond, body } => self.eval_do_while(cond, body),
            Stmt::DoUntil { cond, body } => self.eval_do_until(cond, body),
            Stmt::DoLoopWhile { body, cond } => self.eval_do_loop_while(body, cond),
            Stmt::DoLoopUntil { body, cond } => self.eval_do_loop_until(body, cond),
            Stmt::DoLoop { body } => self.eval_do_loop(body),

            // 函数相关
            Stmt::Sub { name, params, body } | Stmt::Function { name, params, body } => {
                self.register_function(name, params, body)
            }

            // 类定义
            Stmt::Class { name, members } => self.register_class(name, members),

            Stmt::Call { name, args } => self.eval_call(name, args),
            Stmt::ExitFor => self.eval_exit_for(),
            Stmt::ExitDo => self.eval_exit_do(),
            Stmt::ExitFunction => self.eval_exit_function(),
            Stmt::ExitSub => self.eval_exit_sub(),
            Stmt::ExitProperty => self.eval_exit_property(),

            // 其他语句
            Stmt::OptionExplicit => {
                // Option Explicit: 要求所有变量必须先声明
                // 当前实现暂时忽略，不强制检查
                Ok(Value::Empty)
            }
            Stmt::Expr(expr) => self.eval_expr(expr),

            // 动态执行语句
            Stmt::Execute(expr) => self.eval_execute(expr),
            Stmt::ExecuteGlobal(expr) => self.eval_execute_global(expr),

            _ => Err(RuntimeError::Generic(format!("Unimplemented: {:?}", stmt))),
        }
    }

    /// 从 RuntimeError 中提取错误信息
    fn extract_error_info(&self, error: &RuntimeError) -> (i32, String) {
        match error {
            RuntimeError::DivisionByZero => {
                (vb_error::DIVISION_BY_ZERO, "Division by zero".to_string())
            }
            RuntimeError::TypeMismatch(msg) => {
                (vb_error::TYPE_MISMATCH, format!("Type mismatch: {}", msg))
            }
            RuntimeError::ObjectRequired => {
                (vb_error::OBJECT_REQUIRED, "Object required".to_string())
            }
            RuntimeError::UndefinedFunction(name) => {
                (vb_error::UNDEFINED_FUNCTION, format!("Undefined function: {}", name))
            }
            RuntimeError::IndexOutOfBounds(_) => {
                (vb_error::SUBSCRIPT_OUT_OF_RANGE, "Subscript out of range".to_string())
            }
            RuntimeError::CreateObjectFailed(msg) => {
                (vb_error::CANT_CREATE_OBJECT, format!("Server.CreateObject: {}", msg))
            }
            _ => (0, format!("{:?}", error)),
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
    pub(crate) fn register_function(&mut self, name: &str, params: &[Param], body: &[Stmt]) -> Result<Value, RuntimeError> {
        self.context.functions.insert(
            crate::utils::normalize_identifier(name),
            Function {
                name: name.to_string(),
                params: params.to_vec(),
                body: body.to_vec(),
            },
        );
        Ok(Value::Empty)
    }

    /// 注册类定义（预编译 VbsClass 并缓存）
    pub(crate) fn register_class(&mut self, name: &str, members: &[crate::ast::ClassMember]) -> Result<Value, RuntimeError> {
        let normalized_name = crate::utils::normalize_identifier(name);
        
        // 预编译 VbsClass（只构建一次）
        let vbs_class = VbsClass::from_ast(name.to_string(), members.to_vec());
        
        // 缓存编译后的类
        self.context.classes.insert(
            normalized_name.clone(),
            Rc::new(vbs_class),
        );
        
        // 保留原始定义用于调试
        self.context.class_defs.insert(
            normalized_name,
            ClassDef {
                name: name.to_string(),
                members: members.to_vec(),
            },
        );

        Ok(Value::Empty)
    }

    /// 执行 Execute 语句
    /// Execute 在当前作用域中动态执行 VBScript 代码
    fn eval_execute(&mut self, expr: &crate::ast::Expr) -> Result<Value, RuntimeError> {
        use crate::parser::Parser;
        use crate::parser::lexer::tokenize;

        // 1. 计算表达式得到字符串
        let code_value = self.eval_expr(expr)?;
        let code_str = match &code_value {
            Value::String(s) => s.clone(),
            _ => {
                return Err(RuntimeError::TypeMismatch(
                    "Execute argument must be a string".to_string()
                ));
            }
        };

        // 2. 解析代码字符串
        let spanned_tokens = match tokenize(&code_str) {
            Ok(tokens) => tokens,
            Err(e) => {
                return Err(RuntimeError::Generic(format!(
                    "Execute tokenization error: {}", e
                )));
            }
        };

        let mut parser = Parser::new(spanned_tokens);

        let mut stmts = Vec::new();
        loop {
            // 跳过冒号分隔符
            while matches!(parser.peek(), crate::parser::lexer::Token::Colon) {
                parser.advance();
            }

            match parser.parse_stmt() {
                Ok(Some(stmt)) => stmts.push(stmt),
                Ok(None) => break, // 解析完成
                Err(e) => {
                    return Err(RuntimeError::Generic(format!(
                        "Execute parse error: {}", e
                    )));
                }
            }
        }

        // 3. 预处理：注册类和函数定义
        let mut decl_only_stmts = Vec::new();
        let mut exec_stmts = Vec::new();
        for stmt in &stmts {
            match stmt {
                Stmt::Class { .. } | Stmt::Function { .. } | Stmt::Sub { .. } => {
                    decl_only_stmts.push(stmt);
                }
                _ => {
                    exec_stmts.push(stmt);
                }
            }
        }

        // 先注册所有声明
        for stmt in &decl_only_stmts {
            self.eval_stmt(stmt)?;
        }

        // 然后执行其他语句
        for stmt in &exec_stmts {
            self.eval_stmt(stmt)?;
        }

        Ok(Value::Empty)
    }

    /// 执行 ExecuteGlobal 语句
    /// ExecuteGlobal 在全局作用域中动态执行 VBScript 代码
    fn eval_execute_global(&mut self, expr: &crate::ast::Expr) -> Result<Value, RuntimeError> {
        use crate::parser::Parser;
        use crate::parser::lexer::tokenize;

        // 1. 计算表达式得到字符串
        let code_value = self.eval_expr(expr)?;
        let code_str = match &code_value {
            Value::String(s) => s.clone(),
            _ => {
                return Err(RuntimeError::TypeMismatch(
                    "ExecuteGlobal argument must be a string".to_string()
                ));
            }
        };

        // 2. 解析代码字符串
        let spanned_tokens = match tokenize(&code_str) {
            Ok(tokens) => tokens,
            Err(e) => {
                return Err(RuntimeError::Generic(format!(
                    "ExecuteGlobal tokenization error: {}", e
                )));
            }
        };

        let mut parser = Parser::new(spanned_tokens);

        let mut stmts = Vec::new();
        loop {
            // 跳过冒号分隔符
            while matches!(parser.peek(), crate::parser::lexer::Token::Colon) {
                parser.advance();
            }

            match parser.parse_stmt() {
                Ok(Some(stmt)) => stmts.push(stmt),
                Ok(None) => break, // 解析完成
                Err(e) => {
                    return Err(RuntimeError::Generic(format!(
                        "ExecuteGlobal parse error: {}", e
                    )));
                }
            }
        }

        // 3. 保存当前作用域深度
        let current_scope_depth = self.context.scope_stack.len();

        // 4. 切换到全局作用域（第0层）
        while self.context.scope_stack.len() > 1 {
            self.context.scope_stack.pop();
        }

        // 5. 在全局作用域执行语句
        let result = self.exec_block(&stmts);

        // 6. 恢复原来的作用域深度
        while self.context.scope_stack.len() < current_scope_depth {
            self.context.push_scope();
        }

        result
    }
}
