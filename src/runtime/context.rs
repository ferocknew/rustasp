//! 执行上下文

use super::Scope;
use super::Value;
use std::collections::HashMap;

/// 函数定义
#[derive(Debug, Clone)]
pub struct Function {
    pub name: String,
    pub params: Vec<String>,
    pub body: Vec<crate::ast::Stmt>,
}

/// 执行上下文
#[derive(Debug)]
pub struct Context {
    /// 全局作用域
    pub global: Scope,
    /// 当前作用域栈
    pub scope_stack: Vec<Scope>,
    /// 函数表
    pub functions: HashMap<String, Function>,
    /// 类表
    pub classes: HashMap<String, ClassDef>,
    /// 输出缓冲区
    pub output: String,
    /// 是否应该退出
    pub should_exit: bool,
    /// 请求参数（用于 Request 对象）
    pub request_data: HashMap<String, String>,
}

/// 类定义
#[derive(Debug, Clone)]
pub struct ClassDef {
    pub name: String,
    pub members: Vec<crate::ast::ClassMember>,
}

impl Context {
    /// 创建新上下文
    pub fn new() -> Self {
        Context {
            global: Scope::new(),
            scope_stack: Vec::new(),
            functions: HashMap::new(),
            classes: HashMap::new(),
            output: String::new(),
            should_exit: false,
            request_data: HashMap::new(),
        }
    }

    /// 获取当前作用域
    pub fn current_scope(&self) -> &Scope {
        self.scope_stack.last().unwrap_or(&self.global)
    }

    /// 获取当前作用域（可变）
    pub fn current_scope_mut(&mut self) -> &mut Scope {
        self.scope_stack.last_mut().unwrap_or(&mut self.global)
    }

    /// 获取变量
    pub fn get_var(&self, name: &str) -> Option<&Value> {
        if let Some(scope) = self.scope_stack.last() {
            if let Some(v) = scope.get(name) {
                return Some(v);
            }
        }
        self.global.get(name)
    }

    /// 设置变量
    pub fn set_var(&mut self, name: String, value: Value) {
        let name_lower = name.to_lowercase();

        // 先在当前作用域查找
        if let Some(scope) = self.scope_stack.last_mut() {
            if scope.contains(&name_lower) {
                scope.set(name, value);
                return;
            }
        }

        // 再在全局作用域查找
        if self.global.contains(&name_lower) {
            self.global.set(name, value);
            return;
        }

        // 未找到则在当前作用域创建
        if let Some(scope) = self.scope_stack.last_mut() {
            scope.set(name, value);
        } else {
            self.global.set(name, value);
        }
    }

    /// 定义变量
    pub fn define_var(&mut self, name: String, value: Value) {
        if let Some(scope) = self.scope_stack.last_mut() {
            scope.define(name, value);
        } else {
            self.global.define(name, value);
        }
    }

    /// 压入新作用域
    pub fn push_scope(&mut self) {
        let parent = self
            .scope_stack
            .last()
            .cloned()
            .unwrap_or_else(|| self.global.clone());
        self.scope_stack.push(Scope::with_parent(parent));
    }

    /// 弹出作用域
    pub fn pop_scope(&mut self) {
        self.scope_stack.pop();
    }

    /// 写入输出
    pub fn write(&mut self, s: &str) {
        self.output.push_str(s);
    }

    /// 获取输出
    pub fn get_output(&self) -> &str {
        &self.output
    }

    /// 清空输出
    pub fn clear_output(&mut self) {
        self.output.clear();
    }

    /// 设置请求参数
    pub fn set_request_data(&mut self, data: HashMap<String, String>) {
        self.request_data = data;
    }

    /// 获取请求参数
    pub fn get_request_param(&self, key: &str) -> Option<&String> {
        self.request_data.get(&key.to_lowercase())
    }
}

impl Default for Context {
    fn default() -> Self {
        Self::new()
    }
}
