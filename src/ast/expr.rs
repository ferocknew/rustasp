//! 表达式定义

use super::BinaryOp;
use super::UnaryOp;
use serde::{Deserialize, Serialize};

/// 表达式
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum Expr {
    /// 数字字面量
    Number(f64),
    /// 字符串字面量
    String(String),
    /// 布尔字面量
    Boolean(bool),
    /// Nothing
    Nothing,
    /// Empty
    Empty,
    /// Null
    Null,
    /// 变量引用
    Variable(String),
    /// 二元运算
    Binary {
        left: Box<Expr>,
        op: BinaryOp,
        right: Box<Expr>,
    },
    /// 一元运算
    Unary { op: UnaryOp, operand: Box<Expr> },
    /// 函数调用
    Call { name: String, args: Vec<Expr> },
    /// 方法调用
    Method {
        object: Box<Expr>,
        method: String,
        args: Vec<Expr>,
    },
    /// 属性访问
    Property { object: Box<Expr>, property: String },
    /// 索引访问
    Index { object: Box<Expr>, index: Box<Expr> },
    /// 数组字面量
    Array(Vec<Expr>),
    /// New 表达式
    New(String),
}

impl Expr {
    /// 创建数字表达式
    pub fn number(n: f64) -> Self {
        Expr::Number(n)
    }

    /// 创建字符串表达式
    pub fn string(s: impl Into<String>) -> Self {
        Expr::String(s.into())
    }

    /// 创建布尔表达式
    pub fn boolean(b: bool) -> Self {
        Expr::Boolean(b)
    }

    /// 创建变量表达式
    pub fn variable(name: impl Into<String>) -> Self {
        Expr::Variable(name.into())
    }

    /// 创建二元运算表达式
    pub fn binary(left: Expr, op: BinaryOp, right: Expr) -> Self {
        Expr::Binary {
            left: Box::new(left),
            op,
            right: Box::new(right),
        }
    }

    /// 创建函数调用表达式
    pub fn call(name: impl Into<String>, args: Vec<Expr>) -> Self {
        Expr::Call {
            name: name.into(),
            args,
        }
    }

    /// 创建属性访问表达式
    pub fn property(object: Expr, property: impl Into<String>) -> Self {
        Expr::Property {
            object: Box::new(object),
            property: property.into(),
        }
    }

    /// 创建方法调用表达式
    pub fn method(object: Expr, method: impl Into<String>, args: Vec<Expr>) -> Self {
        Expr::Method {
            object: Box::new(object),
            method: method.into(),
            args,
        }
    }
}
