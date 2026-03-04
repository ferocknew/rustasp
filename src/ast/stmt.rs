//! 语句定义

use super::Expr;
use serde::{Deserialize, Serialize};

/// 语句
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum Stmt {
    /// Dim 声明
    Dim {
        name: String,
        init: Option<Expr>,
        is_array: bool,
        sizes: Vec<Expr>,
    },
    /// Const 常量声明
    Const { name: String, value: Expr },
    /// 赋值语句
    Assignment { target: Expr, value: Expr },
    /// Set 赋值
    Set { target: Expr, value: Expr },
    /// If 语句
    If {
        branches: Vec<IfBranch>,
        else_block: Option<Vec<Stmt>>,
    },
    /// For...Next 循环
    For {
        var: String,
        start: Expr,
        end: Expr,
        step: Option<Expr>,
        body: Vec<Stmt>,
    },
    /// For Each 循环
    ForEach {
        var: String,
        collection: Expr,
        body: Vec<Stmt>,
    },
    /// Do While 循环
    DoWhile { cond: Expr, body: Vec<Stmt> },
    /// Do Until 循环
    DoUntil { cond: Expr, body: Vec<Stmt> },
    /// Do...Loop While
    DoLoopWhile { body: Vec<Stmt>, cond: Expr },
    /// Do...Loop Until
    DoLoopUntil { body: Vec<Stmt>, cond: Expr },
    /// While...Wend 循环
    While { cond: Expr, body: Vec<Stmt> },
    /// Exit For
    ExitFor,
    /// Exit Do
    ExitDo,
    /// Exit Function
    ExitFunction,
    /// Exit Sub
    ExitSub,
    /// Exit Property
    ExitProperty,
    /// Sub 定义
    Sub {
        name: String,
        params: Vec<Param>,
        body: Vec<Stmt>,
    },
    /// Function 定义
    Function {
        name: String,
        params: Vec<Param>,
        body: Vec<Stmt>,
    },
    /// Call 语句
    Call { name: String, args: Vec<Expr> },
    /// 类定义
    Class {
        name: String,
        members: Vec<ClassMember>,
    },
    /// Property Get
    PropertyGet {
        name: String,
        params: Vec<Param>,
        body: Vec<Stmt>,
    },
    /// Property Let
    PropertyLet {
        name: String,
        params: Vec<Param>,
        body: Vec<Stmt>,
    },
    /// Property Set
    PropertySet {
        name: String,
        params: Vec<Param>,
        body: Vec<Stmt>,
    },
    /// ReDim
    ReDim {
        name: String,
        sizes: Vec<Expr>,
        preserve: bool,
    },
    /// Erase
    Erase(String),
    /// Execute
    Execute(Expr),
    /// ExecuteGlobal
    ExecuteGlobal(Expr),
    /// Eval
    Eval(Expr),
    /// Option Explicit
    OptionExplicit,
    /// On Error Resume Next
    OnErrorResumeNext,
    /// On Error Goto 0
    OnErrorGoto0,
    /// Resume Next
    ResumeNext,
    /// 表达式语句
    Expr(Expr),
}

/// If 分支
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct IfBranch {
    pub cond: Expr,
    pub body: Vec<Stmt>,
}

/// 函数参数
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Param {
    pub name: String,
    pub is_byref: bool,
    pub default: Option<Expr>,
}

impl Param {
    pub fn new(name: impl Into<String>) -> Self {
        Param {
            name: name.into(),
            is_byref: false,
            default: None,
        }
    }

    pub fn byref(mut self) -> Self {
        self.is_byref = true;
        self
    }

    pub fn default(mut self, value: Expr) -> Self {
        self.default = Some(value);
        self
    }
}

/// 类成员
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum ClassMember {
    Dim {
        name: String,
        is_public: bool,
    },
    Const {
        name: String,
        value: Expr,
        is_public: bool,
    },
    Sub {
        name: String,
        params: Vec<Param>,
        body: Vec<Stmt>,
        is_public: bool,
    },
    Function {
        name: String,
        params: Vec<Param>,
        body: Vec<Stmt>,
        is_public: bool,
    },
    PropertyGet {
        name: String,
        params: Vec<Param>,
        body: Vec<Stmt>,
        is_public: bool,
    },
    PropertyLet {
        name: String,
        params: Vec<Param>,
        body: Vec<Stmt>,
        is_public: bool,
    },
    PropertySet {
        name: String,
        params: Vec<Param>,
        body: Vec<Stmt>,
        is_public: bool,
    },
}
