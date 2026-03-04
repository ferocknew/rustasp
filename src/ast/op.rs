//! 运算符定义

use serde::{Deserialize, Serialize};

/// 二元运算符
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum BinaryOp {
    /// +
    Add,
    /// -
    Sub,
    /// *
    Mul,
    /// /
    Div,
    /// \
    IntDiv,
    /// Mod
    Mod,
    /// ^
    Pow,
    /// &
    Concat,
    /// =
    Eq,
    /// <>
    Ne,
    /// <
    Lt,
    /// <=
    Le,
    /// >
    Gt,
    /// >=
    Ge,
    /// And
    And,
    /// Or
    Or,
    /// Xor
    Xor,
    /// Is
    Is,
    /// Imp
    Imp,
    /// Eqv
    Eqv,
}

/// 一元运算符
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum UnaryOp {
    /// -
    Neg,
    /// Not
    Not,
}
