//! 运行时错误

use thiserror::Error;

/// 控制流信号（用于 Exit 语句）
#[derive(Debug, Clone, PartialEq)]
pub enum ControlFlow {
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
}

#[derive(Debug, Error)]
pub enum RuntimeError {
    #[error("Undefined variable: {0}")]
    UndefinedVariable(String),

    #[error("Type mismatch: {0}")]
    TypeMismatch(String),

    #[error("Division by zero")]
    DivisionByZero,

    #[error("Index out of bounds: {0}")]
    IndexOutOfBounds(usize),

    #[error("Object required")]
    ObjectRequired,

    #[error("Method not found: {0}")]
    MethodNotFound(String),

    #[error("Property not found: {0}")]
    PropertyNotFound(String),

    #[error("Argument count mismatch")]
    ArgumentCountMismatch,

    #[error("Undefined function: {0}")]
    UndefinedFunction(String),

    #[error("Not an array")]
    NotAnArray,

    #[error("Invalid assignment target")]
    InvalidAssignment,

    #[error("Invalid index")]
    InvalidIndex,

    #[error("Runtime error: {0}")]
    Generic(String),

    /// 控制流信号（不是真正的错误）
    #[error("Control flow: {0:?}")]
    ControlFlow(ControlFlow),
}
