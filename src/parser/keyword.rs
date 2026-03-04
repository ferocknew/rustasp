//! VBScript 关键字定义

use serde::{Deserialize, Serialize};

/// VBScript 关键字
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum Keyword {
    // 声明
    Dim,
    Const,
    ReDim,
    Preserve,

    // 控制流
    If,
    Then,
    Else,
    ElseIf,
    End,
    Select,
    Case,

    // 循环
    For,
    To,
    Step,
    Next,
    Each,
    In,
    Do,
    While,
    Loop,
    Until,
    Wend,
    Exit,

    // 函数/过程
    Sub,
    Function,
    Call,
    Return,
    ByRef,
    ByVal,

    // 类
    Class,
    Property,
    Get,
    Let,
    Set,
    Public,
    Private,

    // 值
    True,
    False,
    Nothing,
    Empty,
    Null,

    // 运算符
    And,
    Or,
    Not,
    Xor,
    Mod,
    Is,
    Imp,
    Eqv,

    // 错误处理
    On,
    Error,
    Resume,

    // 其他
    Option,
    Explicit,
    Erase,
    Execute,
    ExecuteGlobal,
    Eval,
}

impl Keyword {
    /// 获取关键字的文本表示
    pub fn as_str(&self) -> &'static str {
        match self {
            Keyword::Dim => "Dim",
            Keyword::Const => "Const",
            Keyword::ReDim => "ReDim",
            Keyword::Preserve => "Preserve",
            Keyword::If => "If",
            Keyword::Then => "Then",
            Keyword::Else => "Else",
            Keyword::ElseIf => "ElseIf",
            Keyword::End => "End",
            Keyword::Select => "Select",
            Keyword::Case => "Case",
            Keyword::For => "For",
            Keyword::To => "To",
            Keyword::Step => "Step",
            Keyword::Next => "Next",
            Keyword::Each => "Each",
            Keyword::In => "In",
            Keyword::Do => "Do",
            Keyword::While => "While",
            Keyword::Loop => "Loop",
            Keyword::Until => "Until",
            Keyword::Wend => "Wend",
            Keyword::Exit => "Exit",
            Keyword::Sub => "Sub",
            Keyword::Function => "Function",
            Keyword::Call => "Call",
            Keyword::Return => "Return",
            Keyword::ByRef => "ByRef",
            Keyword::ByVal => "ByVal",
            Keyword::Class => "Class",
            Keyword::Property => "Property",
            Keyword::Get => "Get",
            Keyword::Let => "Let",
            Keyword::Set => "Set",
            Keyword::Public => "Public",
            Keyword::Private => "Private",
            Keyword::True => "True",
            Keyword::False => "False",
            Keyword::Nothing => "Nothing",
            Keyword::Empty => "Empty",
            Keyword::Null => "Null",
            Keyword::And => "And",
            Keyword::Or => "Or",
            Keyword::Not => "Not",
            Keyword::Xor => "Xor",
            Keyword::Mod => "Mod",
            Keyword::Is => "Is",
            Keyword::Imp => "Imp",
            Keyword::Eqv => "Eqv",
            Keyword::On => "On",
            Keyword::Error => "Error",
            Keyword::Resume => "Resume",
            Keyword::Option => "Option",
            Keyword::Explicit => "Explicit",
            Keyword::Erase => "Erase",
            Keyword::Execute => "Execute",
            Keyword::ExecuteGlobal => "ExecuteGlobal",
            Keyword::Eval => "Eval",
        }
    }

    /// 是否是一元运算符
    pub fn is_unary_op(&self) -> bool {
        matches!(self, Keyword::Not)
    }

    /// 是否是逻辑与运算符
    pub fn is_and(&self) -> bool {
        matches!(self, Keyword::And)
    }

    /// 是否是逻辑或运算符
    pub fn is_or(&self) -> bool {
        matches!(self, Keyword::Or)
    }
}
