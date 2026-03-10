//! VBScript 关键字定义和匹配

use serde::{Deserialize, Serialize};

/// VBScript 关键字枚举
#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
pub enum Keyword {
    // 声明
    Dim,
    Const,
    ReDim,
    Preserve,
    Option,
    Explicit,
    Private,
    Public,

    // 控制流
    If,
    Then,
    Else,
    ElseIf,
    End,
    For,
    Each,
    To,
    Step,
    Next,
    In,
    While,
    Wend,
    Do,
    Loop,
    Until,
    Select,
    Case,

    // 函数和过程
    Function,
    Sub,
    Call,
    Exit,
    Return,
    ByRef,
    ByVal,

    // 错误处理
    On,
    Error,
    Resume,

    // 对象
    Set,
    New,
    Nothing,
    Class,
    Property,
    Get,
    Let,

    // 运算符关键字
    Mod,
    And,
    Or,
    Not,

    // 类型检查
    Is,
    Type,

    // 逻辑值
    True,
    False,
    Null,
    Empty,

    // 动态执行
    Execute,
    ExecuteGlobal,
    Eval,
}

impl Keyword {
    /// 将关键字转换为字符串
    pub fn as_str(&self) -> &'static str {
        match self {
            Keyword::Dim => "Dim",
            Keyword::Const => "Const",
            Keyword::ReDim => "ReDim",
            Keyword::Preserve => "Preserve",
            Keyword::Option => "Option",
            Keyword::Explicit => "Explicit",
            Keyword::Private => "Private",
            Keyword::Public => "Public",
            Keyword::If => "If",
            Keyword::Then => "Then",
            Keyword::Else => "Else",
            Keyword::ElseIf => "ElseIf",
            Keyword::End => "End",
            Keyword::For => "For",
            Keyword::Each => "Each",
            Keyword::To => "To",
            Keyword::Step => "Step",
            Keyword::Next => "Next",
            Keyword::In => "In",
            Keyword::While => "While",
            Keyword::Wend => "Wend",
            Keyword::Do => "Do",
            Keyword::Loop => "Loop",
            Keyword::Until => "Until",
            Keyword::Select => "Select",
            Keyword::Case => "Case",
            Keyword::Function => "Function",
            Keyword::Sub => "Sub",
            Keyword::Call => "Call",
            Keyword::Exit => "Exit",
            Keyword::Return => "Return",
            Keyword::ByRef => "ByRef",
            Keyword::ByVal => "ByVal",
            Keyword::On => "On",
            Keyword::Error => "Error",
            Keyword::Resume => "Resume",
            Keyword::Set => "Set",
            Keyword::New => "New",
            Keyword::Nothing => "Nothing",
            Keyword::Class => "Class",
            Keyword::Property => "Property",
            Keyword::Get => "Get",
            Keyword::Let => "Let",
            Keyword::Mod => "Mod",
            Keyword::And => "And",
            Keyword::Or => "Or",
            Keyword::Not => "Not",
            Keyword::Is => "Is",
            Keyword::Type => "Type",
            Keyword::True => "True",
            Keyword::False => "False",
            Keyword::Null => "Null",
            Keyword::Empty => "Empty",
            Keyword::Execute => "Execute",
            Keyword::ExecuteGlobal => "ExecuteGlobal",
            Keyword::Eval => "Eval",
        }
    }

    /// 检查是否是一元运算符
    pub fn is_unary_op(&self) -> bool {
        matches!(self, Keyword::Not)
    }

    /// 检查是否是 Or 运算符
    pub fn is_or(&self) -> bool {
        matches!(self, Keyword::Or)
    }

    /// 检查是否是 And 运算符
    pub fn is_and(&self) -> bool {
        matches!(self, Keyword::And)
    }
}

/// 关键字查找表
static KEYWORDS: &[(&str, Keyword)] = &[
    ("dim", Keyword::Dim),
    ("const", Keyword::Const),
    ("redim", Keyword::ReDim),
    ("preserve", Keyword::Preserve),
    ("option", Keyword::Option),
    ("explicit", Keyword::Explicit),
    ("private", Keyword::Private),
    ("public", Keyword::Public),
    ("if", Keyword::If),
    ("then", Keyword::Then),
    ("else", Keyword::Else),
    ("elseif", Keyword::ElseIf),
    ("end", Keyword::End),
    ("for", Keyword::For),
    ("each", Keyword::Each),
    ("to", Keyword::To),
    ("step", Keyword::Step),
    ("next", Keyword::Next),
    ("in", Keyword::In),
    ("while", Keyword::While),
    ("wend", Keyword::Wend),
    ("do", Keyword::Do),
    ("loop", Keyword::Loop),
    ("until", Keyword::Until),
    ("select", Keyword::Select),
    ("case", Keyword::Case),
    ("function", Keyword::Function),
    ("sub", Keyword::Sub),
    ("call", Keyword::Call),
    ("exit", Keyword::Exit),
    ("return", Keyword::Return),
    ("byref", Keyword::ByRef),
    ("byval", Keyword::ByVal),
    ("on", Keyword::On),
    ("error", Keyword::Error),
    ("resume", Keyword::Resume),
    ("set", Keyword::Set),
    ("new", Keyword::New),
    ("nothing", Keyword::Nothing),
    ("class", Keyword::Class),
    ("property", Keyword::Property),
    ("get", Keyword::Get),
    ("let", Keyword::Let),
    ("mod", Keyword::Mod),
    ("and", Keyword::And),
    ("or", Keyword::Or),
    ("not", Keyword::Not),
    ("is", Keyword::Is),
    ("type", Keyword::Type),
    ("true", Keyword::True),
    ("false", Keyword::False),
    ("null", Keyword::Null),
    ("empty", Keyword::Empty),
    ("execute", Keyword::Execute),
    ("executeglobal", Keyword::ExecuteGlobal),
    ("eval", Keyword::Eval),
];

/// 查找关键字
pub fn lookup_keyword(s: &str) -> Option<Keyword> {
    let lower = s.to_ascii_lowercase();
    KEYWORDS
        .iter()
        .find(|(k, _)| *k == lower)
        .map(|(_, v)| *v)
}
