//! 内置函数 Token 识别模块
//!
//! 使用 Token Key 而非字符串匹配，提高内置函数调用性能
//! 将函数名映射为整数 ID，通过 match/switch 快速分发

use crate::runtime::{RuntimeError, Value, ValueConversion};
use rand::Rng;
use std::collections::HashMap;

/// 内置函数 Token ID
/// 每个内置函数对应一个唯一的整数 ID，用于快速匹配
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u16)]
pub enum BuiltinToken {
    // ========== 数学函数 (1-50) ==========
    Abs = 1,
    Sqr = 2,
    Sin = 3,
    Cos = 4,
    Tan = 5,
    Atn = 6,
    Log = 7,
    Exp = 8,
    Int = 9,
    Fix = 10,
    Round = 11,
    Rnd = 12,
    Sgn = 13,

    // ========== 类型转换函数 (51-100) ==========
    CStr = 51,
    CInt = 52,
    CLng = 53,
    CSng = 54,
    CDbl = 55,
    CBool = 56,
    CDate = 57,
    CByte = 58,

    // ========== 字符串函数 (101-150) ==========
    Len = 101,
    Trim = 102,
    LTrim = 103,
    RTrim = 104,
    Left = 105,
    Right = 106,
    Mid = 107,
    UCase = 108,
    LCase = 109,
    InStr = 110,
    InStrRev = 111,
    StrComp = 112,
    Replace = 113,
    Split = 114,
    Join = 115,
    StrReverse = 116,
    Space = 117,
    String_ = 118,  // String 是 Rust 关键字，加下划线
    Asc = 119,
    Chr = 120,
    AscW = 121,
    ChrW = 122,

    // ========== 日期时间函数 (151-200) ==========
    Now = 151,
    Date = 152,
    Time = 153,
    Year = 154,
    Month = 155,
    Day = 156,
    Hour = 157,
    Minute = 158,
    Second = 159,
    WeekDay = 160,
    DateAdd = 161,
    DateDiff = 162,
    DatePart = 163,
    DateSerial = 164,
    DateValue = 165,
    TimeSerial = 166,
    TimeValue = 167,
    FormatDateTime = 168,
    MonthName = 169,
    WeekDayName = 170,

    // ========== 数组函数 (201-220) ==========
    Array = 201,
    UBound = 202,
    LBound = 203,
    Filter = 204,
    IsArray = 205,

    // ========== 检验函数 (221-250) ==========
    IsNumeric = 221,
    IsDate = 222,
    IsEmpty = 223,
    IsNull = 224,
    IsObject = 225,
    IsNothing = 226,
    TypeName = 227,
    VarType = 228,

    // ========== 交互函数 (251-270) ==========
    MsgBox = 251,
    InputBox = 252,

    // ========== 其他函数 (271-300) ==========
    CreateObject = 271,
    GetObject = 272,
    Eval = 273,
    Execute = 274,
    RGB = 275,

    // 未知函数
    Unknown = 0,
}
