//! 值类型转换

use super::Value;

/// 值转换 trait
pub trait ValueConversion {
    /// 转换为布尔值
    fn to_bool(&self) -> bool;

    /// 转换为数字
    fn to_number(&self) -> f64;

    /// 转换为字符串
    fn to_string(&self) -> String;

    /// 是否为真值
    fn is_truthy(&self) -> bool;
}

impl ValueConversion for Value {
    fn to_bool(&self) -> bool {
        match self {
            Value::Boolean(b) => *b,
            Value::Number(n) => *n != 0.0,
            Value::String(s) => {
                if s.is_empty() {
                    false
                } else if s.eq_ignore_ascii_case("true") {
                    true
                } else if s.eq_ignore_ascii_case("false") {
                    false
                } else {
                    s.parse::<f64>().map(|n| n != 0.0).unwrap_or(false)
                }
            }
            Value::Empty => false,
            Value::Null => false,
            Value::Nothing => false,
            Value::Array(_) => true,
            Value::Object(_) => true,
        }
    }

    fn to_number(&self) -> f64 {
        match self {
            Value::Number(n) => *n,
            Value::Boolean(b) => {
                if *b {
                    -1.0
                } else {
                    0.0
                }
            }
            Value::String(s) => {
                if s.is_empty() {
                    0.0
                } else {
                    s.parse::<f64>().unwrap_or(0.0)
                }
            }
            Value::Empty => 0.0,
            Value::Null => 0.0,
            Value::Nothing => 0.0,
            Value::Array(_) => 0.0,
            Value::Object(_) => 0.0,
        }
    }

    fn to_string(&self) -> String {
        match self {
            Value::String(s) => s.clone(),
            Value::Number(n) => {
                if n.fract() == 0.0 {
                    format!("{}", *n as i64)
                } else {
                    format!("{}", n)
                }
            }
            Value::Boolean(b) => if *b { "True" } else { "False" }.to_string(),
            Value::Empty => String::new(),
            Value::Null => "Null".to_string(),
            Value::Nothing => "Nothing".to_string(),
            Value::Array(arr) => {
                let items: Vec<String> = arr.iter().map(|v| ValueConversion::to_string(v)).collect();
                items.join(", ")
            }
            Value::Object(_) => "[object]".to_string(),
        }
    }

    fn is_truthy(&self) -> bool {
        match self {
            Value::Boolean(b) => *b,
            Value::Number(n) => *n != 0.0,
            Value::String(s) => !s.is_empty(),
            Value::Empty => false,
            Value::Null => false,
            Value::Nothing => false,
            Value::Array(_) => true,
            Value::Object(_) => true,
        }
    }
}
