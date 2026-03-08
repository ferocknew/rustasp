//! 值显示

use super::Value;
use std::fmt;

impl fmt::Display for Value {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Value::Empty => write!(f, ""),
            Value::Null => write!(f, "Null"),
            Value::Nothing => write!(f, "Nothing"),
            Value::Boolean(b) => write!(f, "{}", if *b { "True" } else { "False" }),
            Value::Number(n) => {
                if n.fract() == 0.0 {
                    write!(f, "{}", *n as i64)
                } else {
                    write!(f, "{}", n)
                }
            }
            Value::String(s) => write!(f, "{}", s),
            Value::Array(arr) => {
                let locked_arr = arr.lock().unwrap();
                let items: Vec<String> = locked_arr.iter().map(|v| v.to_string()).collect();
                write!(f, "[{}]", items.join(", "))
            }
            Value::Object(obj) => {
                // 尝试作为字典显示
                let locked_obj = obj.lock().unwrap();
                if let Some(dict) = locked_obj.as_any().downcast_ref::<super::super::objects::Dictionary>() {
                    let items: Vec<String> = dict.as_hashmap()
                        .iter()
                        .map(|(k, v)| format!("{}: {}", k, v))
                        .collect();
                    write!(f, "{{{}}}", items.join(", "))
                } else {
                    // 其他对象显示类型名
                    write!(f, "[object]")
                }
            }
        }
    }
}
