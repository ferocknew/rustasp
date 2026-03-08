//! 字符串函数执行器

use crate::runtime::{RuntimeError, Value, ValueConversion};
use super::super::token::BuiltinToken;
use std::sync::{Arc, Mutex};

pub fn execute(token: BuiltinToken, args: &[Value]) -> Result<Option<Value>, RuntimeError> {
    let result = match token {
        BuiltinToken::Len => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            Value::Number(ValueConversion::to_string(&args[0]).len() as f64)
        }
        BuiltinToken::Trim => string_unary(args, |s| s.trim().to_string())?,
        BuiltinToken::LTrim => string_unary(args, |s| s.trim_start().to_string())?,
        BuiltinToken::RTrim => string_unary(args, |s| s.trim_end().to_string())?,
        BuiltinToken::UCase => string_unary(args, |s| s.to_uppercase())?,
        BuiltinToken::LCase => string_unary(args, |s| s.to_lowercase())?,
        BuiltinToken::Left => {
            if args.len() < 2 {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let s = ValueConversion::to_string(&args[0]);
            let n = ValueConversion::to_number(&args[1]) as usize;
            Value::String(s.chars().take(n).collect::<String>())
        }
        BuiltinToken::Right => {
            if args.len() < 2 {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let s = ValueConversion::to_string(&args[0]);
            let n = ValueConversion::to_number(&args[1]) as usize;
            Value::String(s.chars().rev().take(n).collect::<String>().chars().rev().collect::<String>())
        }
        BuiltinToken::Mid => {
            if args.len() < 2 {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let s = ValueConversion::to_string(&args[0]);
            let start = ValueConversion::to_number(&args[1]) as usize;
            let length = args.get(2).map(|v| ValueConversion::to_number(v) as usize).unwrap_or(s.len());
            Value::String(s.chars().skip(start.saturating_sub(1)).take(length).collect::<String>())
        }
        BuiltinToken::Asc => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let s = ValueConversion::to_string(&args[0]);
            Value::Number(s.chars().next().map(|c| c as u8 as f64).unwrap_or(0.0))
        }
        BuiltinToken::AscW => {
            // AscW - 返回 Unicode 码点（16 位）
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let s = ValueConversion::to_string(&args[0]);
            Value::Number(s.chars().next().map(|c| c as u32 as f64).unwrap_or(0.0))
        }
        BuiltinToken::Chr => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let n = ValueConversion::to_number(&args[0]) as u32;
            Value::String(char::from_u32(n).unwrap_or('\0').to_string())
        }
        BuiltinToken::ChrW => {
            // ChrW - 返回 Unicode 字符（与 Chr 相同，因为 Rust 使用 UTF-8）
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let n = ValueConversion::to_number(&args[0]) as u32;
            Value::String(char::from_u32(n).unwrap_or('\0').to_string())
        }
        BuiltinToken::InStr => {
            // InStr([start, ]string1, string2[, compare])
            // 返回 string2 在 string1 中首次出现的位置
            if args.len() < 2 {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let (start, string1, string2) = if args.len() >= 3 {
                // 有 start 参数
                (Some(ValueConversion::to_number(&args[0]) as usize),
                 ValueConversion::to_string(&args[1]),
                 ValueConversion::to_string(&args[2]))
            } else {
                // 没有 start 参数
                (None, ValueConversion::to_string(&args[0]), ValueConversion::to_string(&args[1]))
            };

            // 从指定位置开始搜索（VBScript 位置从 1 开始）
            let search_str = if let Some(pos) = start {
                if pos > string1.len() {
                    return Ok(Some(Value::Number(0.0)));
                } else if pos > 1 {
                    string1.chars().skip(pos - 1).collect::<String>()
                } else {
                    string1
                }
            } else {
                string1
            };

            // 查找子串（不区分大小写）
            let search_lower = search_str.to_lowercase();
            let target_lower = string2.to_lowercase();
            let pos = search_lower.find(&target_lower);
            let result = if let Some(p) = pos {
                // VBScript 位置从 1 开始
                let base_pos = start.unwrap_or(1);
                (base_pos + p) as f64
            } else {
                0.0
            };
            Value::Number(result)
        }
        BuiltinToken::InStrRev => {
            // InStrRev(string1, string2[, start[, compare]])
            // 从字符串末尾开始查找
            if args.len() < 2 {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let string1 = ValueConversion::to_string(&args[0]);
            let string2 = ValueConversion::to_string(&args[1]);
            let start = if args.len() >= 3 {
                Some(ValueConversion::to_number(&args[2]) as usize)
            } else {
                None
            };

            // 从末尾查找子串
            let search_str = if let Some(pos) = start {
                if pos > string1.len() {
                    return Ok(Some(Value::Number(0.0)));
                } else if pos > 0 {
                    string1.chars().take(pos).collect::<String>()
                } else {
                    string1.clone()
                }
            } else {
                string1.clone()
            };

            // 使用 rfind 从右查找（不区分大小写）
            let search_lower = search_str.to_lowercase();
            let target_lower = string2.to_lowercase();
            let pos = search_lower.rfind(&target_lower);
            let result = if let Some(p) = pos {
                (p + 1) as f64  // VBScript 位置从 1 开始
            } else {
                0.0
            };
            Value::Number(result)
        }
        BuiltinToken::StrComp => {
            // StrComp(string1, string2[, compare])
            // 比较两个字符串，返回 -1, 0, 或 1
            if args.len() < 2 {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let string1 = ValueConversion::to_string(&args[0]);
            let string2 = ValueConversion::to_string(&args[1]);

            let result = if string1.eq_ignore_ascii_case(&string2) {
                0.0
            } else if string1.to_lowercase() < string2.to_lowercase() {
                -1.0
            } else {
                1.0
            };
            Value::Number(result)
        }
        BuiltinToken::Replace => {
            // Replace(string, find, replacewith[, start[, count[, compare]]])
            if args.len() < 3 {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let string = ValueConversion::to_string(&args[0]);
            let find = ValueConversion::to_string(&args[1]);
            let replace_with = ValueConversion::to_string(&args[2]);
            let start = if args.len() >= 4 {
                ValueConversion::to_number(&args[3]) as usize
            } else {
                1
            };
            let count = if args.len() >= 5 {
                ValueConversion::to_number(&args[4]) as usize
            } else {
                usize::MAX
            };

            // 获取要搜索的子串
            let search_str = if start > 1 && start <= string.len() {
                &string[start - 1..]
            } else if start > string.len() {
                return Ok(Some(Value::String(string)));
            } else {
                &string
            };

            // 执行替换
            let result = if count == usize::MAX {
                search_str.replacen(&find, &replace_with, usize::MAX)
            } else {
                search_str.replacen(&find, &replace_with, count)
            };

            // 组合结果
            let final_result = if start > 1 && start <= string.len() {
                format!("{}{}", &string[..start - 1], result)
            } else {
                result
            };

            Value::String(final_result)
        }
        BuiltinToken::Split => {
            // Split(string[, delimiter[, count[, compare]]])
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let string = ValueConversion::to_string(&args[0]);
            let delimiter = args.get(1).map(|v| ValueConversion::to_string(v)).unwrap_or(" ".to_string());

            let parts: Vec<Value> = if delimiter.is_empty() {
                // 空分隔符，返回单个字符数组
                string.chars().map(|c| Value::String(c.to_string())).collect()
            } else {
                string.split(&delimiter).map(|s| Value::String(s.to_string())).collect()
            };

            use crate::runtime::VbsArray;
            Value::Array(Arc::new(Mutex::new(VbsArray::from_vec(parts))))
        }
        BuiltinToken::Join => {
            // Join(array[, delimiter])
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            match &args[0] {
                Value::Array(ref arr) => {
                    let delimiter = args.get(1).map(|v| ValueConversion::to_string(v)).unwrap_or(" ".to_string());
                    let locked_arr = arr.lock().unwrap();
                    Value::String(locked_arr.data.iter().map(|v| ValueConversion::to_string(v)).collect::<Vec<_>>().join(&delimiter))
                }
                _ => return Err(RuntimeError::TypeMismatch("Expected array".to_string())),
            }
        }
        BuiltinToken::Space => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let n = ValueConversion::to_number(&args[0]) as usize;
            Value::String(" ".repeat(n.min(1000000)))
        }
        BuiltinToken::String_ => {
            // String(number, character) - 返回重复的字符
            if args.len() < 2 {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let n = ValueConversion::to_number(&args[0]) as usize;
            let ch = ValueConversion::to_string(&args[1]).chars().next().unwrap_or(' ');
            Value::String(ch.to_string().repeat(n.min(1000000)))
        }
        BuiltinToken::StrReverse => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            Value::String(ValueConversion::to_string(&args[0]).chars().rev().collect::<String>())
        }
        BuiltinToken::LenB => {
            // LenB - 返回字符串的字节长度
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            Value::Number(ValueConversion::to_string(&args[0]).as_bytes().len() as f64)
        }
        BuiltinToken::LeftB => {
            // LeftB - 返回左边指定字节数
            if args.len() < 2 {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let s = ValueConversion::to_string(&args[0]);
            let n = ValueConversion::to_number(&args[1]) as usize;
            let bytes = s.as_bytes();
            // 截取前 n 个字节，注意不能在多字节字符中间截断
            let end = bytes.len().min(n);
            // 使用 from_utf8 取前 end 个字节
            match std::str::from_utf8(&bytes[..end]) {
                Ok(result) => Value::String(result.to_string()),
                Err(_) => {
                    // 如果截断位置不在字符边界，找到最近的边界
                    let valid_end = s.floor_char_boundary(end);
                    Value::String(s[..valid_end].to_string())
                }
            }
        }
        BuiltinToken::RightB => {
            // RightB - 返回右边指定字节数
            if args.len() < 2 {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let s = ValueConversion::to_string(&args[0]);
            let n = ValueConversion::to_number(&args[1]) as usize;
            let bytes = s.as_bytes();
            // 从末尾取 n 个字节
            let start = if bytes.len() > n { bytes.len() - n } else { 0 };
            // 找到最近的字符边界
            let valid_start = s.ceil_char_boundary(start);
            Value::String(s[valid_start..].to_string())
        }
        BuiltinToken::MidB => {
            // MidB - 返回中间指定字节数
            if args.len() < 2 {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let s = ValueConversion::to_string(&args[0]);
            let start = ValueConversion::to_number(&args[1]) as usize;
            let length = args.get(2).map(|v| ValueConversion::to_number(v) as usize).unwrap_or(s.as_bytes().len());
            let bytes = s.as_bytes();
            // VBScript 位置从 1 开始，转换为从 0 开始
            let start_byte = start.saturating_sub(1).min(bytes.len());
            let end_byte = (start_byte + length).min(bytes.len());
            // 找到最近的字符边界
            let valid_start = s.ceil_char_boundary(start_byte);
            let valid_end = s.floor_char_boundary(end_byte);
            if valid_start < valid_end {
                Value::String(s[valid_start..valid_end].to_string())
            } else {
                Value::String(String::new())
            }
        }
        BuiltinToken::InStrB => {
            // InStrB - 按字节查找字符串
            if args.len() < 2 {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let (start, string1, string2) = if args.len() >= 3 {
                // 有 start 参数
                (Some(ValueConversion::to_number(&args[0]) as usize),
                 ValueConversion::to_string(&args[1]),
                 ValueConversion::to_string(&args[2]))
            } else {
                // 没有 start 参数
                (None, ValueConversion::to_string(&args[0]), ValueConversion::to_string(&args[1]))
            };

            let bytes1 = string1.as_bytes();
            let bytes2 = string2.as_bytes();

            // 从指定字节位置开始搜索（VBScript 位置从 1 开始）
            let search_bytes = if let Some(pos) = start {
                let start_byte = pos.saturating_sub(1).min(bytes1.len());
                &bytes1[start_byte..]
            } else {
                bytes1
            };

            // 查找子串（区分大小写，字节匹配）
            let pos = search_bytes.windows(bytes2.len()).position(|w| w == bytes2);
            let result = if let Some(p) = pos {
                let base_pos = start.unwrap_or(1);
                (base_pos + p) as f64
            } else {
                0.0
            };
            Value::Number(result)
        }
        BuiltinToken::AscB => {
            // AscB - 返回第一个字节的 ASCII 值
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let s = ValueConversion::to_string(&args[0]);
            Value::Number(s.as_bytes().first().map(|&b| b as f64).unwrap_or(0.0))
        }
        BuiltinToken::ChrB => {
            // ChrB - 返回指定字节对应的单字节字符串
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            let n = (ValueConversion::to_number(&args[0]) as u32) & 0xFF; // 只取低 8 位
            // 将字节转换为 Latin-1 字符（单字节编码）
            let ch = (n as u8) as char;
            Value::String(ch.to_string())
        }
        _ => return Ok(None),
    };
    Ok(Some(result))
}

fn string_unary<F>(args: &[Value], f: F) -> Result<Value, RuntimeError>
where
    F: FnOnce(&str) -> String,
{
    if args.is_empty() {
        return Err(RuntimeError::ArgumentCountMismatch);
    }
    Ok(Value::String(f(&ValueConversion::to_string(&args[0]))))
}
