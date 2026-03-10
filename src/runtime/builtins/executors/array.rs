//! 数组函数执行器

use crate::runtime::{RuntimeError, Value, ValueConversion, VbsArray};
use super::super::token::BuiltinToken;
use std::sync::{Arc, Mutex};

pub fn execute(token: BuiltinToken, args: &[Value]) -> Result<Option<Value>, RuntimeError> {
    let result = match token {
        BuiltinToken::UBound => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            match &args[0] {
                Value::Array(ref arr) => {
                    let locked_arr = arr.lock().unwrap();
                    let dim = if args.len() > 1 {
                        let dim_num = ValueConversion::to_number(&args[1]) as usize;
                        if dim_num == 0 {
                            0
                        } else {
                            dim_num - 1  // 转换为 0-based 索引
                        }
                    } else {
                        0
                    };
                    let ubound = if dim < locked_arr.dims.len() {
                        locked_arr.dims[dim].saturating_sub(1)
                    } else {
                        0
                    };
                    Value::Number(ubound as f64)
                }
                _ => return Err(RuntimeError::TypeMismatch("Expected array".to_string())),
            }
        }
        BuiltinToken::LBound => {
            // VBScript 数组的下界始终是 0
            Value::Number(0.0)
        }
        BuiltinToken::Array => {
            // Array 函数：将参数转换为一维数组
            let vbs_arr = VbsArray::from_vec(args.to_vec());
            Value::Array(Arc::new(Mutex::new(vbs_arr)))
        }
        BuiltinToken::Filter => {
            if args.len() < 2 {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            match &args[0] {
                Value::Array(ref arr) => {
                    let locked_arr = arr.lock().unwrap();
                    let criteria = ValueConversion::to_string(&args[1]);
                    let include = args.get(2).map(|v| ValueConversion::to_bool(v)).unwrap_or(true);
                    let filtered: Vec<Value> = locked_arr.data.iter()
                        .filter(|v| {
                            let s = ValueConversion::to_string(*v);
                            if include { s.contains(&criteria) } else { !s.contains(&criteria) }
                        })
                        .cloned()
                        .collect();
                    Value::Array(Arc::new(Mutex::new(VbsArray::from_vec(filtered))))
                }
                _ => return Err(RuntimeError::TypeMismatch("Expected array".to_string())),
            }
        }
        BuiltinToken::IsArray => {
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            Value::Boolean(matches!(args[0], Value::Array(_)))
        }
        BuiltinToken::Erase => {
            // Erase - 清除数组元素，将其设置为 Empty
            if args.is_empty() {
                return Err(RuntimeError::ArgumentCountMismatch);
            }
            match &args[0] {
                Value::Array(ref arr) => {
                    let locked_arr = arr.lock().unwrap();
                    // 创建一个新数组，保持相同维度，所有元素设置为 Empty
                    let erased = VbsArray {
                        dims: locked_arr.dims.clone(),
                        data: vec![Value::Empty; locked_arr.data.len()],
                    };
                    Value::Array(Arc::new(Mutex::new(erased)))
                }
                _ => return Err(RuntimeError::TypeMismatch("Expected array".to_string())),
            }
        }
        _ => return Ok(None),
    };
    Ok(Some(result))
}
