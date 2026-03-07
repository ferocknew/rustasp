//! 内置函数 Token 识别模块
//!
//! 使用 Token Key 而非字符串匹配，提高内置函数调用性能
//! 将函数名映射为整数 ID，通过 match/switch 快速分发

use crate::runtime::{RuntimeError, Value};

mod token;
mod registry;
mod executors;

#[allow(unused_imports)]
pub use token::BuiltinToken;
pub use registry::TokenRegistry;
pub use executors::BuiltinExecutor;

/// 便捷函数：通过函数名直接执行内置函数
#[allow(dead_code)]
pub fn execute_builtin(name: &str, args: &[Value]) -> Option<Result<Value, RuntimeError>> {
    let registry = TokenRegistry::new();
    registry.lookup(name).map(|token| BuiltinExecutor::execute(token, args))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_token_registry() {
        let registry = TokenRegistry::new();

        assert_eq!(registry.lookup("abs"), Some(BuiltinToken::Abs));
        assert_eq!(registry.lookup("ABS"), Some(BuiltinToken::Abs));
        assert_eq!(registry.lookup("len"), Some(BuiltinToken::Len));
        assert_eq!(registry.lookup("unknown"), None);
    }

    #[test]
    fn test_math_functions() {
        use crate::runtime::Value;
        let result = BuiltinExecutor::execute(BuiltinToken::Abs, &[Value::Number(-5.0)]).unwrap();
        assert_eq!(result, Value::Number(5.0));

        let result = BuiltinExecutor::execute(BuiltinToken::Sqr, &[Value::Number(16.0)]).unwrap();
        assert_eq!(result, Value::Number(4.0));

        let result = BuiltinExecutor::execute(BuiltinToken::Int, &[Value::Number(3.7)]).unwrap();
        assert_eq!(result, Value::Number(3.0));
    }

    #[test]
    fn test_string_functions() {
        use crate::runtime::Value;
        let result = BuiltinExecutor::execute(BuiltinToken::Len, &[Value::String("hello".to_string())]).unwrap();
        assert_eq!(result, Value::Number(5.0));

        let result = BuiltinExecutor::execute(BuiltinToken::UCase, &[Value::String("hello".to_string())]).unwrap();
        assert_eq!(result, Value::String("HELLO".to_string()));

        let result = BuiltinExecutor::execute(BuiltinToken::Trim, &[Value::String("  hello  ".to_string())]).unwrap();
        assert_eq!(result, Value::String("hello".to_string()));
    }

    #[test]
    fn test_type_conversion() {
        use crate::runtime::Value;
        let result = BuiltinExecutor::execute(BuiltinToken::CInt, &[Value::String("42".to_string())]).unwrap();
        assert_eq!(result, Value::Number(42.0));

        let result = BuiltinExecutor::execute(BuiltinToken::CStr, &[Value::Number(42.0)]).unwrap();
        assert_eq!(result, Value::String("42".to_string()));
    }

    #[test]
    fn test_inspection_functions() {
        use crate::runtime::Value;
        let result = BuiltinExecutor::execute(BuiltinToken::IsNumeric, &[Value::Number(42.0)]).unwrap();
        assert_eq!(result, Value::Boolean(true));

        let result = BuiltinExecutor::execute(BuiltinToken::IsNumeric, &[Value::String("abc".to_string())]).unwrap();
        assert_eq!(result, Value::Boolean(false));

        let result = BuiltinExecutor::execute(BuiltinToken::IsEmpty, &[Value::Empty]).unwrap();
        assert_eq!(result, Value::Boolean(true));

        let result = BuiltinExecutor::execute(BuiltinToken::TypeName, &[Value::Number(42.0)]).unwrap();
        assert_eq!(result, Value::String("Double".to_string()));
    }
}
