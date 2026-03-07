pub fn execute_builtin(name: &str, args: &[Value]) -> Option<Result<Value, RuntimeError>> {
    let registry = TokenRegistry::new();
    registry.lookup(name).map(|token| {
        BuiltinExecutor::execute(token, args)
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_token_registry() {
        let registry = TokenRegistry::new();

        // 测试查找
        assert_eq!(registry.lookup("abs"), Some(BuiltinToken::Abs));
        assert_eq!(registry.lookup("ABS"), Some(BuiltinToken::Abs)); // 不区分大小写
        assert_eq!(registry.lookup("len"), Some(BuiltinToken::Len));

        // 测试未注册的函数
        assert_eq!(registry.lookup("unknown"), None);
    }

    #[test]
    fn test_math_functions() {
        // Abs
        let result = BuiltinExecutor::execute(BuiltinToken::Abs, &[Value::Number(-5.0)]).unwrap();
        assert_eq!(result, Value::Number(5.0));

        // Sqr
        let result = BuiltinExecutor::execute(BuiltinToken::Sqr, &[Value::Number(16.0)]).unwrap();
        assert_eq!(result, Value::Number(4.0));

        // Int
        let result = BuiltinExecutor::execute(BuiltinToken::Int, &[Value::Number(3.7)]).unwrap();
        assert_eq!(result, Value::Number(3.0));
    }

    #[test]
    fn test_string_functions() {
        // Len
        let result = BuiltinExecutor::execute(BuiltinToken::Len, &[Value::String("hello".to_string())]).unwrap();
        assert_eq!(result, Value::Number(5.0));

        // UCase
        let result = BuiltinExecutor::execute(BuiltinToken::UCase, &[Value::String("hello".to_string())]).unwrap();
        assert_eq!(result, Value::String("HELLO".to_string()));

        // Trim
        let result = BuiltinExecutor::execute(BuiltinToken::Trim, &[Value::String("  hello  ".to_string())]).unwrap();
        assert_eq!(result, Value::String("hello".to_string()));
    }

    #[test]
    fn test_type_conversion() {
        // CInt
        let result = BuiltinExecutor::execute(BuiltinToken::CInt, &[Value::String("42".to_string())]).unwrap();
        assert_eq!(result, Value::Number(42.0));

        // CStr
        let result = BuiltinExecutor::execute(BuiltinToken::CStr, &[Value::Number(42.0)]).unwrap();
        assert_eq!(result, Value::String("42".to_string()));
    }

    #[test]
    fn test_inspection_functions() {
        // IsNumeric
        let result = BuiltinExecutor::execute(BuiltinToken::IsNumeric, &[Value::Number(42.0)]).unwrap();
        assert_eq!(result, Value::Boolean(true));

        let result = BuiltinExecutor::execute(BuiltinToken::IsNumeric, &[Value::String("abc".to_string())]).unwrap();
        assert_eq!(result, Value::Boolean(false));

        // IsEmpty
        let result = BuiltinExecutor::execute(BuiltinToken::IsEmpty, &[Value::Empty]).unwrap();
        assert_eq!(result, Value::Boolean(true));

        // TypeName
        let result = BuiltinExecutor::execute(BuiltinToken::TypeName, &[Value::Number(42.0)]).unwrap();
        assert_eq!(result, Value::String("Double".to_string()));
    }
}
