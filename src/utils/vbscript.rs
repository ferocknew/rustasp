//! VBScript 公共特性工具模块
//!
//! 提供 VBScript 语言特性的公共辅助函数，如：
//! - 大小写不敏感的标识符比较
//! - 标识符规范化

/// 将标识符转换为规范形式（小写）
///
/// VBScript 是大小写不敏感的语言，内部使用小写形式作为统一表示
#[inline]
pub fn normalize_identifier(name: &str) -> String {
    name.to_lowercase()
}

/// 比较两个标识符是否相等（大小写不敏感）
#[inline]
pub fn identifier_eq(a: &str, b: &str) -> bool {
    a.eq_ignore_ascii_case(b)
}

/// 检查标识符是否匹配指定名称（大小写不敏感）
#[inline]
pub fn identifier_matches(name: &str, target: &str) -> bool {
    name.eq_ignore_ascii_case(target)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_normalize_identifier() {
        assert_eq!(normalize_identifier("Response"), "response");
        assert_eq!(normalize_identifier("REQUEST"), "request");
        assert_eq!(normalize_identifier("Server"), "server");
    }

    #[test]
    fn test_identifier_eq() {
        assert!(identifier_eq("Response", "response"));
        assert!(identifier_eq("REQUEST", "request"));
        assert!(identifier_eq("Server", "SERVER"));
        assert!(!identifier_eq("Server", "Client"));
    }

    #[test]
    fn test_identifier_matches() {
        assert!(identifier_matches("Response", "response"));
        assert!(identifier_matches("REQUEST", "request"));
    }
}
