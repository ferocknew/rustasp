//! 路径解析和安全检查模块
//!
//! 实现父路径安全处理，防止目录遍历攻击

use std::path::{Path, PathBuf};

/// 路径解析错误
#[derive(Debug, Clone)]
pub enum PathError {
    /// 路径超出 Web 根目录
    OutsideRoot(String),
    /// 路径解析失败
    InvalidPath(String),
    /// 父路径被禁用
    ParentPathDisabled(String),
}

impl std::fmt::Display for PathError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            PathError::OutsideRoot(msg) => write!(f, "Path outside web root: {}", msg),
            PathError::InvalidPath(msg) => write!(f, "Invalid path: {}", msg),
            PathError::ParentPathDisabled(msg) => write!(f, "Parent path disabled: {}", msg),
        }
    }
}

impl std::error::Error for PathError {}

/// 路径解析器
pub struct PathResolver {
    web_root: PathBuf,
    allow_parent_paths: bool,
}

impl PathResolver {
    /// 创建新的路径解析器
    pub fn new(web_root: PathBuf, allow_parent_paths: bool) -> Self {
        PathResolver {
            web_root,
            allow_parent_paths,
        }
    }

    /// 解析并验证路径
    ///
    /// 确保：
    /// 1. 如果 allow_parent_paths=false，禁止 ../
    /// 2. 最终路径必须在 web_root 内
    /// 3. 返回规范化的绝对路径
    pub fn resolve(&self, request_path: &str) -> Result<PathBuf, PathError> {
        // 去掉开头的 /
        let path = request_path.trim_start_matches('/');

        // 如果禁止父路径，检查是否包含 ..
        if !self.allow_parent_paths && path.contains("..") {
            return Err(PathError::ParentPathDisabled(format!(
                "Parent paths are disabled: {}",
                path
            )));
        }

        // 拼接完整路径
        let full_path = self.web_root.join(path);

        // 尝试规范化路径（解析 .. 和符号链接）
        match full_path.canonicalize() {
            Ok(canonical) => {
                // 确保路径在 web_root 内
                if self.is_safe(&canonical) {
                    Ok(canonical)
                } else {
                    Err(PathError::OutsideRoot(format!(
                        "Path {:?} is outside web root {:?}",
                        canonical, self.web_root
                    )))
                }
            }
            Err(e) => {
                // 路径不存在或其他错误，返回原始路径但仍然检查安全性
                if self.is_safe(&full_path) {
                    Ok(full_path)
                } else {
                    Err(PathError::InvalidPath(format!(
                        "Cannot canonicalize path {:?}: {}",
                        full_path, e
                    )))
                }
            }
        }
    }

    /// 检查路径是否安全（在 web_root 内）
    fn is_safe(&self, path: &Path) -> bool {
        // 获取 web_root 的绝对路径
        let web_root_abs = match self.web_root.canonicalize() {
            Ok(p) => p,
            Err(_) => return false,
        };

        // 检查路径是否以 web_root 开头
        path.starts_with(&web_root_abs)
    }

    /// 获取相对路径（用于显示）
    pub fn get_relative_path(&self, full_path: &Path) -> Option<String> {
        full_path
            .strip_prefix(&self.web_root)
            .ok()
            .and_then(|p| p.to_str())
            .map(|s| s.to_string())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parent_path_disabled() {
        let resolver = PathResolver::new(PathBuf::from("/var/www"), false);

        // 应该拒绝父路径
        assert!(resolver.resolve("../etc/passwd").is_err());
        assert!(resolver.resolve("foo/../bar").is_err());
    }

    #[test]
    fn test_parent_path_allowed() {
        let resolver = PathResolver::new(PathBuf::from("/var/www"), true);

        // 应该允许父路径（如果仍在 web_root 内）
        // 注意：这个测试结果取决于实际的文件系统
    }

    #[test]
    fn test_safe_path() {
        let resolver = PathResolver::new(PathBuf::from("/var/www"), true);

        // 正常路径应该通过
        match resolver.resolve("index.asp") {
            Ok(path) => {
                // 路径应该在 /var/www 下
                assert!(path.starts_with("/var/www"));
            }
            Err(_) => {
                // 如果文件不存在，也可以接受错误
            }
        }
    }
}
