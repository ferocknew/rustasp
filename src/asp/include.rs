//! ASP Include 指令处理器
//!
//! 支持 <!--#include file="..." --> 和 <!--#include virtual="..." -->

use std::collections::HashSet;
use std::path::{Path, PathBuf};

/// Include 指令类型
#[derive(Debug, Clone, PartialEq)]
enum IncludeType {
    /// 相对于当前文件的路径
    File,
    /// 相对于网站根目录的路径
    Virtual,
}

/// 解析 include 指令
fn parse_include_directive(line: &str) -> Option<(IncludeType, String)> {
    let line = line.trim();

    // 匹配 <!--#include file="..." -->
    if let Some(rest) = line.strip_prefix("<!--#include") {
        let rest = rest.trim();

        // 匹配 file=
        if let Some(path_rest) = rest.strip_prefix("file") {
            let path_rest = path_rest.trim_start();
            if let Some(path_rest) = path_rest.strip_prefix('=') {
                let path_rest = path_rest.trim();
                // 提取引号中的路径
                let path = extract_quoted_path(path_rest)?;
                return Some((IncludeType::File, path));
            }
        }

        // 匹配 virtual=
        if let Some(path_rest) = rest.strip_prefix("virtual") {
            let path_rest = path_rest.trim_start();
            if let Some(path_rest) = path_rest.strip_prefix('=') {
                let path_rest = path_rest.trim();
                let path = extract_quoted_path(path_rest)?;
                return Some((IncludeType::Virtual, path));
            }
        }
    }

    None
}

/// 提取引号中的路径
fn extract_quoted_path(s: &str) -> Option<String> {
    let s = s.trim_start();
    if s.is_empty() {
        return None;
    }

    let quote_char = s.chars().next()?;
    if quote_char != '"' && quote_char != '\'' {
        return None;
    }

    let rest = &s[quote_char.len_utf8()..];
    if let Some(end_pos) = rest.find(quote_char) {
        Some(rest[..end_pos].to_string())
    } else {
        None
    }
}

/// 处理 include 指令
///
/// # Arguments
/// * `source` - ASP 源代码
/// * `current_file` - 当前文件的完整路径
/// * `web_root` - 网站根目录路径
/// * `processed_files` - 已处理的文件集合（用于检测循环 include）
///
/// # Returns
/// * 处理后的源代码
pub fn process_includes(
    source: &str,
    current_file: &Path,
    web_root: &Path,
    processed_files: &mut HashSet<PathBuf>,
) -> Result<String, String> {
    let mut result = String::new();

    for line in source.lines() {
        let trimmed = line.trim();

        // 检查是否是 include 指令
        if let Some((include_type, include_path)) = parse_include_directive(trimmed) {
            // 解析 include 文件的实际路径
            let resolved_path =
                resolve_include_path(&include_type, &include_path, current_file, web_root)?;

            // 检查循环 include
            if processed_files.contains(&resolved_path) {
                return Err(format!(
                    "循环 include 检测: {} -> {}",
                    current_file.display(),
                    resolved_path.display()
                ));
            }

            // 读取 include 文件内容
            let include_content = std::fs::read_to_string(&resolved_path)
                .map_err(|e| format!("无法读取 include 文件 {}: {}", resolved_path.display(), e))?;

            // 标记为已处理
            processed_files.insert(resolved_path.clone());

            // 递归处理 include 文件中的 include
            let processed_include =
                process_includes(&include_content, &resolved_path, web_root, processed_files)?;

            // 在 include 内容前添加文件名标记，方便错误定位
            // 使用相对路径显示，更易读
            let relative_path = resolved_path
                .strip_prefix(web_root)
                .unwrap_or(&resolved_path);
            result.push_str(&format!(
                "' === Included from: {} ===\n",
                relative_path.display()
            ));
            result.push_str(&processed_include);

            // 确保内容以换行符结束，避免与后续代码连在一起
            // 特别处理 <!--#include ... --><% 这种紧挨着的写法
            if !processed_include.ends_with('\n') {
                result.push('\n');
            }
        } else {
            // 普通行，直接添加
            result.push_str(line);
            result.push('\n');
        }
    }

    Ok(result)
}

/// 解析 include 路径
fn resolve_include_path(
    include_type: &IncludeType,
    include_path: &str,
    current_file: &Path,
    web_root: &Path,
) -> Result<PathBuf, String> {
    let resolved = match include_type {
        IncludeType::File => {
            // file: 相对于当前文件所在目录
            let current_dir = current_file
                .parent()
                .ok_or_else(|| format!("无法获取当前文件的父目录: {}", current_file.display()))?;
            current_dir.join(include_path)
        }
        IncludeType::Virtual => {
            // virtual: 相对于网站根目录
            let normalized_path = include_path.trim_start_matches('/');
            web_root.join(normalized_path)
        }
    };

    // 规范化路径（处理 ../ 和 ./）
    let canonical = resolved
        .canonicalize()
        .map_err(|e| format!("路径解析失败 {}: {}", resolved.display(), e))?;

    // 安全检查：确保路径在 web_root 内
    if !canonical.starts_with(web_root) {
        return Err(format!(
            "安全错误: include 路径超出 web root: {}",
            canonical.display()
        ));
    }

    Ok(canonical)
}

/// 预处理 ASP 文件（入口函数）
///
/// # Arguments
/// * `source` - ASP 源代码
/// * `file_path` - 当前文件的完整路径
/// * `web_root` - 网站根目录路径
pub fn preprocess(source: &str, file_path: &Path, web_root: &Path) -> Result<String, String> {
    // 确保 web_root 是 canonicalized 路径
    let web_root = web_root
        .canonicalize()
        .map_err(|e| format!("无法解析 web root 路径: {}", e))?;

    let mut processed_files = HashSet::new();
    processed_files.insert(file_path.to_path_buf());
    process_includes(source, file_path, &web_root, &mut processed_files)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_include_file() {
        let line = "<!--#include file=\"test.asp\" -->";
        let result = parse_include_directive(line);
        assert_eq!(result, Some((IncludeType::File, "test.asp".to_string())));
    }

    #[test]
    fn test_parse_include_virtual() {
        let line = "<!--#include virtual=\"/includes/header.asp\" -->";
        let result = parse_include_directive(line);
        assert_eq!(
            result,
            Some((IncludeType::Virtual, "/includes/header.asp".to_string()))
        );
    }

    #[test]
    fn test_parse_include_with_spaces() {
        let line = "<!--#include   file  =  \"test.asp\"  -->";
        let result = parse_include_directive(line);
        assert_eq!(result, Some((IncludeType::File, "test.asp".to_string())));
    }

    #[test]
    fn test_non_include_line() {
        let line = "<html><body>Test</body></html>";
        let result = parse_include_directive(line);
        assert_eq!(result, None);
    }
}
