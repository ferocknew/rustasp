//! ASP 引擎错误处理
//!
//! 提供格式化的错误信息，包含源代码上下文

/// 格式化错误信息，包含源文件上下文
#[allow(dead_code)]
pub fn format_error_with_context(
    error_type: &str,
    message: &str,
    code: &str,
    source: &str,
    start_line: usize,
) -> String {
    // 获取源文件行
    let source_lines: Vec<&str> = source.lines().collect();

    // 获取代码段行
    let _code_lines: Vec<&str> = code.lines().collect();

    // 确定错误发生的行（从消息中提取相对于代码段的行号）
    let relative_line = extract_error_line(message).unwrap_or(1);
    let absolute_line = start_line + relative_line - 1;

    // 确定显示范围
    let context_lines = 2;
    let show_start = (absolute_line.saturating_sub(context_lines)).max(0);
    let show_end = (absolute_line + context_lines).min(source_lines.len());

    // 构建代码摘要
    let mut code_summary = String::new();

    // 添加错误行的上下文
    for (idx, line) in source_lines
        .iter()
        .enumerate()
        .skip(show_start)
        .take(show_end - show_start)
    {
        let line_num = idx + 1;
        let is_error_line = line_num == absolute_line;
        let marker = if is_error_line { ">>>" } else { "   " };

        code_summary.push_str(&format!("{} {:4} | {}\n", marker, line_num, line));
    }

    // 美化错误类型
    let (error_type_cn, error_icon) = match error_type {
        "Lexer error" => ("词法分析错误", "🔤"),
        "Parser error" => ("语法分析错误", "📝"),
        "Runtime error" => ("运行时错误", "⚙️"),
        _ => (error_type, "❌"),
    };

    // 格式化消息，提取关键信息
    let clean_message = clean_error_message(message);

    format!(
        "{} {} (第 {} 行)\n\n{}\n\n代码上下文:\n{}",
        error_icon,
        error_type_cn,
        absolute_line,
        clean_message,
        code_summary.trim_end()
    )
}

/// 从错误消息中提取行号
#[allow(dead_code)]
pub fn extract_error_line(message: &str) -> Option<usize> {
    let lower = message.to_lowercase();

    // 尝试找 "at line X" 模式
    if let Some(pos) = lower.find("at line") {
        let rest = &message[pos + 7..];
        let rest = rest.trim_start();
        if let Some(num_end) = rest.find(|c: char| !c.is_ascii_digit()) {
            if let Ok(line) = rest[..num_end].parse::<usize>() {
                return Some(line);
            }
        } else if let Ok(line) = rest.parse::<usize>() {
            return Some(line);
        }
    }

    // 尝试找 "line X" 模式
    if let Some(pos) = lower.find("line") {
        let rest = &message[pos + 4..];
        let rest = rest.trim_start();
        if let Some(num_end) = rest.find(|c: char| !c.is_ascii_digit()) {
            if let Ok(line) = rest[..num_end].parse::<usize>() {
                return Some(line);
            }
        }
    }

    None
}

/// 清理错误消息，提取关键信息
#[allow(dead_code)]
fn clean_error_message(message: &str) -> String {
    let msg = message.to_string();

    // 移除技术性前缀
    let msg = msg
        .replace("Parser error: ", "")
        .replace("Lexer error: ", "")
        .replace("Runtime error: ", "");

    // 将英文错误翻译成中文
    let msg = msg
        .replace("Unexpected token in expression:", "表达式中出现意外的标记:")
        .replace("Expected", "期望")
        .replace("found", "但找到")
        .replace("Undefined variable", "未定义的变量")
        .replace("Type mismatch", "类型不匹配")
        .replace("Division by zero", "除零错误")
        .replace("Object required", "需要对象")
        .replace("Property not found", "属性不存在")
        .replace("Method not found", "方法不存在")
        .replace("Invalid assignment", "无效的赋值")
        .replace("at line", "位于第")
        .replace("column", "列");

    msg.trim().to_string()
}
