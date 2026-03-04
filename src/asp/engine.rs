//! ASP 完整引擎
//!
//! 使用 StmtParser + Runtime 执行 ASP 代码

use super::segmenter::{segment_with_pos, Segment};
use vbscript::parser::{parse_expression, parse_program, tokenize};

use crate::http::RequestContext;

/// ASP 执行引擎
pub struct Engine {
    /// 调试模式
    debug: bool,
    /// 请求上下文
    request_context: Option<RequestContext>,
}

impl Engine {
    /// 创建新引擎
    pub fn new() -> Self {
        Engine {
            debug: false,
            request_context: None,
        }
    }

    /// 设置调试模式
    pub fn with_debug(mut self, debug: bool) -> Self {
        self.debug = debug;
        self
    }

    /// 设置请求上下文
    pub fn with_request_context(mut self, ctx: RequestContext) -> Self {
        self.request_context = Some(ctx);
        self
    }

    /// 执行 ASP 文件
    pub fn execute(&mut self, source: &str) -> Result<String, String> {
        // 1. 分割代码（带位置信息）
        let segments_with_pos = segment_with_pos(source)?;

        // 2. 创建解释器
        let mut interpreter = vbscript::runtime::Interpreter::new();

        // 3. 注入内建对象（在执行前）
        if let Some(ref ctx) = self.request_context {
            let context = interpreter.context_mut();
            // 将 QueryString 注入为变量
            for (key, value) in &ctx.query_string {
                context.define_var(
                    key.clone(),
                    vbscript::runtime::Value::String(value.clone()),
                );
            }
            // 将 Form 数据注入为变量
            for (key, value) in &ctx.form_data {
                context.define_var(
                    key.clone(),
                    vbscript::runtime::Value::String(value.clone()),
                );
            }
        }

        // 4. 执行每个代码段
        let mut output = String::new();

        for seg in &segments_with_pos {
            match &seg.segment {
                Segment::Html(html) => {
                    output.push_str(html);
                }
                Segment::Code(code) => {
                    // 词法分析
                    let tokens = tokenize(code).map_err(|e| {
                        format_error_with_context(
                            "Lexer error",
                            &e.to_string(),
                            code,
                            source,
                            seg.start_line,
                        )
                    })?;

                    // 语法分析
                    let program = parse_program(tokens).map_err(|e| {
                        format_error_with_context(
                            "Parser error",
                            &e.to_string(),
                            code,
                            source,
                            seg.start_line,
                        )
                    })?;

                    // 执行
                    interpreter.execute(&program).map_err(|e| {
                        format_error_with_context(
                            "Runtime error",
                            &e.to_string(),
                            code,
                            source,
                            seg.start_line,
                        )
                    })?;

                    // 收集输出
                    {
                        let context = interpreter.context();
                        output.push_str(context.get_output());
                    }
                    interpreter.context_mut().clear_output();
                }
                Segment::Expr(expr) => {
                    // 解析并执行表达式
                    let tokens = tokenize(expr).map_err(|e| {
                        format_error_with_context(
                            "Lexer error",
                            &e.to_string(),
                            expr,
                            source,
                            seg.start_line,
                        )
                    })?;

                    let ast = parse_expression(tokens).map_err(|e| {
                        format_error_with_context(
                            "Parser error",
                            &e.to_string(),
                            expr,
                            source,
                            seg.start_line,
                        )
                    })?;

                    // 求值
                    let value = interpreter.eval_expr(&ast).map_err(|e| {
                        format_error_with_context(
                            "Runtime error",
                            &e.to_string(),
                            expr,
                            source,
                            seg.start_line,
                        )
                    })?;

                    // 输出
                    use vbscript::runtime::ValueConversion;
                    output.push_str(&ValueConversion::to_string(&value));
                }
            }
        }

        Ok(output)
    }
}

impl Default for Engine {
    fn default() -> Self {
        Self::new()
    }
}

/// 格式化错误信息，包含源文件上下文
fn format_error_with_context(
    error_type: &str,
    message: &str,
    _code: &str,
    source: &str,
    start_line: usize,
) -> String {
    // 获取源文件行
    let source_lines: Vec<&str> = source.lines().collect();

    // 确定错误发生的行（从消息中提取相对于代码段的行号）
    let relative_line = extract_error_line(message).unwrap_or(1);
    let absolute_line = start_line + relative_line - 1;

    // 确定显示范围
    let context_lines = 2;
    let show_start = (absolute_line.saturating_sub(context_lines + 1)).max(0);
    let show_end = (absolute_line + context_lines).min(source_lines.len());

    // 构建代码摘要
    let mut code_summary = String::new();
    for (idx, line) in source_lines
        .iter()
        .enumerate()
        .skip(show_start)
        .take(show_end - show_start)
    {
        let line_num = idx + 1;
        let marker = if line_num == absolute_line { " >>>" } else { "    " };
        code_summary.push_str(&format!("{}{:3} | {}\n", marker, line_num, line));
    }

    // 格式化消息
    let msg = if message.starts_with(error_type) {
        message.to_string()
    } else {
        format!("{}: {}", error_type, message)
    };

    format!(
        "{}\n\nError at line {}:\n{}",
        msg,
        absolute_line,
        code_summary.trim_end()
    )
}

/// 从错误消息中提取行号
fn extract_error_line(message: &str) -> Option<usize> {
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
