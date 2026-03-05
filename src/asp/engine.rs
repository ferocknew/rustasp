//! ASP 完整引擎
//!
//! 使用 Parser + Runtime 执行 ASP 代码

use super::segmenter::{segment_with_pos, Segment};
use std::collections::HashMap;
use vbscript::ast::Program;
use vbscript::parser::tokenize;
use vbscript::parser::Parser;

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

            // 构建请求参数数据（合并 QueryString 和 Form，支持多值）
            let mut request_data = HashMap::new();

            // 合并 QueryString（Form 数据优先级更高，后处理）
            for (key, values) in &ctx.query_string {
                if let Some(first_value) = values.first() {
                    // 注入第一个值为变量
                    context.define_var(
                        key.clone(),
                        vbscript::runtime::Value::String(first_value.clone()),
                    );
                }
                request_data.insert(key.to_lowercase(), values.clone());
            }

            // 合并 Form 数据（覆盖 QueryString）
            for (key, values) in &ctx.form_data {
                if let Some(first_value) = values.first() {
                    // 注入第一个值为变量
                    context.define_var(
                        key.clone(),
                        vbscript::runtime::Value::String(first_value.clone()),
                    );
                }
                request_data.insert(key.to_lowercase(), values.clone());
            }

            // 设置请求数据（用于 Request("key") 语法）
            context.set_request_data(request_data);
        }

        // 4. 执行每个代码段
        let mut output = String::new();

        for seg in &segments_with_pos {
            eprintln!("DEBUG: Segment at line {}: {:?}", seg.start_line, seg.segment);
            match &seg.segment {
                Segment::Html(html) => {
                    output.push_str(html);
                }
                Segment::Code(code) => {
                    eprintln!("DEBUG ASP: 开始解析代码段: '{}'", code);
                    // 词法分析
                    let tokens = tokenize(code).map_err(|e| {
                        let error_msg = format_error_with_context(
                            "Lexer error",
                            &e.to_string(),
                            code,
                            source,
                            seg.start_line,
                        );
                        eprintln!("❌ ASP Error in {}:\n{}\n", 
                            self.request_context.as_ref().map(|ctx| ctx.path.as_str()).unwrap_or("unknown"),
                            error_msg
                        );
                        error_msg
                    })?;

                    // 语法分析
                    let mut parser = Parser::new(tokens);
                    let stmts = parser.parse_program().map_err(|e| {
                        let error_msg = format_error_with_context(
                            "Parser error",
                            &e.to_string(),
                            code,
                            source,
                            seg.start_line,
                        );
                        eprintln!("❌ ASP Error in {}:\n{}\n", 
                            self.request_context.as_ref().map(|ctx| ctx.path.as_str()).unwrap_or("unknown"),
                            error_msg
                        );
                        error_msg
                    })?;

                    // 执行
                    let program = Program { statements: stmts };
                    interpreter.execute(&program).map_err(|e| {
                        let error_msg = format_error_with_context(
                            "Runtime error",
                            &e.to_string(),
                            code,
                            source,
                            seg.start_line,
                        );
                        eprintln!("❌ ASP Error in {}:\n{}\n", 
                            self.request_context.as_ref().map(|ctx| ctx.path.as_str()).unwrap_or("unknown"),
                            error_msg
                        );
                        error_msg
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
                        let error_msg = format_error_with_context(
                            "Lexer error",
                            &e.to_string(),
                            expr,
                            source,
                            seg.start_line,
                        );
                        eprintln!("❌ ASP Error in {}:\n{}\n", 
                            self.request_context.as_ref().map(|ctx| ctx.path.as_str()).unwrap_or("unknown"),
                            error_msg
                        );
                        error_msg
                    })?;

                    let mut parser = Parser::new(tokens);
                    let ast = parser.parse_expr(0).map_err(|e| {
                        let error_msg = format_error_with_context(
                            "Parser error",
                            &e.to_string(),
                            expr,
                            source,
                            seg.start_line,
                        );
                        eprintln!("❌ ASP Error in {}:\n{}\n", 
                            self.request_context.as_ref().map(|ctx| ctx.path.as_str()).unwrap_or("unknown"),
                            error_msg
                        );
                        error_msg
                    })?;

                    // 求值
                    let value = interpreter.eval_expr(&ast).map_err(|e| {
                        let error_msg = format_error_with_context(
                            "Runtime error",
                            &e.to_string(),
                            expr,
                            source,
                            seg.start_line,
                        );
                        eprintln!("❌ ASP Error in {}:\n{}\n", 
                            self.request_context.as_ref().map(|ctx| ctx.path.as_str()).unwrap_or("unknown"),
                            error_msg
                        );
                        error_msg
                    })?;

                    // 输出
                    use vbscript::runtime::ValueConversion;
                    output.push_str(&ValueConversion::to_string(&value));
                }
                Segment::Directive(_) => {
                    // 指令暂不处理，直接跳过
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

/// 格式化错误信息，包含源文件上下文
fn format_error_with_context(
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

/// 清理错误消息，提取关键信息
fn clean_error_message(message: &str) -> String {
    let msg = message.to_string();
    
    // 移除技术性前缀
    let msg = msg.replace("Parser error: ", "")
                 .replace("Lexer error: ", "")
                 .replace("Runtime error: ", "");
    
    // 将英文错误翻译成中文
    let msg = msg.replace("Unexpected token in expression:", "表达式中出现意外的标记:")
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
