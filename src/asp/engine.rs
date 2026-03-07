//! ASP 完整引擎
//!
//! 使用 Parser + Runtime 执行 ASP 代码
//! 支持跨代码块的语句（如 If...Else...End If 分散在多个代码块中）

use super::segmenter::{segment_with_pos, Segment};
use std::collections::HashMap;
use vbscript::ast::Program;
use vbscript::parser::tokenize;
use vbscript::parser::Parser;
use vbscript::runtime::Value;
use vbscript::Response;

use vbscript::builtins::{Session, SessionManager};
use crate::http::RequestContext;

/// ASP 执行结果
pub struct ExecutionResult {
    /// 输出内容
    pub output: String,
    /// Response 对象（包含状态码、ContentType、Headers 等）
    pub response: Response,
}

/// 将 Session 转换为 HashMap
fn session_to_map(session: &Session) -> HashMap<String, Value> {
    let mut map = HashMap::new();
    // 存储 Session ID
    map.insert("sessionid".to_string(), Value::String(session.session_id().to_string()));
    map.insert("timeout".to_string(), Value::Number(session.timeout() as f64));
    map
}

/// ASP 执行引擎
pub struct Engine {
    /// 调试模式
    debug: bool,
    /// 请求上下文
    request_context: Option<RequestContext>,
    /// Session 管理器
    session_manager: Option<SessionManager>,
}

impl Engine {
    /// 创建新引擎
    pub fn new() -> Self {
        Engine {
            debug: false,
            request_context: None,
            session_manager: None,
        }
    }

    /// 设置 Session 管理器
    pub fn with_session_manager(mut self, manager: SessionManager) -> Self {
        self.session_manager = Some(manager);
        self
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
    pub fn execute(&mut self, source: &str) -> Result<ExecutionResult, String> {
        // 1. 分割代码（带位置信息）
        let segments_with_pos = segment_with_pos(source)?;

        // 2. 构建完整的 VBScript 程序
        let full_program = self.build_full_program(&segments_with_pos)?;

        // 3. 词法分析
        let tokens = tokenize(&full_program).map_err(|e| {
            let error_msg = format!("Lexer error: {}", e);
            eprintln!("❌ ASP Error: {}\n", error_msg);
            error_msg
        })?;

        // 4. 语法分析
        let mut parser = Parser::new(tokens);
        let stmts = parser.parse_program().map_err(|e| {
            let error_msg = format!("Parser error: {}", e);
            eprintln!("❌ ASP Error: {}\n", error_msg);
            error_msg
        })?;

        // 5. 创建解释器和 Response 对象
        let mut interpreter = vbscript::runtime::Interpreter::new();
        // Response 对象将在 Context 中首次使用时自动创建

        // 5.1 初始化 Session
        let session_id = if let Some(ref ctx) = self.request_context {
            // 从 Cookie 中读取 Session ID
            if let Some(existing_id) = ctx.cookie("ASPSESSIONID") {
                existing_id.to_string()
            } else {
                // 生成新的 Session ID
                SessionManager::generate_session_id()
            }
        } else {
            // 测试模式：从环境变量读取或生成新的
            std::env::var("TEST_SESSION_ID")
                .unwrap_or_else(|_| SessionManager::generate_session_id())
        };

        // 尝试从 SessionManager 加载或创建新的 Session
        let session_id_for_session = session_id.clone();
        let mut session = if let Some(ref mut manager) = self.session_manager {
            match manager.load_session(&session_id) {
                Ok(Some(s)) => s,
                Ok(None) => {
                    // 创建新的 Session
                    match manager.create_session(session_id.clone(), 20) {
                        Ok(s) => s,
                        Err(e) => {
                            eprintln!("警告: 无法创建 Session: {}", e);
                            Session::new(session_id_for_session)
                        }
                    }
                }
                Err(e) => {
                    eprintln!("警告: 无法加载 Session: {}", e);
                    Session::new(session_id_for_session)
                }
            }
        } else {
            // 如果没有 SessionManager，使用内存中的 Session
            Session::new(session_id_for_session)
        };

        // 创建 Session 存储对象 - 使用 HashMap 包装
        let mut session_map: std::collections::HashMap<String, Value> = std::collections::HashMap::new();

        // 从 Session 对象中复制已有数据
        for (key, value) in session.get_all_data() {
            session_map.insert(key, value);
        }

        // 将 Session ID 和 Timeout 作为特殊属性存储
        session_map.insert("__session_id".to_string(), Value::String(session_id.clone()));
        session_map.insert("sessionid".to_string(), Value::String(session.session_id().to_string()));
        session_map.insert("timeout".to_string(), Value::Number(session.timeout() as f64));

        interpreter.context_mut().define_var(
            "Session".to_string(),
            Value::Object(session_map),
        );

        // 6. 注入内建对象（在执行前）
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

        // 7. 执行完整的程序
        let program = Program { statements: stmts };
        interpreter.execute(&program).map_err(|e| {
            let error_msg = format!("Runtime error: {}", e);
            eprintln!("❌ ASP Error: {}\n", error_msg);
            error_msg
        })?;

        // 7.5 保存 Session 数据（如果在执行过程中被修改）
        if let Some(Value::Object(session_map)) = interpreter.context().get_var("Session") {
            // 先将 context 中的 session_map 数据复制到 Session 对象（以便后续使用）
            for (key, value) in session_map {
                // 跳过特殊属性
                if key.starts_with("__") || key == "sessionid" || key == "timeout" {
                    continue;
                }
                session.set(key.clone(), value.clone());
            }

            // 从 context 中的 session_map 创建新的 SessionData 并保存
            let now = std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap_or_default()
                .as_secs();

            // 将 HashMap<String, Value> 转换为 HashMap<String, serde_json::Value>
            let mut data = std::collections::HashMap::new();
            for (key, value) in session_map {
                // 跳过特殊属性
                if key.starts_with("__") || key == "sessionid" || key == "timeout" {
                    continue;
                }
                // 简化处理：只支持基本类型
                let json_value = match value {
                    Value::String(s) => serde_json::Value::String(s.clone()),
                    Value::Number(n) => {
                        if n.fract() == 0.0 && *n >= i64::MIN as f64 && *n <= i64::MAX as f64 {
                            serde_json::Value::Number(serde_json::Number::from(*n as i64))
                        } else {
                            serde_json::Value::String(n.to_string())
                        }
                    }
                    Value::Boolean(b) => serde_json::Value::Bool(*b),
                    _ => serde_json::Value::Null,
                };
                data.insert(key.clone(), json_value);
            }

            // 创建 SessionData
            let session_data = vbscript::builtins::session_manager::SessionData {
                session_id: session.session_id().to_string(),
                timeout: session.timeout(),
                created_at: now - 100, // 假设创建于 100 秒前
                last_accessed: now,
                data,
            };

            // 保存到 SessionManager
            if let Some(ref mut manager) = self.session_manager {
                if let Err(e) = manager.save_session_data(&session_data) {
                    eprintln!("警告: 无法保存 Session: {}", e);
                }
            }
        }

        // 8. 收集输出和 Response 对象
        let mut response = interpreter.context_mut().take_response().unwrap_or_default();

        // 8.1 将 Response.buffer 合并到输出中
        let response_buffer = response.get_output().to_string();
        let context_output = interpreter.context().get_output().to_string();

        let output = if !response_buffer.is_empty() && !context_output.is_empty() {
            format!("{}{}", response_buffer, context_output)
        } else if !response_buffer.is_empty() {
            response_buffer
        } else {
            context_output
        };

        // 8.1 设置 Session Cookie（如果 Session ID 存在且与请求中的不同）
        if let Some(ref ctx) = self.request_context {
            let cookie_id = ctx.cookie("ASPSESSIONID");
            // 如果请求中没有 Cookie，或者 Cookie ID 与当前 Session ID 不一致，则设置新的 Cookie
            if cookie_id != Some(&session_id) {
                response.add_header(
                    "Set-Cookie",
                    &format!("ASPSESSIONID={}; Path=/; HttpOnly", &session_id)
                );
            }
        }

        Ok(ExecutionResult { output, response })
    }

    /// 构建完整的 VBScript 程序
    /// 将所有段（HTML、代码、表达式）合并成一个完整的 VBScript 程序
    fn build_full_program(&self, segments: &[super::segmenter::SegmentWithPos]) -> Result<String, String> {
        let mut program = String::new();

        for seg in segments {
            match &seg.segment {
                Segment::Html(html) => {
                    // HTML 段转换为 Response.Write 语句
                    // 转义双引号
                    let escaped_html = html.replace('\\', "\\\\").replace('"', "\"\"");
                    // 处理多行 HTML
                    for (i, line) in escaped_html.lines().enumerate() {
                        if i > 0 {
                            program.push_str("Response.Write vbNewLine\n");
                        }
                        program.push_str(&format!("Response.Write \"{}\"\n", line));
                    }
                }
                Segment::Code(code) => {
                    // 代码段直接添加
                    program.push_str(code);
                    program.push('\n');
                }
                Segment::Expr(expr) => {
                    // 表达式段转换为 Response.Write 语句
                    program.push_str(&format!("Response.Write {}\n", expr));
                }
                Segment::Directive(_) => {
                    // 指令暂不处理，直接跳过
                }
            }
        }

        Ok(program)
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
