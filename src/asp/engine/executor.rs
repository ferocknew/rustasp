//! ASP 引擎执行器
//!
//! 负责执行 ASP 代码，管理 Session、Request、Response 等对象

use super::super::segmenter::{Segment, SegmentWithPos, segment_with_pos};
use crate::http::RequestContext;
use std::collections::HashMap;
use std::sync::{Arc, Mutex};
use vbscript::ast::Program;
use vbscript::parser::tokenize;
use vbscript::parser::Parser;
use vbscript::runtime::Value;
use vbscript::runtime::objects::{Session, SessionManager};

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
    pub fn execute(&mut self, source: &str) -> Result<super::ExecutionResult, String> {
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

        // 5. 创建解释器
        let mut interpreter = vbscript::runtime::Interpreter::new();

        // 5.1 创建 Response 对象并放入变量表（与 Request、Session 保持一致）
        let response = vbscript::runtime::objects::Response::new();
        interpreter.context_mut().define_var(
            "Response".to_string(),
            Value::Object(Arc::new(Mutex::new(response))),
        );

        // 5.2 初始化 Session
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
        let session = if let Some(ref mut manager) = self.session_manager {
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

        // 将 Session 对象作为 BuiltinObject 存储
        interpreter.context_mut().define_var(
            "Session".to_string(),
            Value::Object(Arc::new(Mutex::new(session))),
        );

        // 5.2 定义 VBScript 内置常量
        Self::define_vbscript_constants(&mut interpreter);

        // 6. 注入内建对象（在执行前）
        if let Some(ref ctx) = self.request_context {
            let context = interpreter.context_mut();

            // 创建 Request 对象
            let mut request = vbscript::runtime::objects::Request::new();

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
                // 设置所有值到 Request 对象（支持多值访问）
                request.set_query_string_multiple(key.clone(), values.clone());
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
                // 设置所有值到 Request 对象（支持多值访问）
                request.set_form_multiple(key.clone(), values.clone());
                request_data.insert(key.to_lowercase(), values.clone());
            }

            // 将 Request 对象作为变量存储
            context.define_var(
                "Request".to_string(),
                vbscript::runtime::Value::Object(Arc::new(Mutex::new(request))),
            );

            // 设置请求数据（用于兼容性）
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
        if let Some(Value::Object(ref session_obj)) = interpreter.context().get_var("Session") {
            // 从 BuiltinObject 中获取 Session 对象
            #[allow(unused_imports)]
            use vbscript::runtime::BuiltinObject;
            if let Ok(session_guard) = session_obj.lock() {
                if let Some(session) = session_guard.as_any().downcast_ref::<Session>() {
                    // 创建 SessionData 并保存
                    match session.to_session_data() {
                        Ok(session_data) => {
                            // 保存到 SessionManager
                            if let Some(ref mut manager) = self.session_manager {
                                if let Err(e) = manager.save_session_data(&session_data) {
                                    eprintln!("警告: 无法保存 Session: {}", e);
                                }
                            }
                        }
                        Err(e) => {
                            eprintln!("警告: 无法序列化 Session: {}", e);
                        }
                    }
                }
            }
        }

        // 8. 收集输出和 Response 对象
        // 从变量表中获取 Response 对象
        let mut response = if let Some(Value::Object(ref response_obj)) = interpreter.context().get_var("Response") {
            // 从 BuiltinObject 中获取 Response
            #[allow(unused_imports)]
            use vbscript::runtime::BuiltinObject;
            if let Ok(resp_guard) = response_obj.lock() {
                if let Some(resp) = resp_guard.as_any().downcast_ref::<vbscript::runtime::objects::Response>() {
                    resp.clone()
                } else {
                    vbscript::runtime::objects::Response::new()
                }
            } else {
                vbscript::runtime::objects::Response::new()
            }
        } else {
            vbscript::runtime::objects::Response::new()
        };

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

        Ok(super::ExecutionResult { output, response })
    }

    /// 定义 VBScript 内置常量
    fn define_vbscript_constants(interpreter: &mut vbscript::runtime::Interpreter) {
        let context = interpreter.context_mut();

        // ===== MsgBox 返回值常量 =====
        context.define_var("vbOK".to_string(), Value::Number(1.0));
        context.define_var("vbCancel".to_string(), Value::Number(2.0));
        context.define_var("vbAbort".to_string(), Value::Number(3.0));
        context.define_var("vbRetry".to_string(), Value::Number(4.0));
        context.define_var("vbIgnore".to_string(), Value::Number(5.0));
        context.define_var("vbYes".to_string(), Value::Number(6.0));
        context.define_var("vbNo".to_string(), Value::Number(7.0));

        // ===== MsgBox 按钮常量 =====
        context.define_var("vbOKOnly".to_string(), Value::Number(0.0));
        context.define_var("vbOKCancel".to_string(), Value::Number(1.0));
        context.define_var("vbAbortRetryIgnore".to_string(), Value::Number(2.0));
        context.define_var("vbYesNoCancel".to_string(), Value::Number(3.0));
        context.define_var("vbYesNo".to_string(), Value::Number(4.0));
        context.define_var("vbRetryCancel".to_string(), Value::Number(5.0));

        // ===== MsgBox 图标常量 =====
        context.define_var("vbCritical".to_string(), Value::Number(16.0));
        context.define_var("vbQuestion".to_string(), Value::Number(32.0));
        context.define_var("vbExclamation".to_string(), Value::Number(48.0));
        context.define_var("vbInformation".to_string(), Value::Number(64.0));

        // ===== MsgBox 默认按钮常量 =====
        context.define_var("vbDefaultButton1".to_string(), Value::Number(0.0));
        context.define_var("vbDefaultButton2".to_string(), Value::Number(256.0));
        context.define_var("vbDefaultButton3".to_string(), Value::Number(512.0));
        context.define_var("vbDefaultButton4".to_string(), Value::Number(768.0));

        // ===== MsgBox 模式常量 =====
        context.define_var("vbApplicationModal".to_string(), Value::Number(0.0));
        context.define_var("vbSystemModal".to_string(), Value::Number(4096.0));

        // ===== VarType 常量 =====
        context.define_var("vbEmpty".to_string(), Value::Number(0.0));
        context.define_var("vbNull".to_string(), Value::Number(1.0));
        context.define_var("vbInteger".to_string(), Value::Number(2.0));
        context.define_var("vbLong".to_string(), Value::Number(3.0));
        context.define_var("vbSingle".to_string(), Value::Number(4.0));
        context.define_var("vbDouble".to_string(), Value::Number(5.0));
        context.define_var("vbCurrency".to_string(), Value::Number(6.0));
        context.define_var("vbDate".to_string(), Value::Number(7.0));
        context.define_var("vbString".to_string(), Value::Number(8.0));
        context.define_var("vbObject".to_string(), Value::Number(9.0));
        context.define_var("vbError".to_string(), Value::Number(10.0));
        context.define_var("vbBoolean".to_string(), Value::Number(11.0));
        context.define_var("vbVariant".to_string(), Value::Number(12.0));
        context.define_var("vbDataObject".to_string(), Value::Number(13.0));
        context.define_var("vbDecimal".to_string(), Value::Number(14.0));
        context.define_var("vbByte".to_string(), Value::Number(17.0));
        context.define_var("vbArray".to_string(), Value::Number(8192.0));

        // ===== 三态常量 =====
        context.define_var("vbUseDefault".to_string(), Value::Number(-2.0));
        context.define_var("vbTrue".to_string(), Value::Boolean(true));
        context.define_var("vbFalse".to_string(), Value::Boolean(false));

        // ===== 比较常量 =====
        context.define_var("vbBinaryCompare".to_string(), Value::Number(0.0));
        context.define_var("vbTextCompare".to_string(), Value::Number(1.0));
        context.define_var("vbDatabaseCompare".to_string(), Value::Number(2.0));

        // ===== 日期格式常量 =====
        context.define_var("vbGeneralDate".to_string(), Value::Number(0.0));
        context.define_var("vbLongDate".to_string(), Value::Number(1.0));
        context.define_var("vbShortDate".to_string(), Value::Number(2.0));
        context.define_var("vbLongTime".to_string(), Value::Number(3.0));
        context.define_var("vbShortTime".to_string(), Value::Number(4.0));

        // ===== 星期常量 =====
        context.define_var("vbUseSystemDayOfWeek".to_string(), Value::Number(0.0));
        context.define_var("vbSunday".to_string(), Value::Number(1.0));
        context.define_var("vbMonday".to_string(), Value::Number(2.0));
        context.define_var("vbTuesday".to_string(), Value::Number(3.0));
        context.define_var("vbWednesday".to_string(), Value::Number(4.0));
        context.define_var("vbThursday".to_string(), Value::Number(5.0));
        context.define_var("vbFriday".to_string(), Value::Number(6.0));
        context.define_var("vbSaturday".to_string(), Value::Number(7.0));

        // ===== 年周常量 =====
        context.define_var("vbFirstJan1".to_string(), Value::Number(1.0));
        context.define_var("vbFirstFourDays".to_string(), Value::Number(2.0));
        context.define_var("vbFirstFullWeek".to_string(), Value::Number(3.0));

        // ===== 颜色常量 =====
        context.define_var("vbBlack".to_string(), Value::Number(0.0));
        context.define_var("vbBlue".to_string(), Value::Number(16711680.0));      // 0xFF0000
        context.define_var("vbCyan".to_string(), Value::Number(16776960.0));      // 0xFFFF00
        context.define_var("vbGreen".to_string(), Value::Number(65280.0));        // 0x00FF00
        context.define_var("vbMagenta".to_string(), Value::Number(16711935.0));   // 0xFF00FF
        context.define_var("vbRed".to_string(), Value::Number(255.0));            // 0x0000FF
        context.define_var("vbWhite".to_string(), Value::Number(16777215.0));     // 0xFFFFFF
        context.define_var("vbYellow".to_string(), Value::Number(65535.0));       // 0x00FFFF

        // ===== 字符串常量 =====
        context.define_var("vbCr".to_string(), Value::String("\r".to_string()));
        context.define_var("vbCrLf".to_string(), Value::String("\r\n".to_string()));
        context.define_var("vbNewLine".to_string(), Value::String("\r\n".to_string()));
        context.define_var("vbFormFeed".to_string(), Value::String("\x0C".to_string()));
        context.define_var("vbLf".to_string(), Value::String("\n".to_string()));
        context.define_var("vbNullChar".to_string(), Value::String("\x00".to_string()));
        context.define_var("vbNullString".to_string(), Value::String(String::new()));
        context.define_var("vbTab".to_string(), Value::String("\t".to_string()));
        context.define_var("vbVerticalTab".to_string(), Value::String("\x0B".to_string()));

        // ===== 其他常量 =====
        context.define_var("vbUseSystem".to_string(), Value::Number(0.0));
        context.define_var("vbObjectError".to_string(), Value::Number(-2147221504.0));
    }

    /// 构建完整的 VBScript 程序
    /// 将所有段（HTML、代码、表达式）合并成一个完整的 VBScript 程序
    fn build_full_program(&self, segments: &[SegmentWithPos]) -> Result<String, String> {
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
