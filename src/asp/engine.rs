//! ASP 完整引擎
//!
//! 使用 StmtParser + Runtime 执行 ASP 代码

use super::segmenter::{segment, Segment};
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
        // 1. 分割代码
        let segments = segment(source)?;

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

        for (i, segment) in segments.iter().enumerate() {
            match segment {
                Segment::Html(html) => {
                    output.push_str(html);
                }
                Segment::Code(code) => {
                    // 词法分析
                    let tokens = tokenize(code).map_err(|e| {
                        format_error_with_code("Lexer error", &e.to_string(), code, i + 1)
                    })?;

                    // 语法分析
                    let program = parse_program(tokens).map_err(|e| {
                        format_error_with_code("Parser error", &e.to_string(), code, i + 1)
                    })?;

                    // 执行
                    interpreter.execute(&program).map_err(|e| {
                        format_error_with_code("Runtime error", &e.to_string(), code, i + 1)
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
                        format_error_with_code("Lexer error", &e.to_string(), expr, i + 1)
                    })?;

                    let ast = parse_expression(tokens).map_err(|e| {
                        format_error_with_code("Parser error", &e.to_string(), expr, i + 1)
                    })?;

                    // 求值
                    let value = interpreter.eval_expr(&ast).map_err(|e| {
                        format_error_with_code("Runtime error", &e.to_string(), expr, i + 1)
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

/// 格式化错误信息，包含代码段
fn format_error_with_code(error_type: &str, message: &str, code: &str, segment_num: usize) -> String {
    // 截取代码摘要（最多 200 字符）
    let code_preview = if code.len() > 200 {
        format!("{}...", &code[..200])
    } else {
        code.to_string()
    };

    // 转义换行符以便单行显示
    let code_single_line = code_preview
        .replace('\n', "\\n")
        .replace('\r', "\\r")
        .replace('\t', "\\t");

    format!(
        "{}: {}\n[Segment #{} Code]: {}",
        error_type,
        message,
        segment_num,
        code_single_line
    )
}
