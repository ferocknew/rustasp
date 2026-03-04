//! ASP 完整引擎
//!
//! 使用 StmtParser + Runtime 执行 ASP 代码

use super::segmenter::{segment, Segment};
use vbscript::parser::{parse_expression, parse_program, tokenize};
use vbscript::runtime::{Interpreter, Value, ValueConversion};

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
        if self.debug {
            eprintln!("=== ASP 引擎开始 ===");
            eprintln!("源文件长度: {} 字节", source.len());
        }

        // 1. 分割代码
        let segments = segment(source)?;

        if self.debug {
            eprintln!("分段数量: {}", segments.len());
        }

        // 2. 创建解释器
        let mut interpreter = Interpreter::new();

        // 3. 注入内建对象（在执行前）
        if let Some(ref ctx) = self.request_context {
            let context = interpreter.context_mut();
            // 将 QueryString 注入为变量
            for (key, value) in &ctx.query_string {
                context.define_var(key.clone(), Value::String(value.clone()));
            }
            // 将 Form 数据注入为变量
            for (key, value) in &ctx.form_data {
                context.define_var(key.clone(), Value::String(value.clone()));
            }
        }

        // 4. 执行每个代码段
        let mut output = String::new();

        for (i, segment) in segments.iter().enumerate() {
            match segment {
                Segment::Html(html) => {
                    if self.debug {
                        eprintln!("[段 #{}] HTML: {} 字节", i + 1, html.len());
                    }
                    output.push_str(html);
                }
                Segment::Code(code) => {
                    if self.debug {
                        eprintln!("[段 #{}] 代码: {:?}", i + 1, code);
                    }

                    // 词法分析
                    let tokens = tokenize(code)
                        .map_err(|e| format!("Lexer error: {:?}", e))?;

                    // 语法分析
                    let program = parse_program(tokens)
                        .map_err(|e| format!("Parser error: {:?}", e))?;

                    // 执行
                    interpreter.execute(&program)
                        .map_err(|e| format!("Runtime error: {:?}", e))?;

                    // 收集输出
                    {
                        let context = interpreter.context();
                        output.push_str(context.get_output());
                    }
                    interpreter.context_mut().clear_output();
                }
                Segment::Expr(expr) => {
                    if self.debug {
                        eprintln!("[段 #{}] 表达式: {:?}", i + 1, expr);
                    }

                    // 解析并执行表达式
                    let tokens = tokenize(expr)
                        .map_err(|e| format!("Lexer error: {:?}", e))?;

                    let ast = parse_expression(tokens)
                        .map_err(|e| format!("Parser error: {:?}", e))?;

                    // 求值
                    let value = interpreter.eval_expr(&ast)
                        .map_err(|e| format!("Runtime error: {:?}", e))?;

                    // 输出
                    output.push_str(&ValueConversion::to_string(&value));
                }
            }
        }

        if self.debug {
            eprintln!("=== ASP 引擎结束 ===");
            eprintln!("输出长度: {} 字节", output.len());
        }

        Ok(output)
    }
}

impl Default for Engine {
    fn default() -> Self {
        Self::new()
    }
}
