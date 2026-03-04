//! ASP 执行引擎

use crate::parser;
use crate::runtime::Interpreter;
use crate::ast::Expr;
use super::segmenter::{Segment, Segmenter};

/// ASP 引擎
pub struct Engine {
    segmenter: Segmenter,
    interpreter: Interpreter,
}

impl Engine {
    /// 创建新引擎
    pub fn new() -> Self {
        Engine {
            segmenter: Segmenter::new(),
            interpreter: Interpreter::new(),
        }
    }

    /// 执行 ASP 文件
    pub fn execute(&mut self, source: &str) -> Result<String, String> {
        // 分段
        let segments = self.segmenter.segment(source);

        // 清空输出
        self.interpreter.context_mut().clear_output();

        // 逐段执行
        for segment in segments {
            match segment {
                Segment::Html(html) => {
                    self.interpreter.context_mut().write(&html);
                }
                Segment::Code(code) => {
                    // 解析并执行代码
                    let program = parser::parse(&code)
                        .map_err(|e| format!("Parse error: {}", e))?;
                    self.interpreter.execute(&program)
                        .map_err(|e| format!("Runtime error: {}", e))?;
                }
                Segment::Expr(expr_code) => {
                    // 解析并执行表达式
                    let expr = parser::parse_expr(&expr_code)
                        .map_err(|e| format!("Parse error: {}", e))?;
                    // 这里需要调用解释器的表达式求值
                    // 暂时简化处理
                    self.interpreter.context_mut().write(&format!("{{{{ {} }}}}", expr_code));
                }
            }
        }

        Ok(self.interpreter.context().get_output().to_string())
    }

    /// 获取解释器（用于注册内置对象）
    pub fn interpreter(&mut self) -> &mut Interpreter {
        &mut self.interpreter
    }
}

impl Default for Engine {
    fn default() -> Self {
        Self::new()
    }
}
