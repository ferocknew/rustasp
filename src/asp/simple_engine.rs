//! 简化的 ASP 引擎（不依赖 chumsky 解析器）

use std::collections::HashMap;

/// 简化的 ASP 引擎
pub struct SimpleEngine {
    variables: HashMap<String, String>,
    debug: bool,
}

impl SimpleEngine {
    pub fn new() -> Self {
        SimpleEngine {
            variables: HashMap::new(),
            debug: false,
        }
    }

    /// 设置调试模式
    pub fn with_debug(mut self, debug: bool) -> Self {
        self.debug = debug;
        self
    }

    /// 执行 ASP 文件（简化版）
    pub fn execute(&mut self, source: &str) -> Result<String, String> {
        if self.debug {
            eprintln!("=== ASP 解析开始 ===");
            eprintln!("源文件长度: {} 字节", source.len());
        }

        let mut output = String::new();
        let mut pos = 0;
        let mut segment_count = 0;

        // 使用正则表达式查找 <% %> 标签
        while let Some(start) = source[pos..].find("<%") {
            segment_count += 1;

            // 添加前面的 HTML
            let html_part = &source[pos..pos + start];
            if !html_part.is_empty() {
                if self.debug {
                    eprintln!("[分段 #{}] HTML: {} 字节", segment_count, html_part.len());
                }
                output.push_str(html_part);
            }

            let tag_start = pos + start + 2;
            pos = tag_start;

            // 查找结束标签 %>
            if let Some(end) = source[pos..].find("%>") {
                let code = &source[pos..pos + end];
                pos += end + 2;

                if self.debug {
                    eprintln!("[分段 #{}] 代码: {:?}", segment_count, code.trim());
                }

                // 执行代码
                match self.eval_code(code, segment_count) {
                    Ok(result) => {
                        if self.debug && !result.is_empty() {
                            eprintln!("  → 输出: {:?}", result);
                        }
                        output.push_str(&result);
                    }
                    Err(e) => {
                        eprintln!("  ❌ 错误: {}", e);
                        output.push_str(&format!("[Error: {}]", e));
                    }
                }
            } else {
                return Err("Unclosed <% tag".to_string());
            }
        }

        // 添加剩余内容
        let remaining = &source[pos..];
        if !remaining.is_empty() {
            if self.debug {
                eprintln!("[分段 #{}] HTML: {} 字节", segment_count + 1, remaining.len());
            }
            output.push_str(remaining);
        }

        if self.debug {
            eprintln!("=== ASP 解析完成 ===");
            eprintln!("总分段数: {}", segment_count);
            eprintln!("输出长度: {} 字节", output.len());
        }

        Ok(output)
    }

    /// 执行简单的代码
    fn eval_code(&mut self, code: &str, segment_id: usize) -> Result<String, String> {
        let code = code.trim();

        if self.debug {
            eprintln!("  [代码块 #{}] 分析: {:?}", segment_id, code);
        }

        // 处理 <%= expr %> 表达式
        if code.starts_with('=') {
            let expr = code[1..].trim();
            if self.debug {
                eprintln!("    类型: 表达式输出");
                eprintln!("    表达式: {:?}", expr);
            }
            return Ok(format!("[{}]", expr));
        }

        // 处理 Response.Write("xxx")
        if code.contains("Response.Write") {
            if let Some(start) = code.find('(') {
                if let Some(end) = code.rfind(')') {
                    let content = &code[start + 1..end];
                    let content = content.trim_matches('"');
                    if self.debug {
                        eprintln!("    类型: Response.Write");
                        eprintln!("    内容: {:?}", content);
                    }
                    return Ok(content.to_string());
                }
            }
        }

        // 简单的变量赋值: Dim x = "value"
        if code.starts_with("Dim") {
            let rest = code[3..].trim();
            if let Some(eq_pos) = rest.find('=') {
                let name = rest[..eq_pos].trim().to_string();
                let value = rest[eq_pos + 1..].trim().trim_matches('"').to_string();
                if self.debug {
                    eprintln!("    类型: 变量声明");
                    eprintln!("    变量: {:?} = {:?}", name, value);
                }
                self.variables.insert(name, value);
                return Ok(String::new());
            }
        }

        if self.debug {
            eprintln!("    类型: 未知（忽略）");
        }

        Ok(String::new())
    }
}

impl Default for SimpleEngine {
    fn default() -> Self {
        Self::new()
    }
}
