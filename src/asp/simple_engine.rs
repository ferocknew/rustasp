//! 简化的 ASP 引擎（不依赖 chumsky 解析器）

use std::collections::HashMap;

/// 简化的 ASP 引擎
pub struct SimpleEngine {
    variables: HashMap<String, String>,
}

impl SimpleEngine {
    pub fn new() -> Self {
        SimpleEngine {
            variables: HashMap::new(),
        }
    }

    /// 执行 ASP 文件（简化版）
    pub fn execute(&mut self, source: &str) -> Result<String, String> {
        let mut output = String::new();
        let mut pos = 0;

        // 使用正则表达式查找 <% %> 标签
        while let Some(start) = source[pos..].find("<%") {
            // 添加前面的 HTML
            output.push_str(&source[pos..pos + start]);

            let tag_start = pos + start + 2;
            pos = tag_start;

            // 查找结束标签 %>
            if let Some(end) = source[pos..].find("%>") {
                let code = &source[pos..pos + end];
                pos += end + 2;

                // 执行代码
                match self.eval_code(code) {
                    Ok(result) => output.push_str(&result),
                    Err(e) => output.push_str(&format!("[Error: {}]", e)),
                }
            } else {
                return Err("Unclosed <% tag".to_string());
            }
        }

        // 添加剩余内容
        output.push_str(&source[pos..]);
        Ok(output)
    }

    /// 执行简单的代码
    fn eval_code(&mut self, code: &str) -> Result<String, String> {
        let code = code.trim();

        // 处理 <%= expr %> 表达式
        if code.starts_with('=') {
            let expr = code[1..].trim();
            return Ok(format!("[{}]", expr));
        }

        // 处理 Response.Write("xxx")
        if code.contains("Response.Write") {
            if let Some(start) = code.find('(') {
                if let Some(end) = code.rfind(')') {
                    let content = &code[start + 1..end];
                    let content = content.trim_matches('"');
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
                self.variables.insert(name, value);
                return Ok(String::new());
            }
        }

        Ok(String::new())
    }
}

impl Default for SimpleEngine {
    fn default() -> Self {
        Self::new()
    }
}
