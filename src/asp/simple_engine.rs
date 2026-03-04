//! 简化的 ASP 引擎（不依赖 chumsky 解析器）

use std::collections::HashMap;
use crate::http::RequestContext;

/// 简化的 ASP 引擎
pub struct SimpleEngine {
    variables: HashMap<String, String>,
    debug: bool,
    request_context: Option<RequestContext>,
}

impl SimpleEngine {
    pub fn new() -> Self {
        SimpleEngine {
            variables: HashMap::new(),
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
            return self.eval_expression(expr);
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

        // 简单的变量赋值: Dim x = "value" 或 x = Request("key")
        if code.starts_with("Dim") {
            let rest = code[3..].trim();
            if let Some(eq_pos) = rest.find('=') {
                let name = rest[..eq_pos].trim().to_string();
                let value = self.eval_value_expr(rest[eq_pos + 1..].trim());
                if self.debug {
                    eprintln!("    类型: 变量声明");
                    eprintln!("    变量: {:?} = {:?}", name, value);
                }
                self.variables.insert(name, value);
                return Ok(String::new());
            }
        }

        // 处理变量赋值: x = value
        if code.contains('=') && !code.starts_with("If") && !code.starts_with("ElseIf") {
            let parts: Vec<&str> = code.splitn(2, '=').collect();
            if parts.len() == 2 {
                let name = parts[0].trim().to_string();
                let value = self.eval_value_expr(parts[1].trim());
                if self.debug {
                    eprintln!("    类型: 变量赋值");
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

    /// 求值表达式（用于 <%= %> 输出）
    fn eval_expression(&self, expr: &str) -> Result<String, String> {
        // 处理 Request("key")
        if expr.starts_with("Request(") {
            return self.eval_request_call(expr);
        }
        // 处理变量
        if let Some(value) = self.variables.get(expr) {
            return Ok(value.clone());
        }
        // 处理字面量
        if expr.starts_with('"') && expr.ends_with('"') {
            return Ok(expr[1..expr.len()-1].to_string());
        }
        Ok(format!("[{}]", expr))
    }

    /// 求值表达式（用于赋值右侧）
    fn eval_value_expr(&self, expr: &str) -> String {
        // 处理 Request("key")
        if expr.starts_with("Request(") {
            return self.eval_request_call(expr).unwrap_or_default();
        }
        // 处理 Request.QueryString("key")
        if expr.starts_with("Request.QueryString(") {
            return self.eval_request_querystring(expr).unwrap_or_default();
        }
        // 处理 Request.Form("key")
        if expr.starts_with("Request.Form(") {
            return self.eval_request_form(expr).unwrap_or_default();
        }
        // 处理 CInt() 转换
        if expr.starts_with("CInt(") {
            let inner = self.extract_paren_content(expr, "CInt(");
            let inner_val = self.eval_value_expr(&inner);
            return inner_val.parse::<i32>().unwrap_or(0).to_string();
        }
        // 处理变量
        if let Some(value) = self.variables.get(expr) {
            return value.clone();
        }
        // 处理字面量
        if expr.starts_with('"') && expr.ends_with('"') {
            return expr[1..expr.len()-1].to_string();
        }
        expr.to_string()
    }

    /// 提取括号内容
    fn extract_paren_content(&self, expr: &str, prefix: &str) -> String {
        let start = prefix.len();
        if expr[start..].starts_with('(') {
            if let Some(end) = expr[start..].rfind(')') {
                return expr[start + 1..start + end].trim().to_string();
            }
        }
        // 如果没有额外的括号，直接找结束括号
        if let Some(end) = expr[start..].rfind(')') {
            expr[start..end].trim().to_string()
        } else {
            expr[start..].trim().to_string()
        }
    }

    /// 处理 Request("key") 调用
    fn eval_request_call(&self, expr: &str) -> Result<String, String> {
        // 提取参数: Request("key")
        let key = self.extract_paren_content(expr, "Request(");
        let key = key.trim_matches('"');

        if let Some(ref ctx) = self.request_context {
            // 先查找 Form，再查找 QueryString
            if let Some(value) = ctx.form(key) {
                return Ok(value.clone());
            }
            if let Some(value) = ctx.query(key) {
                return Ok(value.clone());
            }
        }
        Ok(String::new())
    }

    /// 处理 Request.QueryString("key") 调用
    fn eval_request_querystring(&self, expr: &str) -> Result<String, String> {
        let key = self.extract_paren_content(expr, "Request.QueryString(");
        let key = key.trim_matches('"');

        if let Some(ref ctx) = self.request_context {
            if let Some(value) = ctx.query(key) {
                return Ok(value.clone());
            }
        }
        Ok(String::new())
    }

    /// 处理 Request.Form("key") 调用
    fn eval_request_form(&self, expr: &str) -> Result<String, String> {
        let key = self.extract_paren_content(expr, "Request.Form(");
        let key = key.trim_matches('"');

        if let Some(ref ctx) = self.request_context {
            if let Some(value) = ctx.form(key) {
                return Ok(value.clone());
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
