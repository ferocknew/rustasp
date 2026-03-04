//! ASP 代码分段器
//!
//! 将 ASP 文件分割为 HTML 段和代码段

/// 代码段类型
#[derive(Debug, Clone, PartialEq)]
pub enum Segment {
    /// HTML 静态内容
    Html(String),
    /// ASP 代码块 `<% ... %>`
    Code(String),
    /// ASP 表达式 `<%= ... %>`
    Expr(String),
}

/// ASP 分段器
pub struct Segmenter {
    source: String,
    pos: usize,
}

impl Segmenter {
    /// 创建新的分段器
    pub fn new(source: &str) -> Self {
        Segmenter {
            source: source.to_string(),
            pos: 0,
        }
    }

    /// 分割源代码为段列表
    pub fn segment(&mut self) -> Result<Vec<Segment>, String> {
        let mut segments = Vec::new();

        while self.pos < self.source.len() {
            // 查找下一个 <% 标签
            if let Some(start) = self.source[self.pos..].find("<%") {
                // 添加前面的 HTML
                if start > 0 {
                    let html = self.source[self.pos..self.pos + start].to_string();
                    if !html.is_empty() {
                        segments.push(Segment::Html(html));
                    }
                }

                self.pos += start + 2; // 跳过 <%

                // 检查是否是表达式 <%= %>
                let is_expr = self.source[self.pos..].starts_with('=');

                if is_expr {
                    self.pos += 1; // 跳过 =
                }

                // 查找结束标签 %>
                if let Some(end) = self.source[self.pos..].find("%>") {
                    let code = self.source[self.pos..self.pos + end].to_string();

                    if is_expr {
                        segments.push(Segment::Expr(code.trim().to_string()));
                    } else {
                        segments.push(Segment::Code(code.trim().to_string()));
                    }

                    self.pos += end + 2; // 跳过 %>
                } else {
                    return Err("Unclosed <% tag".to_string());
                }
            } else {
                // 没有更多标签，添加剩余 HTML
                let html = self.source[self.pos..].to_string();
                if !html.is_empty() {
                    segments.push(Segment::Html(html));
                }
                break;
            }
        }

        Ok(segments)
    }
}

/// 分割 ASP 源代码（便捷函数）
pub fn segment(source: &str) -> Result<Vec<Segment>, String> {
    let mut segmenter = Segmenter::new(source);
    segmenter.segment()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_html_only() {
        let mut segmenter = Segmenter::new("<html><body>Hello</body></html>");
        let segments = segmenter.segment().unwrap();
        assert_eq!(segments.len(), 1);
        assert!(matches!(&segments[0], Segment::Html(s) if s.contains("Hello")));
    }

    #[test]
    fn test_code_block() {
        let mut segmenter = Segmenter::new("<html><% Response.Write \"Hi\" %></html>");
        let segments = segmenter.segment().unwrap();
        assert_eq!(segments.len(), 3);
        assert!(matches!(&segments[1], Segment::Code(s) if s.contains("Response.Write")));
    }

    #[test]
    fn test_expression() {
        let mut segmenter = Segmenter::new("<html><%= name %></html>");
        let segments = segmenter.segment().unwrap();
        assert_eq!(segments.len(), 3);
        assert!(matches!(&segments[1], Segment::Expr(s) if s == "name"));
    }
}
