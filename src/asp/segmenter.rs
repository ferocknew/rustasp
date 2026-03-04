//! ASP 代码分段器

/// ASP 代码段
#[derive(Debug, Clone, PartialEq)]
pub enum Segment {
    /// HTML 文本
    Html(String),
    /// 代码块 `<% ... %>`
    Code(String),
    /// 表达式块 `<%= ... %>`
    Expr(String),
}

/// 分段器
pub struct Segmenter;

impl Segmenter {
    /// 创建分段器
    pub fn new() -> Self {
        Segmenter
    }

    /// 将 ASP 源代码分段
    pub fn segment(&self, source: &str) -> Vec<Segment> {
        let mut segments = Vec::new();
        let mut chars = source.chars().peekable();
        let mut current_text = String::new();

        while let Some(c) = chars.next() {
            // 检查 `<%` 开始
            if c == '<' && chars.peek() == Some(&'%') {
                chars.next(); // 消费 `%`

                // 保存之前的 HTML
                if !current_text.is_empty() {
                    segments.push(Segment::Html(current_text.clone()));
                    current_text.clear();
                }

                // 检查是否是表达式 `<%=`
                let is_expr = if chars.peek() == Some(&'=') {
                    chars.next(); // 消费 `=`
                    true
                } else {
                    false
                };

                // 收集代码直到 `%>`
                let mut code = String::new();
                while let Some(c) = chars.next() {
                    if c == '%' && chars.peek() == Some(&'>') {
                        chars.next(); // 消费 `>`
                        break;
                    }
                    code.push(c);
                }

                if is_expr {
                    segments.push(Segment::Expr(code.trim().to_string()));
                } else {
                    segments.push(Segment::Code(code.trim().to_string()));
                }
            } else {
                current_text.push(c);
            }
        }

        // 保存最后的 HTML
        if !current_text.is_empty() {
            segments.push(Segment::Html(current_text));
        }

        segments
    }
}

impl Default for Segmenter {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_segment_html() {
        let segmenter = Segmenter::new();
        let segments = segmenter.segment("<h1>Hello</h1>");
        assert_eq!(segments, vec![Segment::Html("<h1>Hello</h1>".to_string())]);
    }

    #[test]
    fn test_segment_code() {
        let segmenter = Segmenter::new();
        let segments = segmenter.segment("<% Response.Write(\"Hello\") %>");
        assert_eq!(segments, vec![Segment::Code("Response.Write(\"Hello\")".to_string())]);
    }

    #[test]
    fn test_segment_expr() {
        let segmenter = Segmenter::new();
        let segments = segmenter.segment("<%= name %>");
        assert_eq!(segments, vec![Segment::Expr("name".to_string())]);
    }

    #[test]
    fn test_segment_mixed() {
        let segmenter = Segmenter::new();
        let segments = segmenter.segment("<html><% Response.Write(\"Hi\") %></html>");
        assert_eq!(segments, vec![
            Segment::Html("<html>".to_string()),
            Segment::Code("Response.Write(\"Hi\")".to_string()),
            Segment::Html("</html>".to_string()),
        ]);
    }
}
