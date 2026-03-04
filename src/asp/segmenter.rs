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

/// 带位置信息的代码段
#[derive(Debug, Clone)]
pub struct SegmentWithPos {
    /// 代码段内容
    pub segment: Segment,
    /// 在源文件中的起始行号（1-indexed）
    pub start_line: usize,
    /// 在源文件中的结束行号（1-indexed）
    pub end_line: usize,
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
        let segments_with_pos = self.segment_with_pos()?;
        Ok(segments_with_pos.into_iter().map(|s| s.segment).collect())
    }

    /// 分割源代码为带位置信息的段列表
    pub fn segment_with_pos(&mut self) -> Result<Vec<SegmentWithPos>, String> {
        let mut segments = Vec::new();

        // 计算行号
        let line_starts: Vec<usize> = std::iter::once(0)
            .chain(self.source.match_indices('\n').map(|(i, _)| i + 1))
            .collect();

        let get_line_number = |pos: usize| -> usize {
            for (idx, &start) in line_starts.iter().enumerate() {
                if start > pos {
                    return idx;
                }
            }
            line_starts.len()
        };

        while self.pos < self.source.len() {
            // 查找下一个 <% 标签
            if let Some(start) = self.source[self.pos..].find("<%") {
                let abs_start = self.pos + start;

                // 添加前面的 HTML
                if start > 0 {
                    let html = self.source[self.pos..abs_start].to_string();
                    if !html.is_empty() {
                        let html_start_line = get_line_number(self.pos);
                        let html_end_line = get_line_number(abs_start);
                        segments.push(SegmentWithPos {
                            segment: Segment::Html(html),
                            start_line: html_start_line,
                            end_line: html_end_line,
                        });
                    }
                }

                self.pos = abs_start + 2; // 跳过 <%

                // 检查是否是表达式 <%= %>
                let is_expr = self.source[self.pos..].starts_with('=');

                if is_expr {
                    self.pos += 1; // 跳过 =
                }

                // 查找结束标签 %>
                if let Some(end) = self.source[self.pos..].find("%>") {
                    let abs_end = self.pos + end;
                    let code = self.source[self.pos..abs_end].to_string();

                    let code_start_line = get_line_number(abs_start);
                    let code_end_line = get_line_number(abs_end);

                    if is_expr {
                        segments.push(SegmentWithPos {
                            segment: Segment::Expr(code.trim().to_string()),
                            start_line: code_start_line,
                            end_line: code_end_line,
                        });
                    } else {
                        segments.push(SegmentWithPos {
                            segment: Segment::Code(code.trim().to_string()),
                            start_line: code_start_line,
                            end_line: code_end_line,
                        });
                    }

                    self.pos = abs_end + 2; // 跳过 %>
                } else {
                    return Err("Unclosed <% tag".to_string());
                }
            } else {
                // 没有更多标签，添加剩余 HTML
                let html = self.source[self.pos..].to_string();
                if !html.is_empty() {
                    let html_start_line = get_line_number(self.pos);
                    let html_end_line = line_starts.len();
                    segments.push(SegmentWithPos {
                        segment: Segment::Html(html),
                        start_line: html_start_line,
                        end_line: html_end_line,
                    });
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

/// 分割 ASP 源代码，返回带位置信息的段列表
pub fn segment_with_pos(source: &str) -> Result<Vec<SegmentWithPos>, String> {
    let mut segmenter = Segmenter::new(source);
    segmenter.segment_with_pos()
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
