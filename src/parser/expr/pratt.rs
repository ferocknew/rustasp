//! Pratt 算法主入口

use crate::ast::Expr;
use crate::parser::ParseError;
use crate::parser::Parser;

impl Parser {
    /// 解析表达式（入口函数）
    ///
    /// min_bp: 最小绑定优先级，低于此优先级的运算符会停止解析
    pub fn parse_expr(&mut self, min_bp: u8) -> Result<Expr, ParseError> {
        // 1. 解析前缀表达式（左侧）
        let mut lhs = self.parse_prefix()?;

        // 2. 循环处理中缀运算符
        loop {
            // 查看下一个 token 是否是中缀运算符
            let (l_bp, r_bp) = match self.infix_binding_power() {
                Some(bp) => bp,
                None => break,
            };

            // 如果左侧优先级小于要求的最小优先级，停止
            if l_bp < min_bp {
                break;
            }

            // 消耗运算符
            let op_token = self.advance().clone();
            let op = self.token_to_binary_op(&op_token)?;

            // 解析右侧表达式（使用右侧优先级）
            let rhs = self.parse_expr(r_bp)?;

            // 构建二元运算 AST
            lhs = Expr::Binary {
                left: Box::new(lhs),
                op,
                right: Box::new(rhs),
            };
        }

        Ok(lhs)
    }
}
