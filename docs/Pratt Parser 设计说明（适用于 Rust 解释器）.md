# Pratt Parser 设计说明（适用于 Rust 解释器）

本方案用于替代 chumsky，实现一个编译稳定、可控、适合解释器的表达式解析器。

---

## 一、为什么选择 Pratt Parser

* 编译时间稳定（无泛型爆炸问题）
* 实现简单，可控性高
* 非常适合表达式解析（VBScript 风格）
* 易扩展运算符优先级

---

## 二、核心思想

每个运算符定义“绑定优先级（binding power）”。

解析函数：

```
parse_expression(min_bp)
```

流程：

1. 解析前缀表达式（数字、变量、一元运算、括号）
2. 循环处理所有优先级 >= min_bp 的中缀运算符
3. 构建 AST

---

## 三、基础结构

### AST

```rust
pub enum Expr {
    Number(f64),
    String(String),
    Variable(String),

    Unary {
        op: UnaryOp,
        rhs: Box<Expr>,
    },

    Binary {
        lhs: Box<Expr>,
        op: BinaryOp,
        rhs: Box<Expr>,
    },
}
```

---

### 绑定优先级示例

```rust
fn infix_binding_power(tok: &Token) -> Option<(u8, u8)> {
    match tok {
        Token::Or => Some((1, 2)),
        Token::And => Some((3, 4)),
        Token::Plus | Token::Minus => Some((7, 8)),
        Token::Star | Token::Slash => Some((9, 10)),
        _ => None,
    }
}
```

---

## 四、核心解析函数

```rust
pub fn parse_expression(&mut self, min_bp: u8) -> Expr {
    let mut lhs = parse_prefix();

    loop {
        let (l_bp, r_bp) = match infix_binding_power(self.peek()) {
            Some(bp) => bp,
            None => break,
        };

        if l_bp < min_bp {
            break;
        }

        let op = self.next();
        let rhs = self.parse_expression(r_bp);

        lhs = Expr::Binary {
            lhs: Box::new(lhs),
            op: token_to_binary_op(op),
            rhs: Box::new(rhs),
        };
    }

    lhs
}
```

---

## 五、推荐目录结构

```
src/parser/
 ├── token.rs
 ├── lexer.rs
 ├── ast.rs
 └── expr.rs  (Pratt 实现)
```

---

## 六、总结

Pratt Parser：

* 无泛型链式嵌套
* 编译速度快
* 易维护
* 极适合解释器项目

适合作为 Rust 版 VBScript / ASP Runtime 的表达式解析核心。
