# VBScript 单行 Sub/Function 解析错误分析

## 错误现象

```
Parser error at position 2367: Unexpected token in expression: Colon
```

出错代码：
```vbscript
Public Sub Echo(ByVal s) : Response.Write s : End Sub
```

错误位置在参数列表 `ByVal s)` 之后的冒号 `:` (位置 501:28)。

## 问题分析

### 1. 解析流程

当前解析流程如下（在 `class_parser.rs` 的 `parse_method` 方法中）：

```rust
// 第 193-195 行
let name = self.expect_ident()?;
let params = self.parse_params()?;

// 第 199 行 - 检查是否是单行方法
let is_single_line = self.check(&Token::Colon);
```

### 2. 根本原因

**问题核心**：`parse_params()` 方法在解析完参数列表后，**没有跳过换行符**，但也没有处理可能的冒号。

让我们看看 `parse_params` 的实现（`proc_stmt.rs` 第 72-109 行）：

```rust
pub fn parse_params(&mut self) -> Result<Vec<Param>, ParseError> {
    let mut params = vec![];

    if self.match_token(&Token::LParen) {
        // ... 解析参数 ...
        self.expect(Token::RParen)?;  // 第 105 行：期望右括号
    }

    Ok(params)  // 第 108 行：直接返回，没有处理任何后续 token
}
```

### 3. 问题表现

当解析器遇到以下代码时：

```vbscript
Public Sub Echo(ByVal s) : Response.Write s : End Sub
```

执行顺序：

1. `parse_method` 消耗 `Public Sub Echo`
2. `parse_params` 消耗 `(ByVal s)`，当前 token 现在指向冒号 `:`
3. `is_single_line = self.check(&Token::Colon)` - **检测到冒号，判断为单行方法**
4. 进入单行方法解析循环（第 205-239 行）
5. 第 206 行：`self.skip_newlines()` - 跳过换行（但没有换行）
6. 第 209 行：检查 `End` 关键字（当前是冒号，不匹配）
7. 第 218 行：调用 `self.parse_stmt()` 解析语句

**这里出现问题**：

- 当前 token 仍然是冒号 `:`
- `parse_stmt()` 不处理冒号（冒号不在 `parse_stmt` 的分发逻辑中）
- 最终 `parse_stmt()` 会调用 `parse_expr_stmt()` → `parse_expr()` → `parse_prefix()`
- `parse_prefix()` 遇到冒号，抛出错误：**"Unexpected token in expression: Colon"**

### 4. 为什么会这样？

查看代码逻辑发现，单行方法的解析流程中存在**逻辑缺陷**：

**第 203-239 行的单行方法解析循环**：

```rust
if is_single_line {
    // 单行 Sub/Function，解析冒号分隔的语句直到 End Function/End Sub
    loop {
        self.skip_newlines();  // 第 206 行

        // 检查是否到达 End Function/End Sub
        if self.check_keyword(Keyword::End) { ... }

        // 解析语句
        if let Some(stmt) = self.parse_stmt()? {  // 第 218 行
            body.push(stmt);
        }

        // 检查是否有冒号分隔符
        self.skip_newlines();
        if !self.match_token(&Token::Colon) {  // 第 224 行
            // 错误处理...
        }
    }
}
```

**问题在于**：

1. 循环开始时，当前 token 是冒号（因为 `parse_params` 后就指向冒号了）
2. 循环第一次执行 `parse_stmt()` 时，**冒号还没有被消耗**
3. `parse_stmt()` 不识别冒号作为语句开始，尝试解析表达式
4. 表达式解析器（Pratt parser）无法处理冒号，抛出错误

### 5. 对比其他解析器的处理

查看 `proc_stmt.rs` 中的普通 Sub/Function 解析（第 42-69 行）：

```rust
pub fn parse_sub(&mut self) -> Result<Option<Stmt>, ParseError> {
    self.expect_keyword(Keyword::Sub)?;
    let name = self.expect_ident()?;

    let params = self.parse_params()?;
    self.skip_newlines();  // ← 注意：这里跳过了换行

    let mut body = vec![];
    loop {
        if self.is_at_end() || self.check_keyword(Keyword::End) {
            break;
        }
        // 处理冒号语句分隔符
        if self.match_token(&Token::Colon) {  // ← 在循环开始时处理冒号
            continue;
        }
        match self.parse_stmt()? {
            Some(stmt) => body.push(stmt),
            None => break,
        }
        self.skip_newlines();
    }
    // ...
}
```

**关键区别**：

- 普通方法在循环开始时会**先检查并消耗冒号**（第 55 行）
- 单行方法的逻辑**期望先解析语句，再消耗冒号**（第 224 行）
- 但单行方法**忘记在第一次循环前先消耗冒号**

## 解决方案

需要在单行方法解析循环开始前，**先消耗初始的冒号**：

```rust
if is_single_line {
    // ← 需要在这里消耗冒号
    self.expect(Token::Colon)?;  // 消费检查时的冒号

    // 单行 Sub/Function，解析冒号分隔的语句直到 End Function/End Sub
    loop {
        self.skip_newlines();
        // ...
    }
}
```

或者在循环逻辑中调整，让第一次循环也能正确处理冒号：

```rust
if is_single_line {
    let mut body = Vec::new();

    loop {
        self.skip_newlines();

        // 检查是否到达 End Function/End Sub
        if self.check_keyword(Keyword::End) {
            if (is_function && self.peek_next_is_keyword(Keyword::Function))
                || (is_sub && self.peek_next_is_keyword(Keyword::Sub))
            {
                break;
            }
        }

        // ← 添加：处理冒号分隔符（在解析语句之前）
        if self.match_token(&Token::Colon) {
            continue;
        }

        // 解析语句
        if let Some(stmt) = self.parse_stmt()? {
            body.push(stmt);
        }

        // 检查是否有冒号分隔符
        self.skip_newlines();
        if !self.match_token(&Token::Colon) {
            // ... 错误处理
        }
    }
}
```

## 总结

这是一个**边界条件错误**：

1. `is_single_line` 检查发现了冒号，但**没有消耗它**
2. 循环逻辑期望冒号在语句之间，但**第一次循环前没有处理初始冒号**
3. `parse_stmt()` 被调用时，当前 token 仍然是冒号
4. 冒号不在 `parse_stmt()` 的处理范围内，导致错误

修复方法：在单行方法解析循环开始前，消耗初始冒号；或在循环中增加冒号的前置检查。
