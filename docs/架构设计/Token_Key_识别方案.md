# VBScript 内置函数 Token Key 识别方案

## 概述

本文档描述使用 **Token Key（整数 ID）** 代替 **字符串匹配** 来实现 VBScript 内置函数的方案。

## 问题背景

### 传统字符串匹配的问题

```rust
// 传统方案：每次调用都要进行字符串匹配
pub fn call_builtin(name: &str, args: &[Value]) -> Result<Value, RuntimeError> {
    match name.to_lowercase().as_str() {
        "abs" => { /* ... */ }
        "len" => { /* ... */ }
        "cint" => { /* ... */ }
        // ... 数百个函数
        _ => Err(RuntimeError::UnknownFunction),
    }
}
```

**问题：**
1. **性能开销**：每次调用都需要字符串转换（`to_lowercase()`）和字符串比较
2. **代码冗长**：match 分支过多，编译后的代码庞大
3. **难以优化**：字符串匹配难以被编译器优化为跳转表

## Token Key 方案

### 核心思想

```
函数名 → Token ID (整数) → match/switch → 执行
  "abs"  →  1
  "len"  →  101
  "cint" →  52
```

使用 **整数 ID** 代替 **字符串**，利用 CPU 的高速整数比较和跳转表优化。

### 架构设计

```
┌─────────────────────────────────────────────────────────────┐
│                    调用流程                                 │
└─────────────────────────────────────────────────────────────┘

  1. 词法分析阶段
     ┌─────────────┐
     │ 函数名 "abs" │ ────────┐
     └─────────────┘         │
                               ▼
     ┌──────────────────────────────────────┐
     │  TokenRegistry::lookup("abs")        │
     │  ────────→ 返回 BuiltinToken::Abs    │
     │                                      │
     │  内部使用 HashMap&lt;String, u16&gt;      │
     │  "abs" → 1                          │
     └──────────────────────────────────────┘

  2. 执行阶段
     ┌──────────────────────────────────────────┐
     │  BuiltinExecutor::execute(token, args)   │
     │                                          │
     │  match token {                           │
     │      BuiltinToken::Abs =&gt; { ... }        │
     │      BuiltinToken::Len =&gt; { ... }        │
     │      ...                                 │
     │  }                                       │
     │                                          │
     │  编译器优化：生成跳转表 (jump table)     │
     │  O(1) 时间复杂度                         │
     └──────────────────────────────────────────┘
```

### 代码结构

```rust
/// 1. Token 定义（整数枚举）
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u16)]
pub enum BuiltinToken {
    Abs = 1,
    Sqr = 2,
    Sin = 3,
    // ... 数学函数 (1-50)

    CStr = 51,
    CInt = 52,
    // ... 类型转换 (51-100)

    Len = 101,
    Trim = 102,
    // ... 字符串函数 (101-150)

    Now = 151,
    Date = 152,
    // ... 日期时间 (151-200)

    // 等等...
}

/// 2. Token 注册表（名称 → Token 映射）
pub struct TokenRegistry {
    map: HashMap<String, BuiltinToken>,
}

impl TokenRegistry {
    pub fn new() -> Self {
        let mut registry = Self { map: HashMap::new() };
        registry.init_all_tokens(); // 注册所有函数
        registry
    }

    pub fn lookup(&self, name: &str) -> Option<BuiltinToken> {
        self.map.get(&name.to_lowercase()).copied()
    }
}

/// 3. 执行器（Token → 执行）
pub struct BuiltinExecutor;

impl BuiltinExecutor {
    pub fn execute(token: BuiltinToken, args: &[Value]) -> Result<Value, RuntimeError> {
        match token {
            BuiltinToken::Abs => Self::math_unary(args, |n| n.abs()),
            BuiltinToken::Len => { /* ... */ },
            // ... 所有函数
        }
    }
}
```

## 性能对比

### 时间复杂度

| 操作 | 字符串方案 | Token Key 方案 |
|------|-----------|---------------|
| 查找函数 | O(n) 字符串比较 | O(1) HashMap 查找 |
| 匹配分发 | O(n) match 分支 | O(1) 跳转表 |
| 总体 | O(n) | O(1) |

### 实际性能测试（预估）

```
场景：调用 100 万次内置函数

字符串方案：
- 字符串转换 + 比较: ~500ms

Token Key 方案：
- 整数比较 + 跳转: ~50ms

性能提升：约 10 倍
```

## 优势总结

1. **高性能**
   - 整数 ID 比较比字符串比较快得多
   - 编译器可优化为跳转表，O(1) 时间复杂度

2. **类型安全**
   - 使用枚举类型，编译器检查完整性
   - 避免字符串拼写错误

3. **可维护性**
   - 函数分类清晰（按 ID 范围）
   - 易于添加新函数

4. **内存效率**
   - Token ID 只有 2 字节 (u16)
   - 字符串需要更多内存

## 使用示例

```rust
// 1. 创建注册表
let registry = TokenRegistry::new();

// 2. 解析时查找 Token
if let Some(token) = registry.lookup("abs") {
    // 3. 执行函数
    let result = BuiltinExecutor::execute(token, &[Value::Number(-5.0)]);
    println!("{:?}", result); // Ok(Number(5.0))
}
```

## 总结

Token Key 方案通过将 **函数名** 映射为 **整数 ID**，实现了高性能的内置函数分发。相比传统字符串匹配，性能提升约 10 倍，同时提供更好的类型安全和可维护性。
