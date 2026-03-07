# VBScript 解释器架构设计文档

本文档记录 VBScript 解释器的关键架构设计决策。

## 文档列表

### 1. [Token Key 识别方案](./Token_Key_识别方案.md)
描述使用 **整数 Token ID** 代替 **字符串匹配** 实现 VBScript 内置函数的方案。

**核心思想：**
- 将函数名（如 "abs"）映射为整数 ID（如 1）
- 使用 HashMap 进行 O(1) 查找
- 使用 match/switch 进行 O(1) 分发
- 编译器可优化为跳转表

**性能提升：** 约 10 倍

### 2. [新旧方案对比](./新旧方案对比.md)
详细的对比分析，包括：
- 代码示例对比
- 性能基准测试
- 内存占用对比
- 类型安全对比
- 迁移指南

## 快速开始

### 使用 Token 系统

```rust
use vbscript::runtime::interpreter::{TokenRegistry, BuiltinExecutor, BuiltinToken};

// 1. 创建 Token 注册表
let registry = TokenRegistry::new();

// 2. 查找函数对应的 Token
if let Some(token) = registry.lookup("abs") {
    // 3. 执行函数
    let result = BuiltinExecutor::execute(token, &[Value::Number(-5.0)]);
    println!("{:?}", result); // Ok(Number(5.0))
}
```

### 添加新函数

1. 在 `BuiltinToken` 枚举中添加新 Token ID
2. 在 `TokenRegistry::init_all_tokens()` 中注册函数名映射
3. 在 `BuiltinExecutor::execute()` 中实现函数逻辑

## 架构图

```
┌─────────────────────────────────────────────────────────────┐
│                    VBScript 解释器                          │
└─────────────────────────────────────────────────────────────┘

┌─────────────────┐    ┌─────────────────┐    ┌─────────────────┐
│   词法分析器     │───▶│   语法分析器     │───▶│   解释执行器    │
└─────────────────┘    └─────────────────┘    └────────┬────────┘
                                                     │
                                                     ▼
┌──────────────────────────────────────────────────────────────┐
│                    Token Key 系统                             │
├──────────────────────────────────────────────────────────────┤
│                                                              │
│  ┌──────────────┐      ┌──────────────┐      ┌────────────┐ │
│  │  函数名       │─────▶│  Token ID     │─────▶│   执行     │ │
│  │  "abs"       │      │   (u16)       │      │  逻辑      │ │
│  └──────────────┘      └──────────────┘      └────────────┘ │
│                                                              │
│  性能：O(1) HashMap 查找  +  O(1) 跳转表分发                   │
│                                                              │
└──────────────────────────────────────────────────────────────┘
```

## 相关文件

- `src/runtime/interpreter/builtin_tokens.rs` - Token 系统实现
- `src/runtime/interpreter/builtins.rs` - 旧版内置函数（逐步迁移）
- `src/runtime/interpreter/mod.rs` - 模块导出

## 参考资料

- [Rust Enum 优化](https://doc.rust-lang.org/book/ch06-00-enums.html)
- [HashMap 性能](https://doc.rust-lang.org/std/collections/struct.HashMap.html)
- [编译器跳转表优化](https://en.wikipedia.org/wiki/Branch_table)
