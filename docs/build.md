# 编译与测试指南

## 环境要求

- Rust 1.70+ (推荐使用 rustup 安装)
- Cargo (随 Rust 自动安装)

## 编译命令

### 检查代码（快速）

```bash
cargo check
```

只进行类型检查，不生成二进制文件，速度最快。

### 开发编译

```bash
cargo build
```

生成 debug 版本，包含调试信息，编译速度快但运行较慢。

### Release 编译

```bash
cargo build --release
```

生成优化后的版本，编译较慢但运行性能最佳。

### 详细输出

```bash
cargo build --verbose
```

显示详细的编译过程，用于调试编译问题。

## 运行命令

### 开发运行

```bash
cargo run
```

### Release 运行

```bash
cargo run --release
```

### 带日志运行

```bash
RUST_LOG=debug cargo run
```

## 测试命令

### 运行所有测试

```bash
cargo test
```

### 运行特定测试

```bash
cargo test test_name
```

### 显示测试输出

```bash
cargo test -- --nocapture
```

## 代码检查

### Clippy 静态分析

```bash
cargo clippy
```

### 格式化代码

```bash
cargo fmt
```

### 格式化检查

```bash
cargo fmt -- --check
```

## 常见编译问题

### 依赖下载失败

```bash
# 更新依赖索引
cargo update

# 清理缓存重新编译
cargo clean
cargo build
```

### 类型错误

检查模块导入是否正确：
```rust
use crate::module::Type;
```

### 链接错误

确保所有依赖版本兼容，检查 Cargo.toml。

## 项目特定说明

### 模块结构

本项目采用分层架构：
- `ast/` - 语法树定义（无外部依赖）
- `parser/` - 解析器（依赖 chumsky）
- `runtime/` - 运行时（独立模块）
- `builtins/` - ASP 内建对象
- `asp/` - ASP 引擎
- `http/` - HTTP 服务（依赖 axum）

### 编译顺序

Cargo 自动处理依赖关系，无需手动指定编译顺序。

### 特性开关

未来可能添加的特性：
```toml
[features]
default = ["sqlite"]
sqlite = ["sqlx"]
mssql = ["tiberius"]
```
