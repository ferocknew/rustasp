# Server.CreateObject 白名单说明（简明版）

Classic ASP 的 `Server.CreateObject` 可以用于创建内建或 COM 对象。为了安全和跨平台实现，需要进行白名单管理。

## 一、白名单对象及实现可行性

| 对象名                        | 可行性   | 实现模块                    | 说明                                     |
| -------------------------- | ----- | ----------------------- | -------------------------------------- |
| Scripting.Dictionary       | ✅ 高   | `objects/dictionary.rs` | 哈希表，存 key-value，前文已示例实现                |
| Scripting.FileSystemObject | ✅ 高   | `objects/filesystem.rs` | 文件操作，需沙箱限制路径在 WebRoot 内                |
| MSXML2.DOMDocument         | ✅ 高   | `objects/xml.rs`        | XML 解析，可用 `roxmltree` 或 `quick-xml` 实现 |
| ADODB.Connection           | ⚠ 可选  | `objects/db.rs`         | 数据库连接抽象（SQLite/SQL Server），必须封装安全接口    |
| ADODB.Recordset            | ⚠ 可选  | `objects/db.rs`         | 数据查询结果封装，与 Connection 配套               |
| Scripting.Runtime          | ⚠ 可选  | `objects/runtime.rs`    | Timer、Date 等基础函数，可部分实现                 |
| WScript.Shell              | ❌ 不安全 | N/A                     | 直接执行系统命令，不跨平台，不实现                      |
| 其他 Windows COM             | ❌ 不安全 | N/A                     | 不支持，避免 RCE 风险                          |

> ✅ 高：可安全跨平台实现
> ⚠ 可选：可实现，但需额外抽象
> ❌ 不安全：不实现

## 二、白名单实现思路（Rust）

* 使用 `ObjectRegistry` 维护对象注册表
* 工厂函数返回 `Value::Object(Rc<RefCell<dyn AspObject>>)`
* 未注册对象禁止创建

示例代码：

```rust
let mut registry = ObjectRegistry::new();

registry.register("Scripting.Dictionary", || {
    Value::Object(Rc::new(RefCell::new(Dictionary::new())))
});

registry.register("Scripting.FileSystemObject", || {
    Value::Object(Rc::new(RefCell::new(FileSystemObject::new(web_root))))
});

registry.register("MSXML2.DOMDocument", || {
    Value::Object(Rc::new(RefCell::new(XmlDocument::new())))
});
```

## 三、设计建议

1. 基础内建对象（Request / Response / Session / Application）必须优先实现，不通过 CreateObject 创建。
2. 白名单对象仅限扩展对象，如 Dictionary / 文件 / XML / 数据库抽象。
3. 危险对象（系统命令、Windows COM）绝对禁止。
4. 对象内部应保证跨平台和沙箱安全。
