# ASP 对象系统设计说明（Rust 实现）

本文档说明如何在 Rust 中实现 Classic ASP / VBScript 风格的对象系统，重点覆盖：

* `Server.CreateObject` 机制
* `Scripting.Dictionary` 示例实现
* 统一对象抽象层（trait）
* 运行时 Value 设计
* 对象注册机制
* 扩展建议

---

# 一、设计目标

Classic ASP 允许如下代码：

```vb
Set dict = Server.CreateObject("Scripting.Dictionary")
dict.Add "a", 1
Response.Write dict.Item("a")
```

目标是在 Rust 中实现：

* 通过字符串创建对象
* 支持方法调用
* 支持属性访问
* 支持运行时动态分发
* 可扩展更多内建对象

---

# 二、内建基础对象优先级说明（重要架构调整）

在真实 Classic ASP 运行时中，以下对象是“天然存在”的：

* Request
* Response
* Session
* Application
* Server

它们并不是通过 `Server.CreateObject` 创建的。

因此在架构设计上应当：

1️⃣ **优先实现基础内建对象**（作为 Runtime 全局上下文的一部分）
2️⃣ 再实现 `Server.CreateObject` 作为可选扩展机制

推荐实现顺序：

```
1. Value 类型系统
2. AspObject trait
3. Response（最简单，先能输出）
4. Request（表单 / QueryString）
5. Session（基于 HashMap）
6. Application（全局共享）
7. Server
8. 最后才是 Server.CreateObject + 白名单对象
```

原因：

* Response 是执行 ASP 页面必须存在的
* Request 是页面输入来源
* Session/Application 影响作用域模型
* Server 只是一个工具对象
* CreateObject 是“扩展能力”，不是运行时基础

如果顺序反过来，会导致运行时上下文结构设计被动。

---

# 三、核心架构概览

整体结构：

```
Value::Object
    ↓
Rc<RefCell<dyn AspObject>>
    ↓
具体对象实现（Dictionary / Request / Response / ...）
```

---

# 三、Value 运行时类型设计

```rust
use std::rc::Rc;
use std::cell::RefCell;

pub enum Value {
    Number(f64),
    String(String),
    Bool(bool),
    Null,
    Object(Rc<RefCell<dyn AspObject>>),
}
```

说明：

* 使用 `Rc` 支持多变量引用同一对象
* 使用 `RefCell` 支持运行时可变性
* `Object` 使用 trait 对象实现动态分发

---

# 四、统一对象抽象：AspObject

```rust
pub trait AspObject {
    fn call_method(
        &mut self,
        name: &str,
        args: Vec<Value>,
    ) -> Result<Value, RuntimeError>;

    fn get_property(&self, name: &str) -> Option<Value>;
}
```

说明：

* `call_method` 用于执行对象方法
* `get_property` 用于读取属性
* 所有内建对象必须实现该 trait

---

# 五、Dictionary 示例实现

## 1️⃣ 结构定义

```rust
use std::collections::HashMap;

pub struct Dictionary {
    data: HashMap<String, Value>,
}
```

## 2️⃣ 方法实现

```rust
impl AspObject for Dictionary {
    fn call_method(
        &mut self,
        name: &str,
        args: Vec<Value>,
    ) -> Result<Value, RuntimeError> {
        match name.to_lowercase().as_str() {
            "add" => {
                let key = args[0].to_string();
                let value = args[1].clone();
                self.data.insert(key, value);
                Ok(Value::Null)
            }
            "item" => {
                let key = args[0].to_string();
                Ok(self.data.get(&key).cloned().unwrap_or(Value::Null))
            }
            "exists" => {
                let key = args[0].to_string();
                Ok(Value::Bool(self.data.contains_key(&key)))
            }
            _ => Err(RuntimeError::MethodNotFound),
        }
    }

    fn get_property(&self, name: &str) -> Option<Value> {
        match name.to_lowercase().as_str() {
            "count" => Some(Value::Number(self.data.len() as f64)),
            _ => None,
        }
    }
}
```

---

# 六、Server.CreateObject 实现

```rust
pub fn create_object(name: &str) -> Result<Value, RuntimeError> {
    match name.to_lowercase().as_str() {
        "scripting.dictionary" => {
            let dict = Dictionary {
                data: HashMap::new(),
            };
            Ok(Value::Object(Rc::new(RefCell::new(dict))))
        }
        _ => Err(RuntimeError::ObjectNotSupported),
    }
}
```

说明：

* 使用白名单机制
* 不允许任意字符串创建对象
* 避免安全风险

---

# 七、对象注册机制（可扩展版本）

推荐将对象创建逻辑抽象为注册表：

```rust
type ObjectFactory = fn() -> Value;

pub struct ObjectRegistry {
    map: HashMap<String, ObjectFactory>,
}
```

注册示例：

```rust
registry.register("Scripting.Dictionary", || {
Value::Object(Rc::new(RefCell::new(Dictionary::new())))
});
```

优点：

* 更易扩展
* 插件化支持
* 可动态注册内建对象

---

# 八、解释器调用流程

当解析到：

```vb
dict.Add "a", 1
```

解释器执行步骤：

1. 查找变量 `dict`
2. 判断其为 `Value::Object`
3. 调用 `call_method`

示例：

```rust
if let Value::Object(obj) = dict_value {
obj.borrow_mut().call_method("Add", args)
}
```

---

# 九、VBScript 特殊行为注意事项

### 1️⃣ 大小写不敏感

建议所有方法名、属性名、键名统一转小写。

---

### 2️⃣ 默认成员（Item）

VBScript 允许：

```vb
dict("a")
```

等价于：

```vb
dict.Item("a")
```

需要在解释器层做语法映射。

---

### 3️⃣ 隐式类型转换

VBScript 会自动做类型转换，Rust 运行时需要实现：

* `to_string()`
* `to_number()`
* `to_bool()`

---

# 十、扩展对象建议

后续可以实现：

* Request
* Response
* Session
* Application
* ADODB.Connection
* ADODB.Recordset

建议每个对象单独文件，不超过 500 行。

---

# 十一、推荐目录结构

```
src/
 ├── runtime/
 │    ├── value.rs
 │    ├── object.rs
 │    ├── registry.rs
 │
 ├── objects/
 │    ├── dictionary.rs
 │    ├── request.rs
 │    ├── response.rs
 │    ├── session.rs
 │
 └── server.rs
```

---

# 十二、总结

本设计通过：

* `Value::Object`
* `dyn AspObject`
* `Rc<RefCell<>>`
* 对象注册表

实现了一个可扩展、跨平台、动态分发的 ASP 对象系统。

该架构可以支持：

* Server.CreateObject
* 方法调用
* 属性访问
* 内建对象扩展

这是一个“模拟 COM 但跨平台”的纯 Rust 运行时对象系统实现方案。
