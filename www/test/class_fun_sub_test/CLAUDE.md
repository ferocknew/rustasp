# VBScript 过程测试用例

## 概述

本测试套件用于验证 VBScript 解释器对 Sub 过程和 Function 过程的支持情况。

## 测试文件列表

| 文件名 | 描述 |
|--------|------|
| test_sub.asp | Sub 过程测试 |
| test_function.asp | Function 过程测试 |
| test_params.asp | 参数传递测试 |
| test_call.asp | Call 语句测试 |

---

## 1. Sub 过程测试 (test_sub.asp)

| 用例编号 | 测试项 | 测试代码 | 预期结果 |
|----------|--------|----------|----------|
| SUB-001 | 无参数 Sub 定义和调用 | `Sub SayHello()...End Sub` | 成功定义和调用 |
| SUB-002 | 有参数 Sub 定义和调用 | `Sub Greet(name)...End Sub` | 输出 "Hello, VBScript!" |
| SUB-003 | 多参数 Sub | `Sub AddAndPrint(a, b)...End Sub` | 输出 "10 + 20 = 30" |
| SUB-004 | 空 Sub 语句 | `Sub EmptySub() End Sub` | 执行完成无错误 |
| SUB-005 | 使用 Call 调用 | `Call SayHello()` | 正常调用 |
| SUB-006 | 不使用 Call 调用无参数 | `SayHello` | 正常调用 |
| SUB-007 | 不使用 Call 调用有参数 | `Greet "Developer"` | 输出 "Hello, Developer!" |

---

## 2. Function 过程测试 (test_function.asp)

| 用例编号 | 测试项 | 测试代码 | 预期结果 |
|----------|--------|----------|----------|
| FN-001 | 无参数 Function | `Function GetVersion()...End Function` | 返回 "1.0.0" |
| FN-002 | 有参数 Function | `Function Square(n)...End Function` | Square(5) = 25 |
| FN-003 | 多参数 Function | `Function Add(a, b)...End Function` | Add(3, 4) = 7 |
| FN-004 | 数学计算 Function | `Function Celsius(fDegrees)...End Function` | Celsius(100) ≈ 37.78 |
| FN-005 | 字符串处理 Function | `Function Concat(str1, str2)...End Function` | 返回 "HelloWorld" |
| FN-006 | 布尔返回 Function | `Function IsEven(n)...End Function` | IsEven(4) = True, IsEven(5) = False |
| FN-007 | 在表达式中使用 | `result = Square(Add(2, 3))` | result = 25 |
| FN-008 | 嵌套调用 | `Square(Add(2, 3))` | 返回 25 |

---

## 3. 参数传递测试 (test_params.asp)

| 用例编号 | 测试项 | 测试代码 | 预期结果 |
|----------|--------|----------|----------|
| PARAM-001 | ByVal 按值传递 | `Sub TestByVal(ByVal x)` | 原变量值不变 |
| PARAM-002 | ByRef 按引用传递 | `Sub TestByRef(ByRef y)` | 原变量值改变 |
| PARAM-003 | 默认传递方式 | `Sub TestDefault(z)` | 默认 ByRef，值改变 |
| PARAM-004 | 混合参数传递 | `Sub TestMixed(ByVal a, ByRef b, c)` | a 不变，b、c 改变 |
| PARAM-005 | 数组参数 | `Sub TestArray(arr)` | 正确传递和遍历数组 |

### ByVal 测试详细

```
调用前 val1 = 5
Sub 内部 x = 15
调用后 val1 = 5 (应该不变)
```

### ByRef 测试详细

```
调用前 val2 = 5
Sub 内部 y = 15
调用后 val2 = 15 (应该改变)
```

---

## 4. Call 语句测试 (test_call.asp)

| 用例编号 | 测试项 | 测试代码 | 预期结果 |
|----------|--------|----------|----------|
| CALL-001 | Call 调用无参数 Sub | `Call MySub1()` | 正常调用 |
| CALL-002 | 不使用 Call 无参数 Sub | `MySub1` | 正常调用 |
| CALL-003 | Call 调用有参数 Sub | `Call MySub2(10, 20)` | 正常调用 |
| CALL-004 | 不使用 Call 有参数 Sub | `MySub2 10, 20` | 正常调用，无括号 |
| CALL-005 | Call 调用 Function | `Call MyFunc2(5, 10)` | 执行但丢弃返回值 |
| CALL-006 | 获取 Function 返回值 | `result = MyFunc2(5, 10)` | result = 15 |
| CALL-007 | 在表达式中使用 Function | `"值: " & MyFunc2(100, 200)` | 输出 "值: 300" |
| CALL-008 | Function 嵌套调用 | `MyFunc2(MyFunc2(1, 2), 3)` | 返回 6 (先计算 1+2=3，再计算 3+3=6) |

---

## 测试执行

### 执行顺序

1. test_sub.asp - 验证 Sub 过程基本功能
2. test_function.asp - 验证 Function 过程基本功能
3. test_params.asp - 验证参数传递机制
4. test_call.asp - 验证 Call 语句用法

### 验证要点

- [ ] Sub 过程定义语法 `Sub ... End Sub`
- [ ] Function 过程定义语法 `Function ... End Function`
- [ ] 函数返回值赋值给函数名
- [ ] Call 语句调用 Sub
- [ ] 不使用 Call 调用 Sub（无括号）
- [ ] ByVal 按值传递参数
- [ ] ByRef 按引用传递参数
- [ ] 默认参数传递方式（ByRef）
- [ ] 数组参数传递
- [ ] Function 在表达式中的使用
- [ ] Function 嵌套调用

---

## 已知限制

| 限制项 | 说明 | 状态 |
|--------|------|------|
| Optional 参数 | VBScript 不支持 Optional 关键字 | N/A |
| ParamArray | VBScript 不支持 ParamArray | N/A |
| 默认值 | VBScript 不支持参数默认值 | N/A |

## 参考文档

- [VBScript 过程](https://www.weistock.com/docs/VBA/VBScript/过程.html)
