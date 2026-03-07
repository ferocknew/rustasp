# VBScript 常用函数及功能

本文档汇总了 VBScript 中常用的内置函数，包括数学函数、转换函数、字符串函数、日期时间函数和检验函数。

---

## 一、数学函数

| 函数名 | 语法 | 功能 |
|--------|------|------|
| Abs | `Abs(number)` | 返回一个数的绝对值 |
| Sqr | `Sqr(number)` | 返回一个数的平方根 |
| Sin | `Sin(number)` | 返回角度的正弦值 |
| Cos | `Cos(number)` | 返回角度的余弦值 |
| Tan | `Tan(number)` | 返回角度的正切值 |
| Atn | `Atn(number)` | 返回角度的反正切值 |
| Log | `Log(number)` | 返回一个数的自然对数 |
| Int | `Int(number)` | 取整函数，返回一个小于 number 的第一个整数 |
| FormatNumber | `FormatNumber(number, numdigitsafterdecimal)` | 转化为指定小数位数的数字 |
| Rnd | `Rnd()` | 返回一个从 0 到 1 的随机数 |
| Ubound | `Ubound(数组名, 维数)` | 返回该数组的最大下标 |
| Lbound | `Lbound(数组名, 维数)` | 返回最小下标数 |

### Rnd 函数详解

**语法：** `Rnd[(number)]`

返回一个随机数。参数 number 可以是任何的数值表达式。

**注解：**
- Rnd 函数返回的随机数介于 0 和 1 之间，可等于 0，但不等于 1

**number 的值会影响 Rnd 返回的随机数：**

| Number 的取值 | 返回值 |
|---------------|--------|
| 小于 0 | 每次都是使用 number 当做随机结果 |
| 大于 0 | 随机序列中的下一个随机数 |
| 等于 0 | 最近一次产生过的随机数 |
| 省略 | 随机序列中的下一个随机数 |

---

## 二、转换函数

| 函数 | 功能 |
|------|------|
| `CStr(variant)` | 将变量 variant 转化为字符串类型 |
| `CDate(variant)` | 将变量 variant 转化为日期类型 |
| `CInt(variant)` | 将变量 variant 转化为整数类型 |
| `CLng(variant)` | 将变量 variant 转化为长整数类型 |
| `CSng(variant)` | 将变量 variant 转化为 Single 类型 |
| `CDbl(variant)` | 将变量 variant 转化为 Double 类型 |
| `CBool(variant)` | 将变量 variant 转化为布尔类型 |

### 数值类型范围

| 类型 | 范围 |
|------|------|
| **整型 (Integer)** | -32,768 到 32,767 |
| **长整型 (Long)** | -2,147,483,648 到 2,147,483,647 |
| **单精度型 (Single)** | 负数：-3.402823E38 到 -1.401298E-45；正数：1.401298E-45 到 3.402823E38 |
| **双精度型 (Double)** | 负数：-1.79769313486232E308 到 -4.94065645841247E-324；正数：4.94065645841247E-324 到 1.79769313486232E308 |

---

## 三、字符串函数

| 函数 | 语法 | 功能 |
|------|------|------|
| Len | `Len(string)` | 返回 string 字符串里的字符数目 |
| Trim | `Trim(string)` | 将字符串前后的空格去掉 |
| Ltrim | `Ltrim(string)` | 将字符串前面的空格去掉 |
| Rtrim | `Rtrim(string)` | 将字符串后面的空格去掉 |
| Mid | `Mid(string, start, length)` | 从 string 字符串的 start 字符开始取得 length 长度的字符串，如果省略第三个参数表示从 start 字符开始到字符串结尾的字符串 |
| Left | `Left(string, length)` | 从 string 字符串的左边取 length 长度的字符串 |
| Right | `Right(string, length)` | 从 string 字符串的右边取得 length 长度的字符串 |
| LCase | `LCase(string)` | 将字符串里的所有大写字母转化成小写字母 |
| UCase | `UCase(string)` | 将字符串里的小写字母转化成大写字母 |
| StrComp | `StrComp(string1, string2)` | 返回 string1 字符串与 string2 字符串的比较结果，如果两个字符串相同，返回 0 |
| InStr | `InStr(string1, string2)` | 返回 string2 字符串在 string1 字符串中第一次出现的位置 |
| Split | `Split(string1, delimiter)` | 将字符串根据 delimiter 拆分成一维数组，其中 delimiter 用于表示子字符串界限的字符，如果省略，使用空格 ("") 当作分隔符 |
| Replace | `Replace(string1, find, replacewith)` | 返回字符串，其中指定的子字符串 (find) 被替换为另一个子字符串 (replacewith) |

---

## 四、日期时间函数

| 函数 | 语法 | 功能 |
|------|------|------|
| Now | `Now()` | 取得系统当前的日期和时间 |
| Date | `Date()` | 取得系统当前的日期 |
| Time | `Time()` | 取得系统当前的时间 |
| Year | `Year(Date)` | 取得给定日期的年份 |
| Month | `Month(Date)` | 取得给定日期的月份 |
| Day | `Day(Date)` | 取得给定日期是几号 |
| Hour | `Hour(time)` | 取得给定时间是第几小时 |
| Minute | `Minute(time)` | 取得给定时间是第几分钟 |
| Second | `Second(time)` | 取得给定时间是第几秒 |
| WeekDay | `WeekDay(Date)` | 取得给定日期是星期几的整数（1 表示星期一，2 表示星期二，依次类推）|
| DateDiff | `DateDiff("Var", Var1, Var2)` | 计算两个日期或时间的间隔 |
| DateAdd | `DateAdd("Var", Var1, Var2)` | 对两个日期或时间作加法 |
| FormatDateTime | `FormatDateTime(Date, format)` | 格式化日期时间 |

### 日期间隔因子

| 间隔因子 | 说明 |
|----------|------|
| yyyy | 年 |
| m | 月 |
| d | 日 |
| ww | 星期 |
| h | 小时 |
| s | 秒 |

**示例：**
```vbscript
DateAdd("d", 10, Date())  ' 10天后是几号
```

---

## 五、检验函数

| 函数 | 功能 |
|------|------|
| `VarType(variant)` | 检查变量类型（0=空，2=整数，7=日期，8=字符串，11=布尔，8192=数组）|
| `IsNumeric(variant)` | 检查是否为数值类型 |
| `IsNull(variant)` | 检查是否为 Null |
| `IsEmpty(variant)` | 检查是否为 Empty |
| `IsObject(variant)` | 检查是否为对象类型 |
| `IsDate(variant)` | 检查是否为日期类型 |
| `IsArray(variant)` | 检查是否为数组类型 |

---

## 参考说明

本文档整理了 VBScript 中的常用内置函数，涵盖了：
- **数学运算**：基本数学计算、随机数、数组边界
- **类型转换**：各种数据类型之间的转换
- **字符串处理**：字符串的截取、查找、替换、大小写转换
- **日期时间**：日期时间的获取、计算和格式化
- **变量检验**：变量类型的检查和验证

更多详细信息请参考 [VBScript 官方文档](https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/windows-scripting/d1wf56tt(v=vs.84))。
