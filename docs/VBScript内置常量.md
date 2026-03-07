# VBScript 内置常量

VBScript 共有 85 个内置常量，分为以下几类：

## MsgBox 相关常量

| 常量 | 值 | 说明 |
|------|-----|------|
| vbOKOnly | 0 | 只显示确定按钮 |
| vbOKCancel | 1 | 显示确定和取消按钮 |
| vbAbortRetryIgnore | 2 | 显示终止、重试和忽略按钮 |
| vbYesNoCancel | 3 | 显示是、否和取消按钮 |
| vbYesNo | 4 | 显示是和否按钮 |
| vbRetryCancel | 5 | 显示重试和取消按钮 |
| vbCritical | 16 | 显示临界消息图标 |
| vbQuestion | 32 | 显示警告查询图标 |
| vbExclamation | 48 | 显示警告消息图标 |
| vbInformation | 64 | 显示信息消息图标 |
| vbDefaultButton1 | 0 | 第一个按钮是默认按钮 |
| vbDefaultButton2 | 256 | 第二个按钮是默认按钮 |
| vbDefaultButton3 | 512 | 第三个按钮是默认按钮 |
| vbDefaultButton4 | 768 | 第四个按钮是默认按钮 |
| vbApplicationModal | 0 | 应用程序模式消息框 |
| vbSystemModal | 4096 | 系统模式消息框 |
| vbMsgBoxHelpButton | 16384 | 添加帮助按钮 |
| vbMsgBoxSetForeground | 65536 | 设置前台窗口 |
| vbMsgBoxRight | 524288 | 文本右对齐 |
| vbMsgBoxRtlReading | 1048576 | 从右向左阅读文本 |

## MsgBox 返回值常量

| 常量 | 值 | 说明 |
|------|-----|------|
| vbOK | 1 | 确定按钮被点击 |
| vbCancel | 2 | 取消按钮被点击 |
| vbAbort | 3 | 终止按钮被点击 |
| vbRetry | 4 | 重试按钮被点击 |
| vbIgnore | 5 | 忽略按钮被点击 |
| vbYes | 6 | 是按钮被点击 |
| vbNo | 7 | 否按钮被点击 |

## VarType 返回值常量

| 常量 | 值 | 说明 |
|------|-----|------|
| vbEmpty | 0 | 未初始化（默认） |
| vbNull | 1 | 不包含任何有效数据 |
| vbInteger | 2 | 整型子类型 |
| vbLong | 3 | 长整型子类型 |
| vbSingle | 4 | 单精度浮点子类型 |
| vbDouble | 5 | 双精度浮点子类型 |
| vbCurrency | 6 | 货币子类型 |
| vbDate | 7 | 日期子类型 |
| vbString | 8 | 字符串子类型 |
| vbObject | 9 | 对象 |
| vbError | 10 | 错误子类型 |
| vbBoolean | 11 | 布尔子类型 |
| vbVariant | 12 | Variant（仅用于变量数组） |
| vbDataObject | 13 | 数据访问对象 |
| vbDecimal | 14 | 十进制子类型 |
| vbByte | 17 | 字节子类型 |
| vbArray | 8192 | 数组 |

## 三态常量

| 常量 | 值 | 说明 |
|------|-----|------|
| vbUseDefault | -2 | 使用默认设置 |
| vbTrue | -1 | True |
| vbFalse | 0 | False |

## 比较常量

| 常量 | 值 | 说明 |
|------|-----|------|
| vbBinaryCompare | 0 | 执行二进制比较 |
| vbTextCompare | 1 | 执行文本比较 |
| vbDatabaseCompare | 2 | 执行数据库比较 |

## 日期格式常量

| 常量 | 值 | 说明 |
|------|-----|------|
| vbGeneralDate | 0 | 显示日期和/或时间 |
| vbLongDate | 1 | 使用区域设置的长日期格式 |
| vbShortDate | 2 | 使用区域设置的短日期格式 |
| vbLongTime | 3 | 使用区域设置的时间格式 |
| vbShortTime | 4 | 使用24小时格式显示时间 |

## 星期常量

| 常量 | 值 | 说明 |
|------|-----|------|
| vbUseSystemDayOfWeek | 0 | 使用系统设置 |
| vbSunday | 1 | 星期日 |
| vbMonday | 2 | 星期一 |
| vbTuesday | 3 | 星期二 |
| vbWednesday | 4 | 星期三 |
| vbThursday | 5 | 星期四 |
| vbFriday | 6 | 星期五 |
| vbSaturday | 7 | 星期六 |

## 年周常量

| 常量 | 值 | 说明 |
|------|-----|------|
| vbFirstJan1 | 1 | 从1月1日所在的周开始 |
| vbFirstFourDays | 2 | 从第一个至少有4天的周开始 |
| vbFirstFullWeek | 3 | 从第一个完整周开始 |

## 颜色常量

| 常量 | 值 | 颜色 |
|------|-----|------|
| vbBlack | 0x000000 | 黑色 |
| vbBlue | 0xFF0000 | 蓝色 |
| vbCyan | 0xFFFF00 | 青色 |
| vbGreen | 0x00FF00 | 绿色 |
| vbMagenta | 0xFF00FF | 洋红 |
| vbRed | 0x0000FF | 红色 |
| vbWhite | 0xFFFFFF | 白色 |
| vbYellow | 0x00FFFF | 黄色 |

## 字符串常量

| 常量 | 值 | 说明 |
|------|-----|------|
| vbCr | Chr(13) | 回车符 |
| vbCrLf | Chr(13) & Chr(10) | 回车换行符 |
| vbNewLine | Chr(13) & Chr(10) 或 Chr(10) | 换行符（平台相关） |
| vbFormFeed | Chr(12) | 换页符 |
| vbLf | Chr(10) | 换行符 |
| vbNullChar | Chr(0) | 空字符 |
| vbNullString | "" | 空字符串 |
| vbTab | Chr(9) | 制表符 |
| vbVerticalTab | Chr(11) | 垂直制表符 |

## 其他常量

| 常量 | 值 | 说明 |
|------|-----|------|
| vbUseSystem | 0 | 使用系统区域设置 |
| vbObjectError | -2147221504 | 自定义错误的起始编号 |

## 实现优先级

### 高优先级（核心功能）
1. **MsgBox 返回值常量** - vbOK, vbCancel, vbAbort, vbRetry, vbIgnore, vbYes, vbNo
2. **VarType 常量** - vbEmpty, vbNull, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate, vbString, vbObject, vbError, vbBoolean, vbVariant, vbByte, vbArray
3. **三态常量** - vbTrue, vbFalse, vbUseDefault
4. **比较常量** - vbBinaryCompare, vbTextCompare
5. **日期格式常量** - vbGeneralDate, vbLongDate, vbShortDate, vbLongTime, vbShortTime

### 中优先级（常用功能）
1. **颜色常量** - vbBlack, vbBlue, vbCyan, vbGreen, vbMagenta, vbRed, vbWhite, vbYellow
2. **星期常量** - vbSunday, vbMonday, vbTuesday, vbWednesday, vbThursday, vbFriday, vbSaturday

### 低优先级（扩展功能）
1. **MsgBox 按钮和图标常量**
2. **年周常量**
3. **字符串常量** - vbCr, vbCrLf, vbNewLine, vbTab 等
