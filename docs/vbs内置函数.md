##  日期/时间函数

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#cdate-函数)CDate 函数

  把有效的日期和时间表达式转换为日期（Date）类型，并返回结果。

CDate(date)



参数

| 参数 | 描述                   |
| ---- | ---------------------- |
| date | 任意有效的日期表达式。 |

说明

\* IsDate 函数用于判断 date 是否可以被转换为日期或时间。CDate 识别日期文字和时间文字，以及一些在可接受的日期范围内的数字。在将数字转换为日期时，数字的整数部分被转换为日期，分数部分被转换为从午夜开始计算的时间。 * CDate 根据系统的区域设置识别日期格式。如果数据的格式不能被日期设置识别，则不能判断年、月、日的正确顺序。另外，如果长日期格式包含表示星期几的字符串，则不能被识别。

示例

  下面的示例使用 CDate 函数将字符串转换成日期类型。一般不推荐使用硬件译码日期和时间作为字符串（下面的例子已体现）。而使用时间和日期文字 (如 #10/19/1962#, #4:45:23 PM#)。

```vb
MyDate = "October 19, 1962"       ' 定义日期。
MyShortDate = CDate(MyDate)        ' 转换为日期数据类型。
MyTime = "4:35:47 PM"        ' 定义时间。
MyShortTime = CDate(MyTime)        ' 转换为日期数据类型。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#date-函数)Date 函数

  返回当前系统日期。

Date



示例

```vb
Dim MyDate
MyDate = Date    ' MyDate 包含当前系统日期。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#dateadd-函数)DateAdd 函数

  返回已添加指定时间间隔后的新日期。

DateAdd(interval, number, date)



参数

| 参数     | 描述                                                         |
| -------- | ------------------------------------------------------------ |
| interval | 必需。您想要添加的时间间隔。 可采用下面的值：   yyyy - 年   q - 季度   m - 月   y - 当年的第几天   d - 日   w - 当周的第几天   ww - 当年的第几周   h - 小时   n - 分   s - 秒 |
| number   | 必需。需要添加的时间间隔的数目。可对未来的日期使用正值，对过去的日期使用负值。 |
| date     | 必需。代表被添加的时间间隔的日期的变量或文字。               |

说明

-   可用 DateAdd 函数从日期中添加或减去指定时间间隔。例如可以使用 DateAdd 从当天算起 30 天以后的日期或从现在算起 45 分钟以后的时间。要向 date 添加以“日”为单位的时间间隔，可以使用“一年的日数”（“y”）、“日”（“d”）或“一周的日数”（“w”）。
-   DateAdd 函数不会返回无效日期。如下示例将 95 年 1 月 31 日加上一个月：

```vb
NewDate = DateAdd("m", 1, "31-Jan-95")
```

  在这个示例中，DateAdd 返回 95 年 2 月 28 日，而不是 95 年 2 月 31 日。如果 date 为 96 年 1 月 31 日，则返回 96 年 2 月 29 日，这是因为 1996 是闰年。

-   如果计算的日期是在公元 100 年之前，则会产生错误。
-   如果 number 不是 Long 型值，则在计算前四舍五入为最接近的整数。

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#datediff-函数)DateDiff 函数

  用于判断在两个日期之间存日期间隔

DateDiff(interval, date1, date2 [,firstdayofweek [, firstweekofyear]])



参数

DateDiff 函数的语法有以下参数：

| 参数            | 描述                                                         |
| --------------- | ------------------------------------------------------------ |
| interval        | 必需。计算 date1 和 date2 之间的时间间隔的单位。可采用下面的值：   yyyy - 年   q - 季度   m - 月   y - 当年的第几天   d - 日   w - 统计date1对应的星期数在date1和date2区间出现的次数   ww - 统计区间每周的第一天出现的个数   h - 小时   n - 分   s - 秒 |
| date1,date2     | 必需。日期表达式。在计算中需要使用的两个日期。               |
| firstdayofweek  | 可选。规定一周的日数，即当周的第几天。 可采用下面的值：   0 = vbUseSystemDayOfWeek - 使用区域语言支持（NLS）API 设置   1 = vbSunday - 星期日（默认）   2 = vbMonday - 星期一   3 = vbTuesday - 星期二   4 = vbWednesday - 星期三   5 = vbThursday - 星期四   6 = vbFriday - 星期五   7 = vbSaturday - 星期六 |
| firstweekofyear | 可选。规定一年中的第一周。 可采用下面的值：   0 = vbUseSystem - 使用区域语言支持（NLS）API 设置   1 = vbFirstJan1 - 由 1 月 1 日所在的星期开始（默认）   2 = vbFirstFourDays - 由在新的一年中至少有四天的第一周开始   3 = vbFirstFullWeek - 由在新的一年中第一个完整的周开始 |

说明

-   如果 date1 晚于 date2，则 DateDiff 函数返回负数。
-   firstdayofweek 参数会对使用“w”和“ww”间隔符号的计算产生影响。

示例

  下面的示例利用 DateDiff 函数显示今天与给定日期之间间隔天数:

```vb
Function DiffADate(theDate)
    '要计算 date1 和 date2 相差的天数，可以使用“一年的日数”（“y”）或“日”（“d”）
    DiffADate = "从当天开始的天数:" & DateDiff("y", Now, theDate)
    DiffADate = "从当天开始的天数:" & DateDiff("d", Now, theDate)

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '统计date1对应的星期数在date1和date2之间出现的次数（“w”）
    '如果 date1 是星期一，则 DateDiff 计算到 date2 之前星期一的数目。此结果包含 date2 而不包含 date1。
    DiffADate = "当天对应的星期出现的次数:" & DateDiff("w", Now, theDate)

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '统计区间内，每周第一天一共出现的次数，默认：周日为每周第一天。
    '如果 date2 是星期日，DateDiff 将计算 date2，但即使 date1 是星期日，也不会计算 date1
    DiffADate = "每周第一天出现的次数:" & DateDiff("ww", Now, theDate)

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '如果 date1 或 date2 是日期文字，则指定的年度会成为日期的固定部分。
    DiffADate = "指定年度日期下的天数间隔:" & DateDiff("d", "2019-10-01", "2019-10-10")

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '如果日期被包括在引号 (" ") 中并且省略年份，'则在代码中每次计算 date1 或 date2 表达式时，将插入当前年份。这样就可以编写适用'于不同年份的程序代码。
    DiffADate = "当年中两个日期间隔的天数:" & DateDiff("d", "10-01", "10-20")

    '在间隔单位为“年”（“yyyy”）时，比较12月31日和来年的1月1日，虽然实际上只相差一天，DateDiff 返回 1 表示相差一个年份。
    DiffADate = "间隔年数:" & DateDiff("yyyy", "2020-12-31", "2021-01-01")
End Function
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#datepart-函数)DatePart 函数

  返回给定日期的指定部分，用于计算日期并返回指定的时间间隔。

DatePart(interval, date[, firstdayofweek[, firstweekofyear]])



参数

| 参数            | 描述                                                         |
| --------------- | ------------------------------------------------------------ |
| interval        | 必需。计算 date1 和 date2 之间的时间间隔的单位。可采用下面的值：   yyyy - 年   q - 季度   m - 月   y - 当年的第几天   d - 日   w - 统计date1对应的星期数在date1和date2区间出现的次数   ww - 统计区间每周的第一天出现的个数   h - 小时   n - 分   s - 秒 |
| date            | 必需。需计算的日期表达式。                                   |
| firstdayofweek  | 可选。规定一周的日数，即当周的第几天。可采用下面的值：   0 = vbUseSystemDayOfWeek - 使用区域语言支持（NLS）API 设置   1 = vbSunday - 星期日（默认）   2 = vbMonday - 星期一   3 = vbTuesday - 星期二   4 = vbWednesday - 星期三   5 = vbThursday - 星期四   6 = vbFriday - 星期五   7 = vbSaturday - 星期六 |
| firstweekofyear | 可选。规定一年中的第一周。  可采用下面的值：   0 = vbUseSystem - 使用区域语言支持（NLS）API 设置   1 = vbFirstJan1 - 由 1 月 1 日所在的星期开始（默认）   2 = vbFirstFourDays - 由在新的一年中至少有四天的第一周开始   3 = vbFirstFullWeek - 由在新的一年中第一个完整的周（不跨年度）开始 |

说明

-   firstdayofweek 参数会影响使用“w”和“ww”间隔符号的计算。
-   如果 date 是日期文字，则指定的年度会成为日期的固定部分。但是如果 date 被包含在引号 (" ") 中，并且省略年份，则在代码中每次计算 date 表达式时，将插入当前年份。这样就可以编写适用于不同年份的程序代码。.

示例

```vb
    dim theDate,DiffADate
    theDate="2021-02-26 23:10:15"
    '显示该日所在的年份
    DiffADate = "yyyy:" & DatePart("yyyy", "theDate")

    '显示该日所在的季节
    DiffADate = "q:" & DatePart("q",  theDate)

    '显示该日所在的月
    DiffADate = "m:" & DatePart("m",  theDate)

    '显示该日是当年第几天
    DiffADate = "y:" & DatePart("y",  theDate)

    '显示该日是该月第几天
    DiffADate = "d:" & DatePart("d",  theDate)

    '显示该日是本周第几天
    DiffADate = "w:" & DatePart("w",  theDate)

    '显示该日是本月第几周
    DiffADate = "ww:" & DatePart("ww",  theDate)

    '显示该日的时，分，秒
    DiffADate = "hh:" & DatePart("h",  theDate)
    DiffADate = "nn:" & DatePart("n",  theDate)
    DiffADate = "ss:" & DatePart("s",  theDate)
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#dateserial-函数)DateSerial 函数

  DateSerial 函数返回指定的年、月、日的子类型 Date 的 Variant 。

DateSerial(year, month, day)



参数

| 参数  | 描述                                                         |
| ----- | ------------------------------------------------------------ |
| year  | 必需。介于 100 到 9999 的数字，或数值表达式。介于 0 到 99 的值被视为 1900–1999。对于所有其他的 year 参数，请使用完整的 4 位年份。 |
| month | 必需的。任何数值表达式。                                     |
| day   | 必需的。任何数值表达式。                                     |

说明

-   当任何一个参数的取值超出可接受的范围时，则会适当地进位到下一个较大的时间单位。例如，如果指定了 35 天，则这个天数被解释成一个月加上多出来的日数，多出来的日数取决于其年份和月份。但是如果参数值超出 -32,768 到 32,767 的范围，或者由三个参数指定（无论是直接还是通过表达式指定）的日期超出了可以接受的日期范围，就会发生错误。

示例

```vb
Dim MyDate1, MyDate2
MyDate1 = DateSerial(2020, 09, 1)         ' 2020/09/01
MyDate2 = DateSerial(2020 - 10, 8 - 2, 1 - 1)   ' 2010/05/31
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#datevalue-函数)DateValue 函数

  把字符串转换为 Date类型。

DateValue(date)



参数

| 参数 | 描述                                                         |
| ---- | ------------------------------------------------------------ |
| date | 必需。一个介于 100 年 1 月 1 日到 9999 年 12 月 31 日的日期，或者任何可表示日期、时间或日期时间兼有的表达式。 |

说明

\* 如果 date 参数包含时间信息，则 DateValue 不会返回时间信息。但是如果 date 包含无效的时间信息（如 "89:98"），就会出现错误。 * 如果 date 是某一字符串，其中仅包含由有效的日期分隔符分隔开的数字，则 DateValue 将会根据为系统指定的短日期格式识别月、日和年的顺序。DateValue 还会识别包含月份名称（无论是全名还是缩写）的明确日期。例如，除了能够识别 12/30/1991 和 12/30/91 之外，DateValue 还能识别 December 30, 1991 和 Dec 30, 1991。 * 如果省略了 date 的年份部分，DateValue 将使用计算机系统日期中的当前年份。

示例

  下面的示例利用 DateValue 函数将字符串转化成日期。也可以利用日期文字直接将日期分配给 Variant 变量， 例如， MyDate = #9/11/63#.

```vb
Dim MyDate
MyDate = DateValue("September 11, 1963")    ' 返回日期。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#day-函数)Day 函数

  返回 1 到 31 之间的一个整数（包括 1 和31），代表某月中的一天。

Day(date)



参数

| 参数 | 描述                                                         |
| ---- | ------------------------------------------------------------ |
| date | 任意可以代表日期的表达式。如果 date 参数中包含 Null，则返回 Null。 |

示例

  下面例子利用 Day 函数得到一个给定日期月的天数:

```vb
Dim MyDay
MyDay = Day("October 19, 1962")  'MyDay 包含 19。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#formatdatetime-函数)FormatDateTime 函数

  格式化并返回一个有效的日期或时间的表达式

FormatDateTime(Date[, NamedFormat])



参数

| 参数   | 描述                                                         |
| ------ | ------------------------------------------------------------ |
| date   | 必需。任何有效的日期表达式（比如 Date() 或者 Now()）。       |
| format | 可选。规定所使用的日期/时间格式的格式值。 可采用下面的值： 0 = vbGeneralDate - 默认。返回日期：mm/dd/yy 及如果指定时间：hh:mm:ss PM/AM。 1 = vbLongDate - 返回日期：weekday, monthname, year 2 = vbShortDate - 返回日期：mm/dd/yy 3 = vbLongTime - 返回时间：hh:mm:ss PM/AM 4 = vbShortTime - 返回时间：hh:mm |

示例

  下面例子利用 FormatDateTime 函数把表达式格式化为长日期型并且把它赋给 MyDateTime:

```vb
 Function GetCurrentDate 
  'FormatDateTime 把日期型格式化为长日期型。
  GetCurrentDate = FormatDateTime(Date, 1) 
End Function
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#hour-函数)Hour 函数

  返回 0 到 23 之间的一个整数（包括 0 和 23），代表一天中的某一小时。

Hour(time)



参数

| 参数 | 描述                                                         |
| ---- | ------------------------------------------------------------ |
| time | 必需。任意可以代表时间的表达式。如果 time 参数中包含 Null，则返回 Null。 |

示例

  下面的示例利用 Hour 函数得到当前时间的小时：

```vb
Dim MyTime, MyHour
MyTime = Now
MyHour = Hour(MyTime)   ' MyHour 包含代表当前时间的数值。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#isdate-函数)IsDate 函数

  返回 Boolean 值指明某表达式是否可以转换为日期。

IsDate(expression)



参数

| 参数       | 描述                                                         |
| ---------- | ------------------------------------------------------------ |
| expression | 参数可以是任意可被识别为日期和时间的日期表达式或字符串表达式。 |

说明

  如果表达式是日期或可合法地转化为有效日期，则 IsDate 函数返回 True；否则函数返回 False。在 Microsoft Windows 操作系统中，有效的日期范围公元 100 年 1 月 1 日到公元 9999 年 12 月 31 日；合法的日期范围随操作系统不同而不同。 下面的示例利用 IsDate 函数决定表达式是否能转换为日期型：

示例

```vb
Dim MyDate, YourDate, NoDate, MyCheck
MyDate = "October 19, 1962": YourDate = #10/19/62#: NoDate = "Hello"
MyCheck = IsDate(MyDate)            ' 返回 True。
MyCheck = IsDate(YourDate)          ' 返回 True。
MyCheck = IsDate(NoDate)            ' 返回 False。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#minute-函数)Minute 函数

  返回 0 到 59 之间的一个整数（包括 0 和59），代表一小时内的某一分钟。

Minute(time)



参数

| 参数 | 描述                                                       |
| ---- | ---------------------------------------------------------- |
| time | 必需。时间的表达式。如果 time 参数包含 Null，则返回 Null。 |

示例

  下面的示例利用 Minute 函数返回小时的分钟数：

```vb
Dim MyVar
MyVar = Minute(Now) 
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#month-函数)Month 函数

  返回表示年的月份的数字，介于 1 到 12 之间。

Month(date)



参数

| 参数 | 描述                                                     |
| ---- | -------------------------------------------------------- |
| date | 必需。日期表达式。如果 time 参数包含 Null，则返回 Null。 |

示例

下面的示例利用 Month 函数返回当前月：

```vb
Dim MyVar
MyVar = Month(Now)    ' MyVar 包含当前月对应的数字。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#monthname-函数)MonthName 函数

  返回指定的月份的名称

MonthName(month[, abbreviate])



参数

| 参数       | 描述                                                       |
| ---------- | ---------------------------------------------------------- |
| month      | 必需。规定月的数字。（比如一月是 1，二月是 2，依此类推。） |
| abbreviate | 可选。一个布尔值，指示是否缩写月份名称。默认是 False。     |

示例

下面的示例利用MonthName 函数为日期表达式返回月份的缩写：

```vb
Dim MyVar
MyVar = MonthName(10, True) ' MyVar 包含 "Oct"。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#now-函数)Now 函数

  根据计算机系统设定的日期和时间返回当前的日期和时间值。

Now



示例

下面的示例利用 Now 函数返回当前的日期和时间：

```vb
Dim MyVar
MyVar = Now ' MyVar 包含当前的日期和时间。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#second-函数)Second 函数

  返回 0 到 59 之间的一个整数（包括 1 和 59），代表一分钟内的某一秒。

Second(time)



参数

| 参数 | 描述                                                         |
| ---- | ------------------------------------------------------------ |
| time | 必需。任意可以代表时间的表达式。如果 time 参数中包含 Null，则返回 Null。 |

示例

下面的示例利用 Second 函数返回当前秒：

```vb
Dim MySec
MySec = Second(Now)
                    'MySec 包含代表当前秒的数字。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#time-函数)Time 函数

  返回当前的系统时间

Time



示例

下面的示例利用 Time 函数返回当前系统时间：

```vb
Dim MyTime
MyTime = Time    ' 返回当前系统时间。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#timer-函数)Timer 函数

  返回当日0点以来的秒数

Timer



示例

  下面的例子使用 Timer 函数来确定 For...Next 循环 N 次所需的时间：

```vb
Function TimeIt(N)
  Dim StartTime, EndTime
  StartTime = Timer
  For I = 1 To N
  Next
  EndTime = Timer
  TimeIt = EndTime - StartTime
End Function 
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#timeserial-函数)TimeSerial 函数

  返回特定小时、分钟和秒的时间。

TimeSerial(hour,minute, second)



参数

| 参数   | 描述                                   |
| ------ | -------------------------------------- |
| hour   | 必需。介于 0-23 的数字，或数值表达式。 |
| minute | 必需。介于 0-59 的数字，或数值表达式。 |
| second | 必需。介于 0-59 的数字，或数值表达式。 |

说明

-   要指定一时刻，如 11:59:59，TimeSerial 的参数取值应在可接受的范围内；也就是说，小时应介于 0-23 之间，分和秒应介于 0-59 之间。但是，可以使用数值表达式为每个参数指定相对时间，这一表达式代表某时刻之前或之后的时、分或秒数。

示例

  下面的示例使用绝对时间数的表达式。TimeSerial 函数返回中午前 6（12-6）小时前的 15分钟 （-15）， 或 5:45:00 A.M.

```vb
Dim MyTime1
MyTime1 = TimeSerial(12 - 6, -15, 0) ' 返回 5:45:00 AM.
```

  当任何一个参数的取值超出可接受的范围时，它会正确地进位到下一个较大的时间单位中。例如，如果指定了 75 分钟，则这个时间被解释成一小时十五分钟。但是，如果任何一个参数值超出 -32768 到 32767 的范围，就会导致错误。如果使用三个参数直接指定的时间或通过表达式计算出的时间超出可接受的日期范围，也会导致错误。

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#timevalue-函数)TimeValue 函数

  返回包含时间的 Date 子类型的 Variant。

TimeValue(time)



参数

| 参数 | 描述                                                         |
| ---- | ------------------------------------------------------------ |
| time | 必需。介于 0:00:00 (12:00:00 A.M.) - 23:59:59 (11:59:59 P.M.) 的时间，或任何表示此范围内时间的表达式。 |

说明

可以采用 12 或 24 小时时钟格式输入时间。例如 "2:24PM" 和 "14:24" 都是有效的 time 参数。如果 time 参数包含日期信息， TimeValue 函数并不返回日期信息。然而，如果 time 参数包含无效的日期信息，则会出现错误。

示例

下面的示例利用 TimeValue 函数将字符串转化为时间。也可以用 日期文字 直接赋时间给 Variant 类型的变量(例如， MyTime = #4:35:17 PM#).

```vb
Dim MyTime
MyTime = TimeValue("4:35:17 PM")    ' MyTime 包含 "4:35:17 PM"。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#weekday-函数)Weekday 函数

  返回代表一星期中某天的整数。返回表示一周的天数的数字，介于 1 到 7 之间。

Weekday(date, [firstdayofweek])



参数

| 参数           | 描述                                                         |
| -------------- | ------------------------------------------------------------ |
| date           | 必需。要计算的日期表达式。                                   |
| firstdayofweek | 可选。规定一周的第一天。 可采用下面的值： 0 = vbUseSystemDayOfWeek - 使用区域语言支持（NLS）API 设置 1 = vbSunday - 星期日（默认） 2 = vbMonday - 星期一 3 = vbTuesday - 星期二 4 = vbWednesday - 星期三 5 = vbThursday - 星期四 6 = vbFriday - 星期五 7 = vbSaturday - 星期六 |

示例

下面例子利用 Weekday 函数得到指定日期为星期几：

```vb
Dim MyDate, MyWeekDay
MyDate = #October 19, 1962#    ' 分派日期。
MyWeekDay = Weekday(MyDate)    ' 由于 MyWeekDay 包含 6,MyDate 代表星期五。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#weekdayname-函数)WeekDayName 函数

返回一周中指定的一天的星期名

WeekdayName(weekday, abbreviate, firstdayofweek)



参数

| 参数           | 描述                                                         |
| -------------- | ------------------------------------------------------------ |
| weekday        | 必需。一周的第几天的数字。                                   |
| abbreviate     | 可选。布尔值，指示是否缩写星期名。                           |
| firstdayofweek | 可采用下面的值： 0 = vbUseSystemDayOfWeek - 使用区域语言支持（NLS）API 设置 1 = vbSunday - 星期日（默认） 2 = vbMonday - 星期一 3 = vbTuesday - 星期二 4 = vbWednesday - 星期三 5 = vbThursday - 星期四 6 = vbFriday - 星期五 7 = vbSaturday - 星期六 |

示例

下面例子利用 WeekDayName 函数返回指定的某一天：

```vb
Dim MyDate
MyDate = WeekDayName(6, True)  'MyDate 包含 Fri。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#year-函数)Year 函数

  返回一个代表某年的整数。

Year(date)



参数

| 参数 | 描述                           |
| ---- | ------------------------------ |
| date | 必需。任何可表示日期的表达式。 |

示例

下面例子利用 Year 函数得到指定日期的年份：

```vb
Dim MyDate, MyYear
MyDate = #October 19, 1962#   '分派一日期。
MyYear = Year(MyDate)         ' MyYear 包含 1962。
```

## [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#字符函数)字符函数

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#instr-函数)InStr 函数

返回一个字符串在另一个字符串中首次出现的位置。

InStr([start, ]string1, string2[, compare])



| 参数    | 描述                                                         |
| ------- | ------------------------------------------------------------ |
| start   | 可选。规定每次搜索的起始位置。默认的搜索起始位置是第一个字符（1）。如果已规定 compare 参数，则必须有此参数。 |
| string1 | 必需。需要被搜索的字符串。                                   |
| string2 | 必需。需要搜索的字符串表达式。                               |
| compare | 可选。规定要使用的字符串比较类型。默认是 0。 可采用下列的值： 0 = vbBinaryCompare - 执行二进制比较 1 = vbTextCompare - 执行文本比较 |

InStr 函数返回以下值：

| 类型                      | 返回值               |
| ------------------------- | -------------------- |
| string1 为零长度          | 0                    |
| string1 为 Null           | Null                 |
| string2 为零长度          | start                |
| string2 为 Null           | Null                 |
| string2 没有找到          | 0                    |
| 在 string1 中找到 string2 | 找到匹配字符串的位置 |
| start > Len(string2)      | 0                    |

说明 下面的示例利用 InStr 搜索字符串：

```vb
Dim SearchString, SearchChar, MyPos
SearchString ="XXpXXpXXPXXP"   ' String to search in.
SearchChar = "P"   ' Search for "P".
MyPos = Instr(4, SearchString, SearchChar, 1)   ' A textual comparison starting at position 4. Returns 6.
MyPos = Instr(1, SearchString, SearchChar, 0)   ' A binary comparison starting at position 1. Returns 9.    
MyPos = Instr(SearchString, SearchChar)   ' Comparison is binary by default (last argument is omitted). Returns 9.
MyPos = Instr(1, SearchString, "W")   ' A binary comparison starting at position 1. Returns 0 ("W" is not found).
```

注意 InStrB 函数使用包含在字符串中的字节数据，所以 InStrB 返回的不是一个字符串在另一个字符串中第一次出现的字符位置，而是字节位置。

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#instrrev-函数)InStrRev 函数

返回字符串在另一字符串中首次出现的位置。搜索从字符串的末端开始，但是返回的位置是从字符串的起点开始计数的。

InStrRev(string1, string2[, start[, compare]])



参数

| 参数    | 描述                                                         |
| ------- | ------------------------------------------------------------ |
| string1 | 必需。需要被搜索的字符串。                                   |
| string2 | 必需。需要搜索的字符串表达式。                               |
| start   | 可选。用于设置每次搜索的开始位置。如果省略，则默认值为 -1，表示从最后一个字符的位置开始搜索。如果 start 包含 Null，则出现错误 |
| compare | 可选。规定要使用的字符串比较类型。默认是 0。 可采用下列的值： 0 = vbBinaryCompare - 执行二进制比较 1 = vbTextCompare - 执行文本比较 |

返回值

InStrRev 返回以下值：

| 类型                      | 返回值               |
| ------------------------- | -------------------- |
| string1 为零长度          | 0                    |
| string1 为 Null           | Null                 |
| string2 为零长度          | start                |
| string2 为 Null           | Null                 |
| string2 没有找到          | 0                    |
| 在 string1 中找到 string2 | 找到匹配字符串的位置 |
| start > Len(string2)      | 0                    |

说明

下面的示例利用 InStrRev 函数搜索字符串：

```vb
Dim SearchString, SearchChar, MyPos
SearchString ="XXpXXpXXPXXP"   ' String to search in.
SearchChar = "P"   ' Search for "P".
MyPos = InstrRev(SearchString, SearchChar, 10, 0)   ' A binary comparison starting at position 10. Returns 9.
MyPos = InstrRev(SearchString, SearchChar, -1, 1)   ' A textual comparison starting at the last position. Returns 12.
MyPos = InstrRev(SearchString, SearchChar, 8)   ' Comparison is binary by default (last argument is omitted). Returns 0.
```

注意 InStrRev 函数的语法与 InStr 函数的语法并不一样。

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#lcase-函数)LCase 函数

返回字符串的小写形式。

LCase(string)



| 参数   | 描述                                                         |
| ------ | ------------------------------------------------------------ |
| string | 必需。任意有效的字符串表达式。如果 string 参数中包含 Null，则返回 Null。 |

说明 仅大写字母转换成小写字母；所有小写字母和非字母字符保持不变。

下面的示例利用 LCase 函数把大写字母转换为小写字母：

```vb
Dim MyString
Dim LCaseString
MyString = "VBSCript"
LCaseString = LCase(MyString) ' LCaseString 包含 "vbscript"。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#left-函数)Left 函数

从字符串的左侧返回指定数量的字符。

Left(string, length)



| 参数   | 描述                                                         |
| ------ | ------------------------------------------------------------ |
| string | 必需。从其中返回字符的字符串。                               |
| length | 必需。规定需返回多少字符。如果设置为 0，则返回空字符串("")。如果设置为大于或等于字符串的长度，则返回整个字符串。 |

下面的示例利用Left 函数返回MyString 的左边三个字母：

```vb
Dim MyString, LeftString
MyString = "VBSCript"
LeftString = Left(MyString, 3) 'LeftString 包含 "VBS"。
```

注意 LeftB 函数与包含在字符串中字节数据一起使用。length 不是指定返回的字符串数，而是字节数。

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#len-函数)Len 函数

返回字符串内字符的数目，或是存储一变量所需的字节数。

Len(string | varname)



| 参数    | 描述                                                         |
| ------- | ------------------------------------------------------------ |
| string  | 任意有效的字符串表达式。如果 string 参数包含 Null，则返回 Null。 |
| varname | 任意有效的变量名。如果 varname 参数包含 Null，则返回 Null。  |

说明 下面的示例利用 Len 函数返回字符串中的字符数目：

```vb
Dim MyString
MyString = Len("VBSCRIPT") 'MyString 包含 8。
```

注意 LenB 函数与包含在字符串中的字节数据一起使用。LenB 不是返回字符串中的字符数，而是返回用于代表字符串的字节数。

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#ltrim-函数)LTrim 函数

删除字符串左边的空格。

LTrim(string)



| 参数   | 描述                                                         |
| ------ | ------------------------------------------------------------ |
| string | 必需。字符串表达式。如果 string 参数中包含 Null，则返回 Null。 |

说明 下面的示例利用 LTrim,用来除去字符串开始的空格

```vb
Dim MyVar
MyVar = LTrim("  vbscript ")  'MyVar 包含 "vbscript "。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#rtrim-函数)RTrim 函数

删除字符串右边的空格。

RTrim(string)



| 参数   | 描述                                                         |
| ------ | ------------------------------------------------------------ |
| string | 必需。字符串表达式。如果 string 参数中包含 Null，则返回 Null。 |

说明

```vb
Dim MyVar
MyVar = RTrim("  vbscript ")  'MyVar 包含 "  vbscript"。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#trim-函数)Trim 函数

删除字符串两边的空格。

Trim(string)



| 参数   | 描述                                                         |
| ------ | ------------------------------------------------------------ |
| string | 必需。字符串表达式。如果 string 参数中包含 Null，则返回 Null。 |

说明

```vb
Dim MyVar
MyVar = Trim("  vbscript ")   'MyVar 包含"vbscript"。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#mid-函数)Mid 函数

从字符串中返回指定数量的字符

Mid(string, start[, length])



| 参数   | 描述                                                         |
| ------ | ------------------------------------------------------------ |
| string | 必需。从其中返回字符的字符串表达式。                         |
| start  | 必需。规定起始位置。如果设置为大于字符串中的字符数量，则返回空字符串("")。 |
| length | 可选。要返回的字符数量。                                     |

下面的示例利用 Mid 函数返回字符串中从第四个字符开始的六个字符：

```vb
Dim MyVar
MyVar = Mid("VB脚本is fun!", 4, 6) 'MyVar 包含 "Script"。
```

注意 MidB 函数与包含在字符串中的字节数据一起使用。其参数不是指定字符数，而是字节数。

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#replace-函数)Replace 函数

使用另一个字符串替换字符串的指定部分指定的次数。

Replace(expression, find, replacewith[, compare[, count[, start]]])



| 参数                                  | 描述                                                         |
| ------------------------------------- | ------------------------------------------------------------ |
| string                                | 必需。被搜索的字符串。                                       |
| find                                  | 必需。将被替换的字符串部分。                                 |
| replacewith                           | 必需。用于替换的子字符串。                                   |
| start                                 | 可选。指定的开始位置。默认值是 1。起始位置之前的所有字符将被删除。 |
| count                                 | 可选。规定要执行的替换的次数。                               |
| 默认值是 -1，表示进行所有可能的替换。 |                                                              |
| compare                               | 可选。规定要使用的字符串比较类型。默认是 0。 可采用下列的值： 0 = vbBinaryCompare - 执行二进制比较 1 = vbTextCompare - 执行文本比较 |

返回值

| 情况                    | 返回值                                                    |
| ----------------------- | --------------------------------------------------------- |
| expression 为零长度     | 零长度字符串 ("")。                                       |
| expression 为 Null      | 错误。                                                    |
| find 为零长度           | expression 的副本。                                       |
| replacewith 为零长度    | expression 的副本，其中删除了所有由 find 参数指定的内容。 |
| start > Len(expression) | 零长度字符串。                                            |
| count 为 0              | expression 的副本。                                       |

说明 Replace 函数的返回值是经过替换（从由 start 指定的位置开始到 expression 字符串的结尾）后的字符串，而不是原始字符串从开始至结尾的副本。

下面的示例利用 Replace 函数返回字符串：

```vb
Dim MyString
MyString = Replace("XXpXXPXXp", "p", "Y") '二进制比较从字符串左端开始。返回 "XXYXXPXXY"。
MyString = Replace("XXpXXPXXp", "p", "Y", '文本比较从第三个字符开始。返回 "YXXYXXY"。3，, -1, 1) 
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#right-函数)Right 函数

从字符串的右侧返回指定数量的字符。

Right(string, length)



| 参数   | 描述                                                         |
| ------ | ------------------------------------------------------------ |
| string | 必需。从其中返回字符的字符串。                               |
| length | 必需。规定需返回多少字符。如果设置为 0，则返回空字符串("")。如果设置为大于或等于字符串的长度，则返回整个字符串。 |

说明 要确定 string 参数中的字符数目，使用 Len 函数。

下面的示例利用 Right 函数从字符串右边返回指定数目的字符：

```vb
Dim AnyString, MyStr
AnyString = "Hello World"      '定义字符串。
MyStr = Right(AnyString, 1)    '返回 "d"。
MyStr = Right(AnyString, 6)    ' 返回 " World"。
MyStr = Right(AnyString, 20)   ' 返回 "Hello World"。
注意 RightB 函数用于字符串中的字节数据，length 参数指定返回的是字节数目，而不是字符数目。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#space-函数)Space 函数

返回由指定数目的空格组成的字符串。

Space(number)



| 参数   | 描述                       |
| ------ | -------------------------- |
| number | 必需。字符串中的空格数量。 |

说明 下面的示例利用 Space 函数返回由指定数目空格组成的字符串：

```vb
Dim MyString
MyString = Space(10)                     ' 返回具有 10 个空格的字符串。
MyString = "Hello" & Space(10) & "World" ' 在两个字符串之间插入 10 个空格。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#strcomp-函数)StrComp 函数

比较两个字符串

StrComp(string1, string2[, compare])



| 参数    | 描述                                                         |
| ------- | ------------------------------------------------------------ |
| string1 | 必需。字符串表达式。                                         |
| string2 | 必需。字符串表达式。                                         |
| compare | 可选。规定要使用的字符串比较类型。默认是 0。 可采用下列的值： 0 = vbBinaryCompare - 执行二进制比较 1 = vbTextCompare - 执行文本比较 |

返回值 StrComp 函数有以下返回值：

| 条件                       | 返回值 |
| -------------------------- | ------ |
| string1 小于 string2       | -1     |
| string1 等于 string2       | 0      |
| string1 大于 string2       | 1      |
| string1 或 string2 为 Null | Null   |

说明 下面的示例利用 StrComp 函数返回字符串比较的结果。如果第三个参数为 1 执行文本比较；如果第三个参数为 0 或者省略执行二进制比较。

```vb
Dim MyStr1, MyStr2, MyComp
MyStr1 = "ABCD": MyStr2 = "abcd"       '定义变量。
MyComp = StrComp(MyStr1, MyStr2, 1)    ' 返回 0。
MyComp = StrComp(MyStr1, MyStr2, 0)    ' 返回 -1。
MyComp = StrComp(MyStr2, MyStr1)       ' 返回 1。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#string-函数)String 函数

返回包含指定长度的重复字符的一个字符串

String(number, character)



| 参数      | 描述                       |
| --------- | -------------------------- |
| number    | 必需。被返回字符串的长度。 |
| character | 必需。被重复的字符。       |

说明 如果指定的 character 值大于 255，则 String 使用下列公式将该数转换成有效的字符代码： character Mod 256

下面的示例利用 String 函数返回指定长度的由重复字符组成的字符串：

```vb
Dim MyString
MyString = String(5, "*")       ' 返回"*****"。
MyString = String(5, 42)        ' 返回"*****"。
MyString = String(10, "ABC")    ' 返回"AAAAAAAAAA"。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#strreverse-函数)StrReverse 函数

反转一个字符串

StrReverse(string1)



| 参数   | 描述                     |
| ------ | ------------------------ |
| string | 必需。需被反转的字符串。 |

说明 下面的示例利用 StrReverse 函数返回按相反顺序排列的字符串：

```vb
Dim MyStr
MyStr = StrReverse("VBScript") 'MyStr 包含 "tpircSBV"。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#ucase-函数)UCase 函数

字符串转换为大写。

UCase(string)



| 参数   | 描述                           |
| ------ | ------------------------------ |
| string | 必需。需被转换为大写的字符串。 |

下面的示例利用 UCase 函数返回字符串的大写形式：

```vb
Dim MyWord
MyWord = UCase("Hello World")    ' 返回"HELLO WORLD"。
```

## [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#数学函数)数学函数

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#abs-函数)Abs 函数

描述

返回数字的绝对值。数字的绝对值是其无符号的数值大小。

Abs(number)



参数

| 参数   | 描述                                                         |
| ------ | ------------------------------------------------------------ |
| number | 任意有效的数值表达式。如果 number 包含 Null，则返回 Null；如果是未初始化变量，则返回 0。 |

示例

利用 Abs 函数计算数字的绝对值：

```vb
Dim MyNumber
MyNumber = Abs(50.3 )        '返回 50.3。
MyNumber = Abs(-50.3)       '返回 50.3。 
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#atn-函数)Atn 函数

返回数值的反正切值。

Atn(number)



参数

| 参数   | 描述                             |
| ------ | -------------------------------- |
| number | 参数可以是任意有效的数值表达式。 |

示例

Atn 函数计算直角三角形两个边的比值 (number) 并返回对应角的弧度值。此比值是该角对边的长度与邻边长度之比。 结果的范围是从 -pi/2 到 pi/2 弧度。

弧度变换为角度的方法是将弧度乘以 pi/180。反之，角度变换为弧度的方法是将角度乘以180/pi 。 下面的示例利用 Atn 来计算 pi 的值:

```vb
Dim pi
pi = 4 * Atn(1)   ' 计算 pi 的值。
```

注意

注意 Atn 是 Tan（将角作为参数返回直角三角形两边的比值）的反三角函数。不要混淆 Atn 与余切（正切的倒数 (1/tangent)）函数。

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#cos-函数)Cos 函数

返回指定数字（角度）的余弦。

Cos(number)



**参数**

| 参数   | 描述                                               |
| ------ | -------------------------------------------------- |
| number | 参数可以是任何将某个角表示为弧度的有效数值表达式。 |

说明

-   Cos 函数取某个角并返回直角三角形两边的比值。此比值是直角三角形中该角的邻边长度与斜边长度之比。 结果范围在 -1 到 1 之间。
-   角度转化成弧度方法是用角度乘以 pi/180 。 反之，弧度转化成角度的方法是用弧度乘以 180/pi 。

示例

下面的示例利用 Cos 函数返回一个角的余弦值:

```vb
Dim MyAngle, MySecant
MyAngle = 1.3                ' 用弧度定义一个角。
MySecant = 1 / Cos(MyAngle)  ' 计算正割。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#exp-函数)Exp 函数

返回 e（自然对数的底）的幂次方。

Exp(number)



| 参数   | 描述                     |
| ------ | ------------------------ |
| number | 必需。有效的数值表达式。 |

说明

如果 number 参数超过 709.782712893，则出现错误。常数 e 的值约为 2.718282。 注意 Exp 函数完成 Log 函数的反运算，并且有时引用为反对数形式。

示例

下面的示例利用 Exp 函数返回 e 的幂次方:

```vb
Dim MyAngle, MyHSin  ' 用弧度定义角。
MyAngle = 1.3        '计算双曲线的正弦。
MyHSin = (Exp(MyAngle) - Exp(-1 * MyAngle)) / 2 
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#hex-函数)Hex 函数

返回表示十六进制数字值的字符串。

Hex(number)



number 参数是任意有效的表达式。

| 参数   | 描述                                                         |
| ------ | ------------------------------------------------------------ |
| number | 必需。任何有效的表达式。 如果数字是： Null - 那么 Hex 函数返回 Null。 Empty - 那么 Hex 函数返回零（0）。 Any other number - 那么 Hex 函数返回 8 个十六进制字符。 如果 number 参数不是整数，则在进行运算前将其四舍五入为最接近的整数。 |

您可以通过在数字前面添加前缀 &H 来表示十六进制数。例如，在十六进制计数法中，&H10 表示十进制数 16。

示例

下面的示例利用 Hex 函数返回数字的十六进制数：

```vb
Dim MyHex
MyHex = Hex(5)      ' 返回 5。
MyHex = Hex(10)   ' 返回A。
MyHex = Hex(459)   ' 返回 1CB。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#int-函数)Int 函数

返回指定数字的整数部分。

-   删除number 参数的小数部分并返回以整数表示的结果。
-   如果 number 参数为负数时，Int 函数返回小于或等于 number 的第一个负整数，例如，Int 将 -8.4 转换为 -9，

Int(number)



| 参数   | 描述                                                         |
| ------ | ------------------------------------------------------------ |
| number | 必需。参数可以是任意有效的数值表达式。如果 number 参数包含 Null，则返回 Null 。 |

示例

下面的示例说明 Int 和 Fix 函数如何返回数字的整数部分：

```vb
MyNumber = Int(99.8)    ' 返回 99。
MyNumber = Int(-99.8)   ' 返回 -100。
MyNumber = Int(-99.2)   ' 返回 -100。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#fix-函数)Fix 函数

返回指定数字的整数部分。

-   删除 number 参数的小数部分并返回以整数表示的结果。
-   如果 number 参数为负数时， Fix 函数返回大于或等于 number 参数的第一个负整数。例如， Fix 将 -8.4 转换为 -8。

Fix(number)



参数

| 参数   | 描述                                                         |
| ------ | ------------------------------------------------------------ |
| number | 必需。参数可以是任意有效的数值表达式。如果 number 参数包含 Null，则返回 Null 。 |

Fix(number) 等同于： Sgn(number) * Int(Abs(number))

示例

下面的示例说明 Int 和 Fix 函数如何返回数字的整数部分：

```vb
MyNumber = Fix(99.2)    ' 返回 99。
MyNumber = Fix(-99.8)   ' 返回-99。
MyNumber = Fix(-99.2)   ' 返回 -99。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#log-函数)Log 函数

返回数值的自然对数。

Log(number)



| 参数   | 描述                              |
| ------ | --------------------------------- |
| number | 必需。大于 0 的有效的数值表达式。 |

说明 自然对数是以 e 为底的对数。常数 e 的值约为 2.718282。

用 n 的自然对数除 x 的自然对数，可以得到以 n 为底的 x 的对数。如下所示：

Logn(x) = Log(x) / Log(n)

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#oct-函数)Oct 函数

返回表示数字八进制值的字符串。 如果 number 参数不是整数，则在进行运算前，将其四舍五入到最接近的整数。

Oct(number)



参数

| 参数   | 描述                                                         |
| ------ | ------------------------------------------------------------ |
| number | 必需。任何有效的表达式。如果数字是： Null - 那么 Oct 函数返回 Null。 Empty - 那么 Oct 函数返回零（0）。 Any other number - 那么 Oct 函数返回 11 个八进制字符。 |

用户也可以通过直接在数前加上 &O 表示八进制数。例如，&O10 为十进制数 8 的八进制表示法。

示例

下面的示例利用 Oct 函数返回数值的八进制数：

```vb
Dim MyOct
MyOct = Oct(4)     ' 返回 4。
MyOct = Oct(8)     ' 返回 10。
MyOct = Oct(459)   ' 返回 713。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#rnd-函数)Rnd 函数

返回一个随机数。数字总是小于 1 但大于或等于 0 。

Rnd(number)



参数

| 参数                                                         | 描述                     |
| ------------------------------------------------------------ | ------------------------ |
| number                                                       | 可选。有效的数值表达式。 |
| 如果数字是： <0 - Rnd 会每次都返回相同的数字。 >0 - Rnd 会返回序列中的下一个随机数。 =0 - Rnd 会返回最近生成的数。 省略 - Rnd 会返回序列中的下一个随机数。 |                          |

因每一次连续调用 Rnd 函数时都用序列中的前一个数作为下一个数的种子，所以对于任何最初给定的种子都会生成相同的数列。 在调用 Rnd 之前，先使用无参数的 Randomize 语句初始化随机数生成器，该生成器具有基于系统计时器的种子。

要产生指定范围的随机整数，请使用以下公式：

Int((upperbound - lowerbound + 1) * Rnd + lowerbound) 这里， upperbound 是此范围的上界，而 lowerbound 是此范围内的下界。

注意 要重复随机数的序列，请在使用数值参数调用 Randomize 之前，立即用负值参数调用 Rnd。使用同样 number 值的 Randomize 不能重复先前的随机数序列。

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#sgn-函数)Sgn 函数

返回表示数字符号的整数。

Sgn(number)



参数

| 参数   | 描述                                                         |
| ------ | ------------------------------------------------------------ |
| number | 必需。参数可以是任意有效的数值表达式。 如果 number 为 Sgn 返回 大于零 1 等于零 0 小于零 -1 |

说明

number 参数的符号决定 Sgn 函数的返回值。

下面的示例利用 Sgn 函数决定数值的符号：

示例

```vb
Dim MyVar1, MyVar2, MyVar3, MySign
MyVar1 = 12: MyVar2 = -2.4: MyVar3 = 0
MySign = Sgn(MyVar1)    ' 返回 1。
MySign = Sgn(MyVar2)    ' 返回 -1。
MySign = Sgn(MyVar3)    ' 返回 0。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#sin-函数)Sin 函数

返回指定数字（角度）的正弦值。

Sin(number)



参数

| 参数   | 描述                                           |
| ------ | ---------------------------------------------- |
| number | 必需。任何将某个角表示为弧度的有效数值表达式。 |

示例

下面例子利用 Sin 返回角度的正弦：

```vb
Dim MyAngle, MyCosecant
MyAngle = 1.3                   ' 用弧度定义角度。
MyCosecant = 1 / Sin(MyAngle)   '计算余割。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#sqr-函数)Sqr 函数

返回数值的平方根。

Sqr(number)



参数

| 参数   | 描述                              |
| ------ | --------------------------------- |
| number | 必需。大于 0 的有效的数值表达式。 |

说明

下面的示例利用 Sqr 函数计算数值的平方根：

```vb
Dim MySqr
MySqr = Sqr(4)     ' 返回 2。
MySqr = Sqr(23)    ' 返回4.79583152331272。
MySqr = Sqr(0)     ' 返回0。
MySqr = Sqr(-4)    ' 产生实时错误。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#tan-函数)Tan 函数

返回指定数字（角度）的正切。

Tan(number)



参数

| 参数   | 描述                                                     |
| ------ | -------------------------------------------------------- |
| number | 必需。参数可以是任何将某个角表示为弧度的有效数值表达式。 |

说明

Tan 取某个角并返回直角三角形两个直角边的比值。此比值是直角三角形中该角的对边长度与邻边长度之比。 将角度乘以 pi /180 即可转换为弧度，将弧度乘以 180/pi 即可转换为角度。 下面的示例利用 Tan 函数返回角度的正切：

示例

```vb
Dim MyAngle, MyCotangent
MyAngle = 1.3                     ' 用弧度定义角度。
MyCotangent = 1 / Tan(MyAngle)    ' 计算余切。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#round-函数)Round 函数

返回按指定位数进行四舍五入的数值。

Round(expression[, numdecimalplaces])



参数

| 参数             | 描述                                                   |
| ---------------- | ------------------------------------------------------ |
| expression       | 必需。需要被四舍五入的数值表达式。                     |
| numdecimalplaces | 可选。规定对小数点右边的多少位进行四舍五入。默认是 0。 |

示例

下面的示例利用 Round 函数将数值四舍五入到两位小数：

```vb
Dim MyVar, pi
pi = 3.14159
MyVar = Round(pi, 2) 'MyVar contains 3.14。
```

## [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#数组函数)数组函数

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#array-函数)Array 函数

  用于创建一个数组，并返回Variant数据类型的变量。

Array(arglist)



参数

| 参数    | 描述                                                         |
| ------- | ------------------------------------------------------------ |
| arglist | 数组元素值的集合，使用逗号分隔，如果没有指定此参数，则将会创建零长度的数组 |

示例

```vb
'创建名为 A 的变量
Dim A

'一个数组赋值给变量 A
A = Array(10,20,30)

'获取第二个数组元素的结果。
B = A(2)  ' B is now 30。
```

注意

未作为数组声明的变量仍可以包含数组。虽然包含数组的 Variant 变量与包含 Variant 元素的数组变量有概念上的不同，但访问数组元素的方法是相同的。

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#filter-函数)Filter 函数

  返回下标从零开始的数组，其中包含以特定过滤条件为基础的字符串数组的子集。

Filter(InputStrings, Value[, Include[, Compare]])



-   如果在数组中没有找到匹配的字符串，Filter 将返回空数组。如果数组为 Null 或者不是一维数组，则会发生错误。
-   返回的数组仅包含足以包含匹配项数目的元素。

参数

| 参数         | 描述                                                         |
| ------------ | ------------------------------------------------------------ |
| inputstrings | 必需。要检索的一维字符串数组。                               |
| value        | 必需。要搜索的字符串。                                       |
| include      | 可选。Boolean 值，指定返回的子字符串是否包含 Value。如果 Include 为 True，Filter 将返回包含子字符串 Value 的数组子集。如果 Include 为 False，Filter 将返回不包含子字符串 Value 的数组子集。默认值为 True。 |
| compare      | 可选。规定要使用的字符串比较类型。 可采用下列的值： 0 = vbBinaryCompare - 执行二进制比较 1 = vbTextCompare - 执行文本比较 |

示例

   下面例子利用 Filter 函数返回包含搜索条件 "Mon" 的数组:

```vb
Dim MyIndex
Dim MyArray (3)
MyArray(0) = "Sunday"
MyArray(1) = "Monday"
MyArray(2) = "Tuesday"
MyIndex = Filter(MyArray, "Mon") 'MyIndex(0) 包含 "Monday"。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#isarray-函数)IsArray 函数

  返回一个指定的变量是否为数组的布尔值。如果变量为数组，则返回 True，否则返回 False。

IsArray(varname)



参数

| 参数    | 描述              |
| ------- | ----------------- |
| varname | 必需。任意变量 。 |

示例

  下面的示例利用 IsArray 函数验证 MyVariable 是否为一数组：

```vb
Dim MyVariable
Dim MyArray(3)
MyArray(0) = "Sunday"
MyArray(1) = "Monday"
MyArray(2) = "Tuesday"
MyVariable = IsArray(MyArray) ' MyVariable 包含 "True"。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#join-函数)Join 函数

返回一个由数组中若干子字符串组成的字符串

Join(list[，delimiter])



参数

| 参数      | 描述                                                         |
| --------- | ------------------------------------------------------------ |
| list      | 必需。一维数组，其中包含需被连接的子字符串。                 |
| delimiter | 可选。用于在返回的字符串中分割子字符串的字符。默认是空格字符。 |

示例

下面的示例利用 Join 函数联合 MyArray 的子字符串：

```vb
Dim MyString
Dim MyArray(3)
MyArray(0) = "Mr."
MyArray(1) = "John "
MyArray(2) = "Doe "
MyArray(3) = "III"
MyString = Join(MyArray) 'MyString 包含 "Mr. John Doe III"。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#lbound-函数)LBound 函数

返回指定数组维的最小可用下标。

LBound(arrayname[, dimension])



参数

| 参数      | 描述                                                         |
| --------- | ------------------------------------------------------------ |
| arrayname | 必需。数组变量的名称。                                       |
| dimension | 可选。要返回哪一维的下界。 1 = 第一维， 2 = 第二维，以此类推。默认是 1 。 |

说明

  LBound 函数与 UBound 函数共同使用以确定数组的大小。使用 UBound 函数可以找到数组某一维的上界。 任一维的下界都是 0。

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#split-函数)Split 函数

  返回基于 0 的一维数组，此数组包含指定数量的子字符串。

Split(expression[, delimiter[, count[, start]]])



参数

| 参数       | 描述                                                         |
| ---------- | ------------------------------------------------------------ |
| expression | 必需。包含子字符串和分隔符的字符串表达式。                   |
| delimiter  | 可选。用于识别子字符串界限的字符。默认是空格字符。           |
| count      | 可选。需被返回的子字符串的数目。-1 指示返回所有的子字符串。  |
| compare    | 可选。规定要使用的字符串比较类型。 可采用下列的值： 0 = vbBinaryCompare - 执行二进制比较 1 = vbTextCompare - 执行文本比较 |

示例

下面的示例利用 Split 函数从字符串中返回数组。函数对分界符进行文本比较，返回所有的子字符串。

```vb
Dim MyString, MyArray, Msg
MyString = "VBScriptXisXfun!"
MyArray = Split(MyString, "x", -1, 1)
' MyArray(0) contains "VBScript".
' MyArray(1) contains "is".
' MyArray(2) contains "fun!".
Msg = MyArray(0) & " " & MyArray(1)
Msg = Msg   & " " & MyArray(2)
MsgBox Msg
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#ubound-函数)UBound 函数

指定数组维数的最大可用下标。

UBound(arrayname[, dimension])



参数

| 参数      | 描述                                                         |
| --------- | ------------------------------------------------------------ |
| arrayname | 必需。数组变量的名称。                                       |
| dimension | 可选。要返回哪一维的下界。 1 = 第一维， 2 = 第二维，以此类推。默认是 1 。 |

说明

UBound 函数与 LBound 函数一起使用，用于确定数组的大小。使用 LBound 函数可以确定数组某一维的下界。 所有维的下界均为 0。

示例

```vb
Dim A(100,3,4)
语句 返回值 
UBound(A, 1) 100 
UBound(A, 2) 3 
UBound(A, 3) 4 
```

## [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#格式化函数)格式化函数

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#formatcurrency-函数)FormatCurrency 函数

  返回作为货币值被格式化的表达式，使用计算机系统控制面板中定义的货币符号。

FormatCurrency(expression[,NumDigitsAfterDecimal [,IncludeLeadingDigit [,UseParensForNegativeNumbers [,GroupDigits]]]])



参数

| 参数            | 描述                                                         |
| --------------- | ------------------------------------------------------------ |
| expression      | 必需。需被格式化的表达式。                                   |
| NumDigAfterDec  | 可选。指示小数点右侧显示位数的数值。默认值为 -1（使用的是计算机的区域设置）。 |
| IncLeadingDig   | 可选。指示是否显示小数值的前导零：  -2 = TristateUseDefault - 使用计算机的区域设置  -1 = TristateTrue - True   0 = TristateFalse - False |
| UseParForNegNum | 可选。指示是否将负值置于括号中：  -2 = TristateUseDefault - 使用计算机的区域设置  -1 = TristateTrue - True   0 = TristateFalse - False |
| GroupDig        | 可选。指示是否使用计算机区域设置中指定的数字分组符号将数字分组：  -2 = TristateUseDefault - 使用计算机的区域设置  -1 = TristateTrue - True  0 = TristateFalse - False |

说明

\* 当省略一个或多个可选项参数时，由计算机区域设置提供被省略参数的值。 与货币值相关的货币符号的位置由系统的区域设置决定。 * 除“显示起始的零”设置来自区域设置的“数字”附签外，所有其他设置信息均取自区域设置的“货币”附签。 :::

示例

  下面例子利用 FormatCurrency 函数把 expression 格式化为 currency 并且赋值给 MyCurrency: ```VB Dim MyCurrency MyCurrency = FormatCurrency(1000) 'MyCurrency 包含 $1000.00 。 ```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#formatdatetime-函数-2)FormatDateTime 函数

  返回表达式，此表达式已被格式化为日期或时间。

FormatDateTime(Date[, NamedFormat])



参数

| 参数   | 描述                                                         |
| ------ | ------------------------------------------------------------ |
| date   | 必需。任何有效的日期表达式.比如 Date() 或者 Now()。          |
| format | 可选。规定所使用的日期/时间格式的格式值。  可采用下面的值：  0 = vbGeneralDate - 默认。返回日期：mm/dd/yy 及如果指定时间：hh:mm:ss PM/AM。  1 = vbLongDate - 返回日期：weekday, monthname, year  2 = vbShortDate - 返回日期：mm/dd/yy  3 = vbLongTime - 返回时间：hh:mm:ss PM/AM  4 = vbShortTime - 返回时间：hh:mm |

示例

  下面例子利用 FormatDateTime 函数把表达式格式化为长日期型并且把它赋给 MyDateTime:

```vb
Function GetCurrentDate 
  'FormatDateTime 把日期型格式化为长日期型。
  GetCurrentDate = FormatDateTime(Date, 1) 
End Function
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#formatnumber-函数)FormatNumber 函数

  返回表达式，此表达式已被格式化为数值。

FormatNumber(expression [,NumDigitsAfterDecimal [,IncludeLeadingDigit [,UseParensForNegativeNumbers [,GroupDigits]]]])



参数

| 参数            | 描述                                                         |
| --------------- | ------------------------------------------------------------ |
| expression      | 必需。需被格式化的表达式。                                   |
| NumDigAfterDec  | 可选。指示小数点右侧显示位数的数值。默认值为 -1（使用的是计算机的区域设置）。 |
| IncLeadingDig   | 可选。指示是否显示小数值的前导零：  -2 = TristateUseDefault - 使用计算机的区域设置  -1 = TristateTrue - True  0 = TristateFalse - False |
| UseParForNegNum | 可选。指示是否将负值置于括号中：  -2 = TristateUseDefault - 使用计算机的区域设置  -1 = TristateTrue - True  0 = TristateFalse - False |
| GroupDig        | 可选。指示是否使用计算机区域设置中指定的数字分组符号将数字分组：  -2 = TristateUseDefault - 使用计算机的区域设置  -1 = TristateTrue - True  0 = TristateFalse - False |

说明

\* 当省略一个或多个可选项参数时，由计算机区域设置提供被省略参数的值。 * 注意 所有设置信息均取自区域设置的“数字”附签。 :::

示例

  下面例子利用 FormatNumber 函数把数值格式化为带四位小数点的数:

```vb
Function FormatNumberDemo 
    Dim MyAngle, MySecant, MyNumber
    MyAngle = 1.3                ' 用弧度定义角。
    MySecant = 1 / Cos(MyAngle)  ' 计算正割值。
    FormatNumberDemo = FormatNumber(MySecant,4) ' 把 MySecant 格式化为带四位小数点的数。
End Function
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#formatpercent-函数)FormatPercent 函数

  返回表达式，此表达式已被格式化为尾随有 % 符号的百分比（乘以 100 ）。

FormatPercent(expression[,NumDigitsAfterDecimal [,IncludeLeadingDigit [,UseParensForNegativeNumbers [,GroupDigits]]]])



参数

| 参数            | 描述                                                         |
| --------------- | ------------------------------------------------------------ |
| expression      | 必需。需被格式化的表达式。                                   |
| NumDigAfterDec  | 可选。指示小数点右侧显示位数的数值。默认值为 -1（使用的是计算机的区域设置）。 |
| IncLeadingDig   | 可选。指示是否显示小数值的前导零：br> -2 = TristateUseDefault - 使用计算机的区域设置   -1 = TristateTrue - True   0 = TristateFalse - False |
| UseParForNegNum | 可选。指示是否将负值置于括号中：   -2 = TristateUseDefault - 使用计算机的区域设置  -1 = TristateTrue - True   0 = TristateFalse - False |
| GroupDig        | 可选。指示是否使用计算机区域设置中指定的数字分组符号将数字分组：   -2 = TristateUseDefault - 使用计算机的区域设置   -1 = TristateTrue - True   0 = TristateFalse - False |

说明

\* 当省略一个或多个可选项参数时，由计算机区域设置提供被省略参数的值。 * 注意 所有设置信息均取自区域设置的“数字”附签。 :::

示例

  下面例子利用 FormatPercent 函数把表达式格式化为百分数:

```vb
Dim MyPercent
MyPercent = FormatPercent(2/32) 'MyPercent 包含 6.25%。
```

## [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#类型转换函数)类型转换函数

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#asc-函数)Asc 函数

字符串表达式的第一个字符的 ANSI 编码。

Asc(string)



**参数**

| 参数   | 描述                                                         |
| ------ | ------------------------------------------------------------ |
| string | 任意有效的字符串表达式。如果 string 参数未包含字符，则将发生运行时错误。 |

**示例**

下面例子中, Asc 返回每一个字符串首字母的 ANSI 字符代码:

```vb
Dim MyNumber
MyNumber = Asc("A")      '返回 65。
MyNumber = Asc("a")      '返回 97。
MyNumber = Asc("Apple")  '返回 65。
```

注意

注意 AscB 函数和包含字节数据的字符串一起使用。 AscB 不是返回第一个字符的字符代码，而是返回首字节。 AscW 是为使用 Unicode 字符的 32 位平台提供的。 它返回 Unicode （宽型）字符代码，因此可以避免从 ANSI 到 Unicode 的代码转换。

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#cbool-函数)CBool 函数

把表达式转换为布尔（Boolean）类型。

CBool(expression)



**参数**

| 参数       | 描述                                                         |
| ---------- | ------------------------------------------------------------ |
| expression | 是任意有效的表达式。 如果 expression 是零，则返回 False；否则返回 True。如果 expression 不能解释为数值，则将发生运行时错误。 |

**示例**

下面的示例使用 CBool 函数将一个表达式转变成 Boolean 类型。如果表达式所计算的值非零，则 CBool 函数返回 True；否则返回 False。

```vb
Dim A, B, Check
A = 5: B = 5            ' 初始化变量。
Check = CBool(A = B)    '复选框设为 True 。
A = 0                   '定义变量。 
Check = CBool(A)        '复选框设为 False 。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#cbyte-函数)CByte 函数

把表达式转换为字节（Byte）类型。

CByte(expression)



**参数**

| 参数       | 描述                 |
| ---------- | -------------------- |
| expression | 是任意有效的表达式。 |

**说明**

-   通常，可以使用子类型转换函数书写代码，以显示某些操作的结果应被表示为特定的数据类型，而不是默认类型。例如，在出现货币、单精度、双精度或整数运算的情况下，使用 CByte 强制执行字节运算。
-   CByte 函数用于进行从其他数据类型到 Byte 子类型的的国际公认的格式转换。例如对十进制分隔符（如千分符）的识别，可能取决于系统的区域设置。
-   如果 expression 在 Byte 子类型可接受的范围之外，则发生错误。

**示例** 下面的示例利用 CByte 函数把 expression 转换为 byte

```vb
Dim MyDouble, MyByte
MyDouble = 125.5678         ' MyDouble 是一个双精度值。
MyByte = CByte(MyDouble)    ' MyByte 包含 126 。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#ccur-函数)CCur 函数

把表达式转换为货币（Currency）类型。

CCur(expression)



**参数**

| 参数       | 描述               |
| ---------- | ------------------ |
| expression | 任意有效的表达式。 |

**示例** 下面的示例使用 CCur 函数将一个表达式转换成 Currency 类型:

```vb
Dim MyDouble, MyCurr
MyDouble = 543.214588          ' MyDouble 是双精度的。
MyCurr = CCur(MyDouble * 2)    '把 MyDouble * 2 (1086.429176) 的结果转换为 Currency (1086.4292)。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#cdbl-函数)CDbl 函数

把表达式转换为双精度（Double）类型。

CDbl(expression)



**参数**

| 参数       | 描述               |
| ---------- | ------------------ |
| expression | 任意有效的表达式。 |

**说明**

-   通常，您可以使用子类型数据转换函数书写代码，以显示某些操作的结果应当被表达为特定的数据类型，而非默认的数据类型。例如在出现货币或整数运算的情况下，使用 CDbl 或 CSng 函数强制进行双精度或单精度算术运算。
-   CDbl 函数用于进行从其他数据类型到 Double 子类型的国际公认的格式转换。例如，十进制分隔符和千位分隔符的识别取决于系统的区域设置。

**示例** 下面的示例利用 CDbl 函数把 expression 转换为 Double。

```vb
 Dim MyCurr, MyDouble
MyCurr = CCur(234.456784)               ' MyCurr 是 Currency 型 (234.4567)。
MyDouble = CDbl(MyCurr * 8.2 * 0.01)    ' 把结果转换为 Double 型 (19.2254576)。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#chr-函数)Chr 函数

返回与指定的 ANSI 字符代码相对应的字符。

Chr(charcode)



**参数**

| 参数     | 描述                       |
| -------- | -------------------------- |
| charcode | 参数是可以标识字符的数字。 |

说明

-   从 0 到 31 的数字表示标准的不可打印的 ASCII 代码。例如，Chr(10) 返回换行符。

**示例** 下面例子利用 Chr 函数返回与指定的字符代码相对应的字符:

```vb
Dim MyChar
MyChar = Chr(65)    '返回 A。
MyChar = Chr(97)    '返回 a。
MyChar = Chr(62)    '返回 >。
MyChar = Chr(37)    '返回 %。
```

注意

ChrB 函数与包含在字符串中的字节数据一起使用。ChrB 不是返回一个或两个字节的字符，而总是返回单个字节的字符。ChrW 是为使用 Unicode 字符的 32 位平台提供的。它的参数是一个 Unicode (宽字符)的字符代码，因此可以避免将 ANSI 转化为 Unicode 字符。

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#cint-函数)CInt 函数

把表达式转换为整数（Integer）类型。

CInt(expression)



**参数**

| 参数       | 描述                                                         |
| ---------- | ------------------------------------------------------------ |
| expression | 参数是任意有效的表达式。expression值必须是介于 -32768 与 32767 之间的数字。否者则发生错误。 |

**示例**

下面的示例利用 CInt 函数把值转换为 Integer:

```vb
Dim MyDouble, MyInt
MyDouble = 2345.5678      ' MyDouble 是 Double。
MyInt = CInt(MyDouble)    ' MyInt 包含 2346。
```

注意

CInt 不同于 Fix 和 Int 函数删除数值的小数部分，而是采用四舍五入的方式。 当小数部分正好等于 0.5 时, CInt 总是将其四舍五入成最接近该数的偶数。例如， 0.5 四舍五入为 0, 以及 1.5 四舍五入为 2.

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#clng-函数)CLng 函数

把表达式转换为长整型（Long）类型。

CLng(expression)



**参数**

| 参数       | 描述                                                         |
| ---------- | ------------------------------------------------------------ |
| expression | 参数是任意有效的表达式。expression值必须是介于 -2147483648 与 2147483647 之间的数字。否者则发生错误。 |

说明 通常，您可以使用子类型数据转换函数书写代码，以显示某些操作的结果应当被表达为特定的数据类型，而非默认的数据类型。例如，在出现货币运算、单精度或双精度算术运算的情况下，使用 CInt 或 CLng 函数强制进行整数运算。

CLng 函数用于进行从其他数据类型到 Long 子类型的的国际公认的格式转换。例如，对十进制分隔符和千位分隔符的识别取决于系统的区域设置。

如果 expression 取值不在 Long子类型的允许范围内，则会出现错误。 **示例**

下面的示例利用 CLng 函数把值转换为 Long:

```vb
Dim MyVal1, MyVal2, MyLong1, MyLong2

' MyVal1, MyVal2 是双精度值。
MyVal1 = 25427.45
MyVal2 = 25427.55       
MyLong1 = CLng(MyVal1)          ' MyLong1 25427。
MyLong2 = CLng(MyVal2)           ' MyLong2 包含 25428 。
```

注意

CLng 不同于 Fix 和 Int 函数删除小数部分， 而是采用四舍五入的方式。 当小数部分正好等于 0.5 时， CLng 函数总是将其四舍五入为最接近该数的偶数。如， 0.5 四舍五入为 0, 以及 1.5 四舍五入为 2 。

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#csng-函数)CSng 函数

把表达式转换为单精度（Single）类型。

CSng(expression)



**参数**

| 参数       | 描述                                                         |
| ---------- | ------------------------------------------------------------ |
| expression | 必需。任何有效的表达式。包含单精度浮点数，负数范围从 -3.402823E38 到 -1.401298E-45，正数范围从 1.401298E-45 到 3.402823E38。 超出此范围则发生错误。 |

下面的示例利用 CSng 函数把值转换为 Single:

```vb
Dim MyDouble1, MyDouble2, MySingle1, MySingle2  ' MyDouble1, MyDouble2 是双精度值。
MyDouble1 = 75.3421115: MyDouble2 = 75.3421555
MySingle1 = CSng(MyDouble1)  ' MySingle1 包含 75.34211 。
MySingle2 = CSng(MyDouble2)  ' MySingle2 包含 75.34216 。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#cstr-函数)CStr 函数

表达式转换为字符串（String）类型。

CStr(expression)



| 参数       | 描述                                                         |
| ---------- | ------------------------------------------------------------ |
| expression | 必需。任何有效的表达式。                                     |
|            | 如果表达式是： Boolean - CStr 函数将返回一个字符串，其中包含 true 或 false。 Date - CStr 函数将返回一个字符串，其中包含短日期格式的日期。 Null - 将发生 run-time 错误。 Empty - CStr 函数将返回一个空字符串（""）。 Error - CStr 函数将返回一个字符串，其中包含单词 "Error" 和错误号码。 Other numeric - CStr 函数将返回一个字符串，其中包含数字。 |

下面的示例利用 CStr 函数把数字转换为 String:

```vb
Dim MyDouble, MyString
MyDouble = 437.324         ' MyDouble 是双精度值。
MyString = CStr(MyDouble)  ' MyString 包含 "437.324"。
```

## [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#其他函数)其他函数

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#createobject)CreateObject

创建并返回对 Automation 对象的引用。 `Automation`服务器至少提供一种对象类型。例如，字处理应用程序可以提供应用程序对象、文档对象和工具条对象。

CreateObject(servername.typename [, location])



参数

| 参数       | 描述                                 |
| ---------- | ------------------------------------ |
| servername | 必需。提供此对象的应用程序名称。     |
| typename   | 必需。对象的类型或类（type/class）。 |
| location   | 可选。在何处创建对象。               |

说明

要创建 Automation 对象，将 CreateObject 函数返回的对象赋值给某对象变量：

```vb
Dim ExcelSheet
Set ExcelSheet = CreateObject("Excel.Sheet")
```

上述代码启动创建对象（在此实例中，是 Microsoft Excel 电子表格）的应用程序。对象创建后，就可以在代码中使用定义的对象变量引用此对象。在下面的示例中，可使用对象变量、ExcelSheet 和其他 Excel 对象，包括 Application 对象和 Cells 集合访问新对象的属性和方法。例如：

```vb
' 通过 Application 对象使 Excel 可见。
ExcelSheet.Application.Visible = True
' 在工作表的第一个单元中放置文本。
ExcelSheet.Cells(1,1).Value = "这是 A 列第一行"
' 保存工作表。
ExcelSheet.SaveAs "C:\DOCS\TEST.XLS"
' 在 Application 对象中使用 Quit 方法退出 Excel。
ExcelSheet.Application.Quit
' 释放对象变量。
Set ExcelSheet = Nothing
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#eval)Eval

计算一个表达式的值并返回结果。

Eval(expression)



| 参数       | 描述                                                         |
| ---------- | ------------------------------------------------------------ |
| result     | 可选项。 是一个变量，用于接受返回的结果。如果未指定结果，应考虑使用 Execute 语句代替。 |
| expression | 必需。可以是包含任何有效 VBScript 表达式的字符串。           |

说明

在 VBScript 中，x = y 可以有两种解释。第一种方式是赋值语句，将 y 的值赋予 x。第二种解释是测试 x 和 y 是否相等。如果相等，result 为 True；否则 result 为 False。Eval 方法总是采用第二种解释，而 Execute 语句总是采用第一种。
在Microsoft(R) Visual Basic Scripting Edition 中不存在这种比较与赋值的混淆，因为赋值运算符(=)与比较运算符 (==)不同。

示例

下面的例子说明了 Eval 函数的用法：

```vb
Sub GuessANumber
  Dim Guess, RndNum
  RndNum = Int((100) * Rnd(1) + 1)
  Guess = CInt(InputBox("Enter your guess:",,0))
  Do
    If Eval("Guess = RndNum") Then
      MsgBox "祝贺你！猜对了！"
      Exit
    Else
      Guess = CInt(InputBox("对不起，请再试一次",,0))
    End If
  Loop Until Guess = 0
End Sub
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#getlocale)GetLocale

返回当前区域设置 ID 值。

-   locale 是用户参考信息集合，与用户的语言、国家和文化传统有关。locale 决定键盘布局、字母排序顺序和日期、时间、数字与货币格式。 返回值可以是任意一个 32-位 的值

GetLocale()



**示例** 下面举例说明 GetLocale 函数的用法。

```vb
Dim currentLocale
currentLocale = GetLocale() ' 返回2052  对应结果：中文 - 中华人民共和国 zh-cn 
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#getobject)GetObject

返回对文件中 Automation 对象的引用。

GetObject([pathname] [, class])



| 参数     | 描述                                                         |
| -------- | ------------------------------------------------------------ |
| pathname | 可选。包含 automation 对象的文件的完整路径和名称。如果此参数被忽略，就必须有 class 参数。 |
| class    | 可选。automation 对象的类。此参数使用此语法：appname.objectype。 |

class 参数的语法格式为 appname.objectype，其中包括以下部分：

| 部分      | 描述                                   |
| --------- | -------------------------------------- |
| appname   | 必选。字符串，提供对象的应用程序名称。 |
| objectype | 必选。字符串，要创建的对象的类型或类。 |

示例

使用 GetObject 函数可以访问文件中的 Automation 对象，而且可以将该对象赋值给对象变量。使用 Set 语句将 GetObject 返回的对象赋值给对象变量。例如：

```vb
Dim CADObject
Set CADObject = GetObject("C:\CAD\SCHEMA.CAD")
```

在执行上述代码时，就会启动与指定路径名相关联的应用程序，同时激活指定文件中的对象。如果 pathname 是零长度字符串 ("")，GetObject 返回指定类型的新对象实例。如果省略 pathname 参数，GetObject 将返回指定类型的当前活动对象。如果没有指定类型的对象，就会出现错误。

某些应用程序允许只激活文件的一部分，方法是在文件名后加上一个惊叹号 (!) 以及用于标识要激活的文件部分的字符串。有关创建这种字符串的详细信息，请参阅创建对象的应用程序的有关文档。

例如，在绘图应用程序中，一个存放在文件中的图可能有多层。可以使用下述代码来激活图 SCHEMA.CAD 中的某一层：

```vb
Set LayerObject = GetObject("C:\CAD\SCHEMA.CAD!Layer3")
```

如果没有指定对象的类，则 Automation 会根据所提供的文件名，确定要启动的应用程序以及要激活的对象。但是，有些文件可能支持多个对象类。例如，图可能支持三种不同类型的对象：Application 对象、Drawing 对象和 Toolbar 对象，所有这些都是同一个文件中的一部分。使用可选项的 class 参数可以指定文件中要激活的对象。例如：

```vb
Dim MyObject
Set MyObject = GetObject("C:\DRAWINGS\SAMPLE.DRW", "FIGMENT.DRAWING") 
```

在上述样例中，FIGMENT 是绘图应用程序的名称，而 DRAWING 则是它支持的一种对象类型。对象被激活之后，就可以在代码中使用所定义的对象变量来引用它。在前面的例子中，可以使用对象变量 MyObject 访问新对象的属性和方法。例如：

```vb
MyObject.Line 9, 90
MyObject.InsertText 9, 100, "嗨，你好！"
MyObject.SaveAs "C:\DRAWINGS\SAMPLE.DRW"
```

注意 在对象的当前实例存在，或者要用已加载的文件创建对象时，请使用 GetObject 函数。如果没有当前实例，并且不准备使用已加载的文件启动对象，请使用 CreateObject 函数。

如果对象已注册为单个实例的对象，则无论执行多少次 CreateObject，都只能创建该对象的一个实例。若使用单个实例对象，当使用零长度字符串 ("") 语法调用时，GetObject 总是返回同一个实例，而如果省略 pathname 参数，则会出现错误。

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#getref)GetRef

返回一个指向一过程的引用，此过程可绑定某事件。

Set object.eventname = GetRef(procname)



| 参数     | 描述                                                         |
| -------- | ------------------------------------------------------------ |
| object   | 必需。事件所关联的 HTML 对象的名称。                         |
| event    | 必需。要与函数绑定的事件的名称。                             |
| procname | 必需。该字符串中包含 Sub 或 Function 过程的名称，该过程与事件关联。 |

说明 GetRef 函数可以用来将 VBScript 过程 (Function 或 Sub） 与 DHTML (动态 HTML)页面中可用的任何事件联系在一起。DHTML 对象模型为不同对象提供了与各种可用事件有关的信息。

在其他脚本和程序设计语言中，GetRef 所提供的功能被称为函数指针，即它指向了在指定事件发生时要执行的过程的地址。

下面的例子说明了 GetRef 函数的使用：

```vb
Function GetRefTest()
  Dim Splash
  Splash = "GetRefTest Version 1.0"  & vbCrLf
  Splash = Splash & Chr(169) & " YourCompany 1999 "
  MsgBox Splash
End Function
Set Window.Onload = GetRef("GetRefTest")
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#inputbox)InputBox

显示一个对话框，用户可在其中输入文本并/或点击一个按钮。如果用户点击点击 OK 按钮或按键盘上的 ENTER 键， 则 InputBox 函数返回文本框中的文本。如果用户点击 Cancel 按钮，函数返回一个空字符串("")。

InputBox(prompt[,title][,default][,xpos][,ypos][,helpfile,context])



| 参数     | 描述                                                         |
| -------- | ------------------------------------------------------------ |
| prompt   | 必需。显示在对话框中的消息。prompt 的最大长度大约是 1024 个字符，这取决于所使用的字符的宽度。如果 prompt 中包含多个行，则可在各行之间用回车符（Chr(13)）、换行符（Chr(10)）或回车换行符的组合（Chr(13) & Chr(10)）来分隔各行。 |
| title    | 可选。对话框的标题。默认是应用程序的名称。                   |
| default  | 可选。一个在文本框中的默认文本。                             |
| xpos     | 可选。数值表达式，用于指定对话框的左边缘与屏幕左边缘的水平距离（单位为 twips*）。如果省略 xpos，则对话框会在水平方向居中。 |
| ypos     | 可选。数值表达式，用于指定对话框的上边缘与屏幕上边缘的垂直距离（单位为 twips*）。如果省略 ypos，则对话框显示在屏幕垂直方向距下边缘大约三分之一处。 |
| helpfile | 可选。字符串表达式，用于标识为对话框提供上下文相关帮助的帮助文件。必须与 context 参数一起使用。 |
| context  | 可选。数值表达式，用于标识由帮助文件的作者指定给某个帮助主题的上下文编号。必须与 helpfile 参数一起使用。 |

说明 如果同时提供了 helpfile 和 context，就会在对话框中自动添加“帮助”按钮。 如果用户单击确定或按下 ENTER，则 InputBox 函数返回文本框中的内容。如果用户单击取消，则函数返回一个零长度字符串 ("")。 下面例子利用 InputBox 函数显示一输入框并且把字符串赋值给输入变量：

```vb
Dim Input
Input = InputBox("输入名字") 
MsgBox ("输入：" & Input)
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#isempty)IsEmpty

返回一个布尔值，指示指定的变量是否已经初始化。如果变量未初始化则返回 True，否则返回 False。

IsEmpty(expression)



| 参数       | 描述                                   |
| ---------- | -------------------------------------- |
| expression | 必需。一个表达式（通常是一个变量名）。 |

说明 如果变量未初始化或显式地设置为 Empty，则函数IsEmpty返回True；否则函数返回False。如果expression包含一个以上的变量，总返回 False。

下面的示例利用 IsEmpty 函数决定变量是否能被初始化：

```vb
Dim MyVar, MyCheck
MyCheck = IsEmpty(MyVar)      ' 返回 True。
MyVar = Null                  ' 赋为 Null。
MyCheck = IsEmpty(MyVar)      ' 返回 False。
MyVar = Empty                 ' 赋为 Empty。
MyCheck = IsEmpty(MyVar)      ' 返回 True。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#isnull)IsNull

返回一个布尔值，指示指定的表达式是否包含无效数据（Null）。如果表达式是 Null 则返回 True，否则返回 False。

IsNull(expression)



| 参数       | 描述               |
| ---------- | ------------------ |
| expression | 必需。一个表达式。 |

说明

如果 expression 为 Null，则 IsNull 返回 True，即表达式不包含有效数据，否则 IsNull 返回 False。如果 expression 由多个变量组成，则表达式的任何组成变量中的 Null 都会使整个表达式返回 True。

Null 值指出变量不包含有效数据。Null 与 Empty 不同，后者指出变量未经初始化。Null 与零长度字符串 ("") 也不同，零长度字符串往往指的是空串。

重点 使用 IsNull 函数可以判断表达式是否包含 Null 值。在某些情况下想使表达式取值为 True，例如 IfVar=Null 和 IfVar<>Null，但它们通常总是为 False。这是因为任何包含 Null 的表达式本身就为 Null，所以表达式的结果为 False。

下面的示例利用 IsNull 函数决定变量是否包含 Null ：

```vb
Dim MyVar, MyCheck
MyCheck = IsNull(MyVar)      ' 返回 False。
MyVar = Null                 ' 赋为 Null。
MyCheck = IsNull(MyVar)      ' 返回 True。
MyVar = Empty                ' 赋为 Empty。
MyCheck = IsNull(MyVar)      ' 返回 False。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#isnumeric)IsNumeric

返回一个布尔值，指示指定的表达式是否可作为数字来计算。如果表达式作为数字来计算则返回 True ，否则返回 False。

IsNumeric(expression)



| 参数       | 描述               |
| ---------- | ------------------ |
| expression | 必需。一个表达式。 |

下面的示例利用 IsNumeric 函数决定变量是否可以作为数值：

```vb
Dim MyVar, MyCheck
MyVar = 53                    '赋值。
MyCheck = IsNumeric(MyVar)    ' 返回 True。
MyVar = "459.95"              ' 赋值。
MyCheck = IsNumeric(MyVar)    ' 返回True。
MyVar = "45 Help"             ' 赋值。
MyCheck = IsNumeric(MyVar)    ' 返回 False。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#isobject)IsObject

返回一个布尔值，指示指定的表达式是否是 automation 对象。如果表达式是 automation 对象则返回 True，否则返回 False。

IsObject(expression)



| 参数       | 描述               |
| ---------- | ------------------ |
| expression | 必需。一个表达式。 |

下面的示例利用 IsObject 函数决定标识符是否代表对象变量：

```vb
Dim MyInt, MyCheck, MyObject
Set MyObject = Me           
MyCheck = IsObject(MyObject)  ' 返回 True。
MyCheck = IsObject(MyInt)     ' 返回 False。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#loadpicture)LoadPicture

返回一个图片对象。

LoadPicture(picturename)



| 参数        | 描述                               |
| ----------- | ---------------------------------- |
| picturename | 必需。需被载入的图片文件的文件名。 |

说明

可以由 LoadPicture 识别的图形格式有位图文件 (.bmp)、图标文件 (.ico)、行程编码文件 (.rle)、图元文件 (.wmf)、增强型图元文件 (.emf)、GIF (.gif) 文件和 JPEG (.jpg) 文件。

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#msgbox)MsgBox

在对话框中显示消息，等待用户单击按钮，并返回一个值指示用户单击的按钮。

IMsgBox(prompt[, buttons][, title][, helpfile, context])



| 参数     | 描述                                                         |
| -------- | ------------------------------------------------------------ |
| prompt   | 必需。作为消息显示在对话框中的字符串表达式。prompt 的最大长度大约是 1024 个字符，这取决于所使用的字符的宽度。如果 prompt 中包含多个行，则可在各行之间用回车符（Chr(13)）、换行符（Chr(10)）或回车换行符的组合（Chr(13) & Chr(10)）分隔各行。 |
| buttons  | 可选，是表示指定显示按钮的数目和类型、使用的图标样式，默认按钮的标识以及消息框样式的数值的总和。默认值为 0。 0 = vbOKOnly - 只显示 OK 按钮 1 = vbOKCancel - 显示 OK 和 Cancel 按钮 2 = vbAbortRetryIgnore - 显示 Abort、Retry 和 Ignore 按钮 3 = vbYesNoCancel - 显示 Yes、No 和 Cancel 按钮 4 = vbYesNo - 显示 Yes 和 No 按钮 5 = vbRetryCancel - 显示 Retry 和 Cancel 按钮 16 = vbCritical - 显示临界信息图标 32 = vbQuestion - 显示警告查询图标 48 = vbExclamation - 显示警告消息图标 64 = vbInformation - 显示信息消息图标 0 = vbDefaultButton1 - 第一个按钮为默认按钮 256 = vbDefaultButton2 - 第二个按钮为默认按钮 512 = vbDefaultButton3 - 第三个按钮为默认按钮 768 = vbDefaultButton4 - 第四个按钮为默认按钮 0 = vbApplicationModal - 应用程序模式（用户必须响应消息框才能继续在当前应用程序中工作） 4096 = vbSystemModal - 系统模式（在用户响应消息框前，所有应用程序都被挂起） -------------------------- 我们可以把按钮分成四组：第一组值(0-5)用于描述对话框中显示的按钮类型与数目；第二组值(16,32,48,64)用于描述图标的样式；第三组值(0,256,512,768)用于确定默认按钮；而第四组值(0,4096)则决定消息框的样式。在将这些数字相加以生成 buttons 参数值时，只能从每组值中取用一个数字。 |
| title    | 可选。消息框的标题。默认是应用程序的名称。                   |
| helpfile | 可选。字符串表达式，用于标识为对话框提供上下文相关帮助的帮助文件。必须与 context 参数一起使用。 |
| context  | 可选。数值表达式，用于标识由帮助文件的作者指定给某个帮助主题的上下文编号。必须与 helpfile 参数一起使用。 |

说明 如果同时提供了 helpfile 和 context，则用户可以按 F1 键以查看与上下文相对应的帮助主题。

如果对话框显示取消按钮，则按 ESC 键与单击取消的效果相同。如果对话框包含帮助按钮，则有为对话框提供的上下文相关帮助。但是在单击其他按钮之前，不会返回任何值。

当MicroSoft Internet Explorer使用MsgBox函数时，任何对话框的标题总是包含"VBScript",以便于将其与标准对话框区别开来。

下面的例子演示了 MsgBox 函数的用法：

```vb
Dim MyVar
MyVar = MsgBox ("Hello World!", 65, "MsgBox Example")
   ' MyVar contains either 1 or 2, depending on which button is clicked.
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#rgb)RGB

返回表示 RGB 颜色值的数字。

RGB(red, green, blue)



| 参数  | 描述                                                         |
| ----- | ------------------------------------------------------------ |
| red   | 必需。介于 0 - 255 之间（且包括）的数字，代表颜色的红色部分。 |
| green | 必需。介于 0 - 255 之间（且包括）的数字，代表颜色的绿色部分。 |
| blue  | 必需。介于 0 - 255 之间（且包括）的数字，代表颜色的蓝色部分。 |

说明 接受颜色说明的应用程序方法和属性，要求该说明以整数代表 RGB 颜色值。RGB 颜色值指定了红色、绿色、蓝色的相对强度，三色组合形成显示的特定颜色。

低字节值表示红色，中字节值表示绿色，高字节值表示蓝色。

对于要求反转字节顺序的应用程序，下面函数在反转字节顺序下提供相同信息：

```vb
Function RevRGB(red, green, blue)
    RevRGB= CLng(blue + (green * 256) + (red * 65536))
End Function
RGB 函数中任一超过 255 的参数都假定为 255。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#scriptengine)ScriptEngine

返回当前使用的脚本语言。

ScriptEngine



返回值 ScriptEngine 函数可返回下列字符串：

| 字符串   | 描述                                                         |
| -------- | ------------------------------------------------------------ |
| VBScript | 表明当前使用的编写脚本引擎是 Microsoft Visual Basic Scripting Edition。 |
| JScript  | 表明当前使用的编写脚本引擎是 Microsoft JScript               |
| VBA      | 表明当前使用的编写脚本引擎是 Microsoft Visual Basic for Applications。 |

**示例**

下面的示例利用 ScriptEngine 函数返回描述所用书写语言的字符串：

```vb
Function GetScriptEngineInfo
  Dim s
  s = ""   '用必要的信息形成字符串。
  s = ScriptEngine & " Version "
  s = s & ScriptEngineMajorVersion & "."
  s = s & ScriptEngineMinorVersion & "."
  s = s & ScriptEngineBuildVersion 
  GetScriptEngineInfo =  s  '返回结果。
End Function
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#scriptenginebuildversion)ScriptEngineBuildVersion

返回使用的编写脚本引擎的编译版本号。

ScriptEngineBuildVersion



说明 返回值直接对应于所使用的 Scripting 程序语言的 DLL 文件中包含的版本信息。 下面的示例利用 ScriptEngineBuildVersion 函数返回创建的编写脚本引擎版本号：

```vb
 Function GetScriptEngineInfo
 Dim s
 s = ""   '用必要的信息形成字符串。
 s = ScriptEngine & " Version "
 s = s & ScriptEngineMajorVersion & "."
 s = s & ScriptEngineMinorVersion & "."
 s = s & ScriptEngineBuildVersion
 GetScriptEngineInfo = s  '返回结果。
End Function
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#scriptenginemajorversion)ScriptEngineMajorVersion

返回使用的编写脚本引擎的主版本号。

ScriptEngineMajorVersion



说明 返回值直接对应于所使用的脚本程序语言中 DLL 文件包含的版本信息。

下面的示例利用 ScriptEngineMajorVersion 函数返回编写脚本引擎的版本号：

```vb
 Function GetScriptEngineInfo
  Dim s
  s = ""            '用必要的信息形成字符串。
  s = ScriptEngine & " Version "
  s = s & ScriptEngineMajorVersion & "."
  s = s & ScriptEngineMinorVersion & "."
  s = s & ScriptEngineBuildVersion 
  GetScriptEngineInfo = s        '返回结果。
End Function
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#scriptengineminorversion)ScriptEngineMinorVersion

返回使用的编写引擎引擎的次版本号。

ScriptEngineMinorVersion



说明 返回值直接对应于所使用的脚本程序语言中 DLL 文件包含的版本信息。

下面的示例利用 ScriptEngineMinorVersion 函数返回编写引擎的副版本号：

示例



```vb
Function GetScriptEngineInfo
  Dim s
  s = ""   '用必要的信息形成字符串。
  s = ScriptEngine & " Version "
  s = s & ScriptEngineMajorVersion & "."
  s = s & ScriptEngineMinorVersion & "."
  s = s & ScriptEngineBuildVersion 
  GetScriptEngineInfo = s  '返回结果。
End Function
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#setlocale)SetLocale

设置 locale ID

SetLocale(lcid)



lcid 参数可以是任意一个合法的 32 位数值或短字符串，该值必须唯一标识一个地理区域。能被识别的值可以查阅 区域设置 ID 表。

说明

若 lcid 为零，区域被设置为与当前系统设置匹配。 一个 locale 是用户参考信息集合，与用户的语言、国家和文化传统有关。该 locale 决定键盘布局、字母排序顺序和日期、时间、数字与货币格式。

```vb
Dim currentLocale
' Get the current locale
currentLocale = GetLocale

Sub Button1_onclick
  Dim original
  original = SetLocale("en-gb")
  mydate = CDate(UKDate.value)
  ' IE always sets the locale to US English so use the
  ' currentLocale variable to set the locale to US English
  original = SetLocale(currentLocale)
  USDate.value = FormatDateTime(mydate,vbShortDate)
End Sub

Sub button2_onclick
  Dim original
  original = SetLocale("de")
  myvalue = CCur(GermanNumber.value)
  original = SetLocale("en-gb")
  USNumber.value = FormatCurrency(myvalue)
End Sub
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#typename)TypeName

返回指定变量的子类型

TypeName(varname)



| 参数    | 描述               |
| ------- | ------------------ |
| varname | 必需。变量的名称。 |

返回值 TypeName 函数返回值如下：

| 值             | 描述                       |
| -------------- | -------------------------- |
| Byte           | 字节值                     |
| Integer        | 整型值                     |
| Long           | 长整型值                   |
| Single         | 单精度浮点值               |
| Double         | 双精度浮点值               |
| Currency       | 货币值                     |
| Decimal        | 十进制值                   |
| Date           | 日期或时间值               |
| String         | 字符串值                   |
| Boolean        | Boolean                    |
| Empty          | 未初始化                   |
| Null           | 无有效数据                 |
| 实际对象类型名 |                            |
| Object         | 一般对象                   |
| Unknown        | 未知对象类型               |
| Nothing        | 还未引用对象实例的对象变量 |
| Error          | 错误                       |

说明 下面的示例利用 TypeName 函数返回变量信息：

```vb
Dim ArrayVar(4), MyType
NullVar = Null    ' 赋 Null 值。
MyType = TypeName("VBScript")   ' 返回 "String"。
MyType = TypeName(4)            ' 返回 "Integer"。
MyType = TypeName(37.50)        ' 返回 "Double"。
MyType = TypeName(NullVar)      ' 返回 "Null"。
MyType = TypeName(ArrayVar)     ' 返回 "Variant()"。
```

### [#](https://www.weistock.com/docs/VBA/VBScript/函数.html#vartype)VarType

返回指示变量子类型的值。

VarType(varname)



| 参数    | 描述               |
| ------- | ------------------ |
| varname | 必需。变量的名称。 |

返回值

VarType 函数返回下列值：

| 常数         | 值   | 描述                            |
| ------------ | ---- | ------------------------------- |
| vbEmpty      | 0    | Empty（未初始化）               |
| vbNull       | 1    | Null（无有效数据）              |
| vbInteger    | 2    | 整数                            |
| vbLong       | 3    | 长整数                          |
| vbSingle     | 4    | 单精度浮点数                    |
| vbDouble     | 5    | 双精度浮点数                    |
| vbCurrency   | 6    | 货币                            |
| vbDate       | 7    | 日期                            |
| vbString     | 8    | 字符串                          |
| vbObject     | 9    | Automation                      |
| vbError      | 10   | 错误                            |
| vbBoolean    | 11   | Boolean                         |
| vbVariant    | 12   | Variant（只和变量数组一起使用） |
| vbDataObject | 13   | 数据访问对象                    |
| vbByte       | 17   | 字节                            |
| vbArray      | 8192 | 数组                            |

注意 这些常数是由 VBScript 指定的。所以，这些名称可在代码中随处使用，以代替实际值。

说明 VarType 函数从不通过自己返回 Array 的值。它总是要添加一些其他值来指示一个具体类型的数组。当 Variant 的值被添加到 Array 的值中以表明 VarType 函数的参数是一个数组时，它才被返回。例如，对一个整数数组的返回值是 2 + 8192 的计算结果，或 8194。如果一个对象有默认，则 VarType(object) 返回对象默认属性的类型。

下面函数利用 VarType 函数决定变量的子类型.

```vb
Dim MyCheck
MyCheck = VarType(300)           ' 返回 2。
MyCheck = VarType(#10/19/62#)    ' 返回 7。
MyCheck = VarType("VBScript")    ' 返回 8。
```