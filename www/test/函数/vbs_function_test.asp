<%@ Language="VBScript" CodePage="65001" %>
<%
Option Explicit
Session.CodePage = 65001
Response.CodePage = 65001
Response.Charset = "UTF-8"

' 字符串测试常量 (使用 ASCII 避免编码问题)
Dim STR_TEST_DEMO, STR_TEST_LONG, STR_TEST_SENTENCE
STR_TEST_DEMO = "Hello World! ABC! DEF!"
STR_TEST_LONG = "The quick brown fox jumps over the lazy dog"
STR_TEST_SENTENCE = "VBScript-is-awesome-for-web-development"
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>VBScript 函数测试</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Arial, sans-serif; background: #f5f5f5; padding: 20px; }
        .container { max-width: 1200px; margin: 0 auto; background: white; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); overflow: hidden; }
        .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; text-align: center; }
        .header h1 { font-size: 2em; margin-bottom: 10px; }
        .header p { opacity: 0.9; }
        .nav { display: flex; background: #f8f9fa; border-bottom: 2px solid #dee2e6; overflow-x: auto; }
        .nav a { padding: 15px 25px; text-decoration: none; color: #495057; border-right: 1px solid #dee2e6; transition: all 0.3s; white-space: nowrap; cursor: pointer; }
        .nav a:hover, .nav a.active { background: #667eea; color: white; }
        .content { padding: 30px; }
        .section { display: none; }
        .section.active { display: block; }
        .test-group { background: #f8f9fa; border-left: 4px solid #667eea; padding: 20px; margin-bottom: 20px; border-radius: 4px; }
        .test-group h3 { color: #667eea; margin-bottom: 15px; font-size: 1.3em; }
        .test-item { background: white; padding: 15px; margin-bottom: 10px; border-radius: 4px; border: 1px solid #e9ecef; }
        .test-item h4 { color: #495057; margin-bottom: 8px; font-size: 1em; }
        .code { background: #2d3748; color: #a0aec0; padding: 10px; border-radius: 4px; font-family: 'Consolas', 'Monaco', monospace; font-size: 0.9em; margin: 8px 0; overflow-x: auto; }
        .result { background: #c6f6d5; color: #22543d; padding: 10px; border-radius: 4px; margin-top: 8px; font-family: 'Consolas', 'Monaco', monospace; }
        .result-label { font-weight: bold; color: #2f855a; margin-right: 8px; }
        table { width: 100%; border-collapse: collapse; margin: 10px 0; }
        th, td { padding: 12px; text-align: left; border: 1px solid #dee2e6; }
        th { background: #667eea; color: white; font-weight: 600; }
        tr:nth-child(even) { background: #f8f9fa; }
        .info-box { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin: 15px 0; border-radius: 4px; }
        .warning-box { background: #fff3e0; border-left: 4px solid #ff9800; padding: 15px; margin: 15px 0; border-radius: 4px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🧪 VBScript 函数测试实验室</h1>
            <p>实时测试各类 VBScript 内置函数的返回值和用法</p>
        </div>

        <div class="nav">
            <a onclick="showSection('math')" class="active">📊 数学函数</a>
            <a onclick="showSection('string')">📝 字符串函数</a>
            <a onclick="showSection('convert')">🔄 转换函数</a>
            <a onclick="showSection('date')">📅 日期时间</a>
            <a onclick="showSection('array')">📦 数组函数</a>
            <a onclick="showSection('judge')">✔️ 判断函数</a>
            <a onclick="showSection('other')">🔧 其他函数</a>
        </div>

        <div class="content">

            <!-- ==================== 数学函数 ==================== -->
            <div id="math" class="section active">
                <div class="test-group">
                    <h3>📐 基本数学运算</h3>

                    <div class="test-item">
                        <h4>Abs() - 绝对值</h4>
                        <div class="code">Abs(-123.45) = <%= Abs(-123.45) %></div>
                        <div class="code">Abs(123.45) = <%= Abs(123.45) %></div>
                    </div>

                    <div class="test-item">
                        <h4>Fix() / Int() - 取整函数</h4>
                        <div class="code">Fix(12.7) = <%= Fix(12.7) %> | Fix(-12.7) = <%= Fix(-12.7) %></div>
                        <div class="code">Int(12.7) = <%= Int(12.7) %> | Int(-12.7) = <%= Int(-12.7) %></div>
                        <div class="result"><span class="result-label">说明：</span>Fix 向0取整，Int 向下取整</div>
                    </div>

                    <div class="test-item">
                        <h4>Round() - 四舍五入</h4>
                        <div class="code">Round(123.4567, 2) = <%= Round(123.4567, 2) %></div>
                        <div class="code">Round(123.4567, 0) = <%= Round(123.4567, 0) %></div>
                    </div>

                    <div class="test-item">
                        <h4>Sgn() - 符号判断</h4>
                        <div class="code">Sgn(-100) = <%= Sgn(-100) %> | Sgn(0) = <%= Sgn(0) %> | Sgn(100) = <%= Sgn(100) %></div>
                    </div>
                </div>

                <div class="test-group">
                    <h3>📈 三角函数与指数</h3>

                    <div class="test-item">
                        <h4>三角函数（弧度制）</h4>
                        <div class="code">Sin(3.1415926/2) = <%= Sin(3.1415926/2) %> (sin 90°)</div>
                        <div class="code">Cos(0) = <%= Cos(0) %></div>
                        <div class="code">Tan(3.1415926/4) = <%= Tan(3.1415926/4) %> (tan 45°)</div>
                        <div class="code">Atn(1) * 4 = <%= Atn(1) * 4 %> (π的近似值)</div>
                    </div>

                    <div class="test-item">
                        <h4>Exp() / Log() / Sqr() - 指数、对数、平方根</h4>
                        <div class="code">Exp(1) = <%= Exp(1) %> (e的值)</div>
                        <div class="code">Exp(2) = <%= Exp(2) %> (e²)</div>
                        <div class="code">Log(100) / Log(10) = <%= Log(100) / Log(10) %> (lg100)</div>
                        <div class="code">Sqr(144) = <%= Sqr(144) %></div>
                    </div>
                </div>

                <div class="test-group">
                    <h3>🎲 随机数生成</h3>
                    <div class="test-item">
                        <h4>Rnd() - 随机数</h4>
                        <div class="code">Rnd() = <%= Rnd() %></div>
                        <div class="code">生成 1-100 之间的随机整数：<%= Int((100 - 1 + 1) * Rnd + 1) %></div>
                        <div class="code">再生成一次：<%= Int((100 - 1 + 1) * Rnd + 1) %></div>
                        <div class="result"><span class="result-label">提示：</span>刷新页面会得到不同的随机数</div>
                    </div>
                </div>
            </div>

            <!-- ==================== 字符串函数 ==================== -->
            <div id="string" class="section">
                <div class="test-group">
                    <h3>✂️ 字符串截取</h3>

                    <div class="test-item">
                        <h4>Left() / Right() / Mid()</h4>
                        <%
                            ' 使用 ASCII 字符测试中文字符处理
                            Dim strDemo, strCN
                            strDemo = "Hello World!"
                            ' 中文字符串通过 Request.Form 或其他方式获取，这里用 ASCII 演示
                            strCN = ChrW(-24029) & ChrW(-22609) & ChrW(-19990) & ChrW(-30028) ' "你好世界" 的 UTF-16 编码
                        %>
                        <div class="code">strDemo = "<%= strDemo %>"</div>
                        <div class="code">Left(strDemo, 5) = "<%= Left(strDemo, 5) %>"</div>
                        <div class="code">Right(strDemo, 6) = "<%= Right(strDemo, 6) %>"</div>
                        <div class="code">Mid(strDemo, 7, 5) = "<%= Mid(strDemo, 7, 5) %>"</div>
                    </div>

                    <div class="test-item">
                        <h4>Len() - 字符串长度</h4>
                        <div class="code">Len("VBScript") = <%= Len("VBScript") %></div>
                        <div class="code">Len("HelloWorld") = <%= Len("HelloWorld") %></div>
                        <div class="result"><span class="result-label">说明：</span>中文字符在 VBScript 中作为单字符处理</div>
                    </div>
                </div>

                <div class="test-group">
                    <h3>🔍 字符串查找与比较</h3>

                    <div class="test-item">
                        <h4>InStr() - 查找子字符串位置</h4>
                        <div class="code">InStr("Hello World", "World") = <%= InStr("Hello World", "World") %></div>
                        <div class="code">InStr("Hello World", "abc") = <%= InStr("Hello World", "abc") %> (未找到返回0)</div>
                        <div class="code">InStr(5, "VBScript VBScript", "VB") = <%= InStr(5, "VBScript VBScript", "VB") %> (从位置5开始)</div>
                    </div>

                    <div class="test-item">
                        <h4>StrComp() - 字符串比较</h4>
                        <div class="code">StrComp("Apple", "Apple") = <%= StrComp("Apple", "Apple") %> (相等)</div>
                        <div class="code">StrComp("A", "B") = <%= StrComp("A", "B") %> (小于)</div>
                        <div class="code">StrComp("B", "A") = <%= StrComp("B", "A") %> (大于)</div>
                    </div>
                </div>

                <div class="test-group">
                    <h3>🔄 字符串转换</h3>

                    <div class="test-item">
                        <h4>LCase() / UCase() - 大小写转换</h4>
                        <div class="code">LCase("Hello World") = "<%= LCase("Hello World") %>"</div>
                        <div class="code">UCase("Hello World") = "<%= UCase("Hello World") %>"</div>
                    </div>

                    <div class="test-item">
                        <h4>LTrim() / RTrim() / Trim() - 去除空格</h4>
                        <div class="code">Trim("  Hello  ") = "<%= Trim("  Hello  ") %>"</div>
                        <div class="code">LTrim("  Hello  ") = "<%= LTrim("  Hello  ") %>"</div>
                        <div class="code">RTrim("  Hello  ") = "<%= RTrim("  Hello  ") %>"</div>
                    </div>

                    <div class="test-item">
                        <h4>StrReverse() - 字符串反转</h4>
                        <div class="code">StrReverse("Hello") = "<%= StrReverse("Hello") %>"</div>
                        <div class="code">StrReverse("你好世界") = "<%= StrReverse("你好世界") %>"</div>
                    </div>

                    <div class="test-item">
                        <h4>Space() / String() - 重复字符</h4>
                        <div class="code">Space(5) & "Hello" = "<%= Space(5) & "Hello" %>" (前导空格)</div>
                        <div class="code">String(5, "*") = "<%= String(5, "*") %>"</div>
                        <div class="code">String(3, "AB") = "<%= String(3, "AB") %>" (使用首字符)</div>
                    </div>
                </div>

                <div class="test-group">
                    <h3>🔧 高级字符串操作</h3>

                    <div class="test-item">
                        <h4>Split() - 分割字符串为数组</h4>
                        <%
                            Dim arrSplit, i, strSplitResult
                            arrSplit = Split("苹果,香蕉,橙子,葡萄", ",")
                            strSplitResult = ""
                            For i = 0 To UBound(arrSplit)
                                strSplitResult = strSplitResult & "arrSplit(" & i & ") = " & arrSplit(i) & "<br>"
                            Next
                        %>
                        <div class="code">Split("苹果,香蕉,橙子,葡萄", ",")</div>
                        <div class="result"><%= strSplitResult %></div>
                    </div>

                    <div class="test-item">
                        <h4>Replace() - 字符串替换</h4>
                        <div class="code">Replace("Hello World", "World", "VBScript") = "<%= Replace("Hello World", "World", "VBScript") %>"</div>
                        <div class="code">Replace("aaa", "a", "b") = "<%= Replace("aaa", "a", "b") %>" (全部替换)</div>
                    </div>

                    <div class="test-item">
                        <h4>Join() - 数组合并为字符串</h4>
                        <%
                            Dim arrJoin
                            arrJoin = Array("A", "B", "C", "D")
                        %>
                        <div class="code">Join(Array("A","B","C","D"), "-") = "<%= Join(arrJoin, "-") %>"</div>
                        <div class="code">Join(Array("A","B","C","D")) = "<%= Join(arrJoin) %>" (默认无分隔符)</div>
                    </div>
                </div>

                <div class="test-group">
                    <h3>🔢 ASCII 码转换</h3>

                    <div class="test-item">
                        <h4>Asc() / Chr() - 字符与 ASCII 码互转</h4>
                        <div class="code">Asc("A") = <%= Asc("A") %></div>
                        <div class="code">Asc("中") = <%= Asc("中") %></div>
                        <div class="code">Chr(65) = "<%= Chr(65) %>"</div>
                        <div class="code">Chr(20013) = "<%= Chr(20013) %>" (中文"中")</div>
                    </div>
                </div>
            </div>

            <!-- ==================== 转换函数 ==================== -->
            <div id="convert" class="section">
                <div class="test-group">
                    <h3>🔄 类型转换函数</h3>

                    <div class="test-item">
                        <h4>CBool() - 转换为布尔值</h4>
                        <div class="code">CBool(0) = <%= CBool(0) %> (False)</div>
                        <div class="code">CBool(1) = <%= CBool(1) %> (True)</div>
                        <div class="code">CBool(-5) = <%= CBool(-5) %> (True)</div>
                    </div>

                    <div class="test-item">
                        <h4>CInt() / CLng() - 转换为整数</h4>
                        <div class="code">CInt(123.7) = <%= CInt(123.7) %> (四舍五入)</div>
                        <div class="code">CInt(123.2) = <%= CInt(123.2) %></div>
                        <div class="code">CLng(123456789.7) = <%= CLng(123456789.7) %></div>
                    </div>

                    <div class="test-item">
                        <h4>CSng() / CDbl() - 转换为浮点数</h4>
                        <div class="code">CSng("3.14") = <%= CSng("3.14") %></div>
                        <div class="code">CDbl("3.1415926535") = <%= CDbl("3.1415926535") %></div>
                    </div>

                    <div class="test-item">
                        <h4>CStr() - 转换为字符串</h4>
                        <div class="code">CStr(123) = "<%= CStr(123) %>"</div>
                        <div class="code">CStr(3.14) = "<%= CStr(3.14) %>"</div>
                        <div class="code">CStr(True) = "<%= CStr(True) %>"</div>
                    </div>

                    <div class="test-item">
                        <h4>CByte() / CCur() / CDate()</h4>
                        <div class="code">CByte(255) = <%= CByte(255) %> (字节 0-255)</div>
                        <div class="code">CCur(1234.5678) = <%= CCur(1234.5678) %> (货币类型，4位小数)</div>
                        <div class="code">CDate("2024-03-07") = <%= CDate("2024-03-07") %></div>
                    </div>
                </div>

                <div class="test-group">
                    <h3>🔢 进制转换</h3>

                    <div class="test-item">
                        <h4>Hex() - 十六进制</h4>
                        <div class="code">Hex(255) = "<%= Hex(255) %>"</div>
                        <div class="code">Hex(16) = "<%= Hex(16) %>"</div>
                    </div>

                    <div class="test-item">
                        <h4>Oct() - 八进制</h4>
                        <div class="code">Oct(8) = "<%= Oct(8) %>"</div>
                        <div class="code">Oct(64) = "<%= Oct(64) %>"</div>
                    </div>
                </div>

                <div class="test-group">
                    <h3>📋 VarType() - 变量类型检测</h3>
                    <div class="test-item">
                        <%
                            Dim vtEmpty, vtNull, vtInt, vtStr, vtDate, vtArr
                            vtInt = 123
                            vtStr = "Hello"
                            vtDate = Now
                            vtArr = Array(1, 2, 3)
                        %>
                        <table>
                            <tr>
                                <th>变量</th>
                                <th>值</th>
                                <th>VarType()</th>
                                <th>类型</th>
                            </tr>
                            <tr>
                                <td>vtEmpty</td>
                                <td><%= "未初始化" %></td>
                                <td><%= VarType(vtEmpty) %></td>
                                <td>vbEmpty (0)</td>
                            </tr>
                            <tr>
                                <td>vtInt</td>
                                <td><%= vtInt %></td>
                                <td><%= VarType(vtInt) %></td>
                                <td>vbInteger (2)</td>
                            </tr>
                            <tr>
                                <td>vtStr</td>
                                <td>"<%= vtStr %>"</td>
                                <td><%= VarType(vtStr) %></td>
                                <td>vbString (8)</td>
                            </tr>
                            <tr>
                                <td>vtDate</td>
                                <td><%= vtDate %></td>
                                <td><%= VarType(vtDate) %></td>
                                <td>vbDate (7)</td>
                            </tr>
                            <tr>
                                <td>vtArr</td>
                                <td>Array(...)</td>
                                <td><%= VarType(vtArr) %></td>
                                <td>vbArray (8192)</td>
                            </tr>
                        </table>
                    </div>
                </div>
            </div>

            <!-- ==================== 日期时间函数 ==================== -->
            <div id="date" class="section">
                <div class="test-group">
                    <h3>⏰ 获取当前时间</h3>

                    <div class="test-item">
                        <h4>Now / Date / Time</h4>
                        <div class="code">Now = <%= Now %></div>
                        <div class="code">Date = <%= Date %></div>
                        <div class="code">Time = <%= Time %></div>
                    </div>

                    <div class="test-item">
                        <h4>Timer - 从午夜开始的秒数</h4>
                        <div class="code">Timer = <%= Timer %></div>
                        <div class="result"><span class="result-label">用途：</span>用于计算代码执行时间</div>
                    </div>
                </div>

                <div class="test-group">
                    <h3>📅 日期时间组成部分</h3>

                    <div class="test-item">
                        <h4>Year / Month / Day</h4>
                        <div class="code">Year(Date) = <%= Year(Date) %></div>
                        <div class="code">Month(Date) = <%= Month(Date) %></div>
                        <div class="code">Day(Date) = <%= Day(Date) %></div>
                    </div>

                    <div class="test-item">
                        <h4>Hour / Minute / Second</h4>
                        <div class="code">Hour(Now) = <%= Hour(Now) %></div>
                        <div class="code">Minute(Now) = <%= Minute(Now) %></div>
                        <div class="code">Second(Now) = <%= Second(Now) %></div>
                    </div>

                    <div class="test-item">
                        <h4>Weekday / WeekdayName</h4>
                        <div class="code">Weekday(Now) = <%= Weekday(Now) %> (1=周日, 7=周六)</div>
                        <div class="code">WeekdayName(Weekday(Now)) = <%= WeekdayName(Weekday(Now)) %></div>
                        <div class="code">WeekdayName(Weekday(Now), True) = <%= WeekdayName(Weekday(Now), True) %> (缩写)</div>
                    </div>
                </div>

                <div class="test-group">
                    <h3>📆 日期计算</h3>

                    <div class="test-item">
                        <h4>DateAdd() - 日期加减</h4>
                        <div class="code">DateAdd("d", 7, Date) = <%= DateAdd("d", 7, Date) %> (加7天)</div>
                        <div class="code">DateAdd("m", 2, Date) = <%= DateAdd("m", 2, Date) %> (加2月)</div>
                        <div class="code">DateAdd("yyyy", -1, Date) = <%= DateAdd("yyyy", -1, Date) %> (减1年)</div>
                        <div class="code">DateAdd("h", 3, Now) = <%= DateAdd("h", 3, Now) %> (加3小时)</div>
                    </div>

                    <div class="test-item">
                        <h4>DateDiff() - 日期差</h4>
                        <div class="code">DateDiff("d", #2024-01-01#, #2024-12-31#) = <%= DateDiff("d", #2024-01-01#, #2024-12-31#) %> 天</div>
                        <div class="code">DateDiff("m", #2024-01-01#, #2024-12-31#) = <%= DateDiff("m", #2024-01-01#, #2024-12-31#) %> 月</div>
                        <div class="code">DateDiff("yyyy", #2000-01-01#, Date) = <%= DateDiff("yyyy", #2000-01-01#, Date) %> 年</div>
                    </div>
                </div>

                <div class="test-group">
                    <div class="info-box">
                        <strong>📖 DateAdd / DateDiff 的 interval 参数说明：</strong><br>
                        yyyy=年, q=季度, m=月, y=一年的日数, d=日<br>
                        w=一周的日数, ww=周, h=小时, n=分钟, s=秒
                    </div>
                </div>
            </div>

            <!-- ==================== 数组函数 ==================== -->
            <div id="array" class="section">
                <div class="test-group">
                    <h3>📦 数组边界</h3>

                    <div class="test-item">
                        <h4>LBound() / UBound()</h4>
                        <%
                            Dim arrTest
                            arrTest = Array("苹果", "香蕉", "橙子", "葡萄", "芒果")
                        %>
                        <div class="code">arrTest = Array("苹果", "香蕉", "橙子", "葡萄", "芒果")</div>
                        <div class="code">LBound(arrTest) = <%= LBound(arrTest) %> (最小下标)</div>
                        <div class="code">UBound(arrTest) = <%= UBound(arrTest) %> (最大下标)</div>
                        <div class="code">数组长度 = UBound - LBound + 1 = <%= UBound(arrTest) - LBound(arrTest) + 1 %></div>
                    </div>

                    <div class="test-item">
                        <h4>二维数组示例</h4>
                        <%
                            Dim arr2D(1, 2)
                            arr2D(0,0) = "A1" : arr2D(0,1) = "A2" : arr2D(0,2) = "A3"
                            arr2D(1,0) = "B1" : arr2D(1,1) = "B2" : arr2D(1,2) = "B3"
                        %>
                        <div class="code">Dim arr2D(1, 2)</div>
                        <div class="code">第一维: LBound = <%= LBound(arr2D, 1) %>, UBound = <%= UBound(arr2D, 1) %></div>
                        <div class="code">第二维: LBound = <%= LBound(arr2D, 2) %>, UBound = <%= UBound(arr2D, 2) %></div>
                    </div>
                </div>

                <div class="test-group">
                    <h3>🔍 数组过滤</h3>

                    <div class="test-item">
                        <h4>Filter() - 筛选数组元素</h4>
                        <%
                            Dim arrSource, arrFiltered, j
                            arrSource = Array("apple.txt", "image.jpg", "data.txt", "photo.jpg")
                            arrFiltered = Filter(arrSource, ".txt", True)
                        %>
                        <div class="code">arrSource = Array("apple.txt", "image.jpg", "data.txt", "photo.jpg")</div>
                        <div class="code">Filter(arrSource, ".txt", True) - 筛选包含 ".txt" 的元素：</div>
                        <div class="result">
                            <%
                                For j = 0 To UBound(arrFiltered)
                                    Response.Write "arrFiltered(" & j & ") = " & arrFiltered(j) & "<br>"
                                Next
                            %>
                        </div>
                        <%
                            arrFiltered = Filter(arrSource, ".jpg", False)
                        %>
                        <div class="code">Filter(arrSource, ".jpg", False) - 排除包含 ".jpg" 的元素：</div>
                        <div class="result">
                            <%
                                For j = 0 To UBound(arrFiltered)
                                    Response.Write "arrFiltered(" & j & ") = " & arrFiltered(j) & "<br>"
                                Next
                            %>
                        </div>
                    </div>
                </div>

                <div class="test-group">
                    <h3>🔄 数组与字符串转换</h3>

                    <div class="test-item">
                        <h4>Split() / Join() 组合使用</h4>
                        <%
                            Dim arrWords, strJoined, strSentence
                            strSentence = "VBScript-is-awesome"
                            arrWords = Split(strSentence, "-")
                            strJoined = Join(arrWords, " ")
                        %>
                        <div class="code">原始字符串: "<%= strSentence %>"</div>
                        <div class="code">Split(strSentence, "-") = ["<%= Join(arrWords, """, """) %>"]</div>
                        <div class="code">Join(arrWords, " ") = "<%= strJoined %>"</div>
                    </div>
                </div>
            </div>

            <!-- ==================== 判断函数 ==================== -->
            <div id="judge" class="section">
                <div class="test-group">
                    <h3>✔️ 类型判断函数</h3>

                    <div class="test-item">
                        <h4>IsArray() - 判断是否为数组</h4>
                        <div class="code">IsArray(Array(1,2,3)) = <%= IsArray(Array(1,2,3)) %></div>
                        <div class="code">IsArray("Hello") = <%= IsArray("Hello") %></div>
                    </div>

                    <div class="test-item">
                        <h4>IsDate() - 判断是否为有效日期</h4>
                        <div class="code">IsDate("2024-03-07") = <%= IsDate("2024-03-07") %></div>
                        <div class="code">IsDate("2024/03/07") = <%= IsDate("2024/03/07") %></div>
                        <div class="code">IsDate("Hello") = <%= IsDate("Hello") %></div>
                    </div>

                    <div class="test-item">
                        <h4>IsEmpty() - 判断是否已初始化</h4>
                        <%
                            Dim emptyVar, initializedVar
                            initializedVar = "test"
                        %>
                        <div class="code">Dim emptyVar (未赋值): IsEmpty(emptyVar) = <%= IsEmpty(emptyVar) %></div>
                        <div class="code">initializedVar = "test": IsEmpty(initializedVar) = <%= IsEmpty(initializedVar) %></div>
                    </div>

                    <div class="test-item">
                        <h4>IsNull() - 判断是否为 Null</h4>
                        <div class="code">IsNull(Null) = <%= IsNull(Null) %></div>
                        <div class="code">IsNull("") = <%= IsNull("") %> (空字符串不是Null)</div>
                        <div class="code">IsNull(0) = <%= IsNull(0) %></div>
                    </div>

                    <div class="test-item">
                        <h4>IsNumeric() - 判断是否为数字</h4>
                        <div class="code">IsNumeric(123) = <%= IsNumeric(123) %></div>
                        <div class="code">IsNumeric("123") = <%= IsNumeric("123") %></div>
                        <div class="code">IsNumeric("12.3") = <%= IsNumeric("12.3") %></div>
                        <div class="code">IsNumeric("Hello") = <%= IsNumeric("Hello") %></div>
                    </div>

                    <div class="test-item">
                        <h4>IsObject() - 判断是否为对象</h4>
                        <%
                            Dim objDict
                            Set objDict = Server.CreateObject("Scripting.Dictionary")
                        %>
                        <div class="code">IsObject(Server.CreateObject("Scripting.Dictionary")) = <%= IsObject(objDict) %></div>
                        <div class="code">IsObject("Hello") = <%= IsObject("Hello") %></div>
                        <%
                            Set objDict = Nothing
                        %>
                    </div>
                </div>

                <div class="test-group">
                    <div class="warning-box">
                        <strong>⚠️ 注意：Empty 与 Null 的区别</strong><br>
                        • <strong>Empty</strong>：变量未初始化（仅适用于 Variant 变量）<br>
                        • <strong>Null</strong>：表示无效数据或无数据<br>
                        • <strong>""</strong>：空字符串，是有效的字符串值
                    </div>
                </div>
            </div>

            <!-- ==================== 其他函数 ==================== -->
            <div id="other" class="section">
                <div class="test-group">
                    <h3>🎨 颜色处理</h3>

                    <div class="test-item">
                        <h4>RGB() - 生成颜色值</h4>
                        <div class="code">RGB(255, 0, 0) = <%= RGB(255, 0, 0) %> (红色)</div>
                        <div class="code">RGB(0, 255, 0) = <%= RGB(0, 255, 0) %> (绿色)</div>
                        <div class="code">RGB(0, 0, 255) = <%= RGB(0, 0, 255) %> (蓝色)</div>
                        <div class="code">RGB(255, 255, 0) = <%= RGB(255, 255, 0) %> (黄色)</div>
                    </div>
                </div>

                <div class="test-group">
                    <h3>💰 格式化函数</h3>

                    <div class="test-item">
                        <h4>FormatCurrency()</h4>
                        <div class="code">FormatCurrency(1234.567) = <%= FormatCurrency(1234.567) %></div>
                        <div class="code">FormatCurrency(1234.567, 1) = <%= FormatCurrency(1234.567, 1) %> (1位小数)</div>
                    </div>

                    <div class="test-item">
                        <h4>FormatNumber()</h4>
                        <div class="code">FormatNumber(1234.5678) = <%= FormatNumber(1234.5678) %></div>
                        <div class="code">FormatNumber(1234.5678, 2) = <%= FormatNumber(1234.5678, 2) %></div>
                    </div>

                    <div class="test-item">
                        <h4>FormatPercent()</h4>
                        <div class="code">FormatPercent(0.1234) = <%= FormatPercent(0.1234) %></div>
                        <div class="code">FormatPercent(0.5) = <%= FormatPercent(0.5) %></div>
                    </div>

                    <div class="test-item">
                        <h4>FormatDateTime()</h4>
                        <div class="code">FormatDateTime(Now, vbGeneralDate) = <%= FormatDateTime(Now, vbGeneralDate) %></div>
                        <div class="code">FormatDateTime(Now, vbLongDate) = <%= FormatDateTime(Now, vbLongDate) %></div>
                        <div class="code">FormatDateTime(Now, vbShortDate) = <%= FormatDateTime(Now, vbShortDate) %></div>
                    </div>
                </div>

                <div class="test-group">
                    <h3>🖥️ 交互函数</h3>

                    <div class="test-item">
                        <h4>InputBox() - 输入框（仅客户端）</h4>
                        <div class="warning-box">
                            ⚠️ InputBox 和 MsgBox 是客户端函数，在服务器端 ASP 中无法使用。<br>
                            下面展示的是它们的使用方法和语法。
                        </div>
                        <div class="code">' 语法示例（客户端VBScript）:<br>
strName = InputBox("请输入您的姓名:", "输入框", "张三")<br>
MsgBox "你好, " & strName, vbOKOnly, "问候"</div>
                    </div>

                    <div class="test-item">
                        <h4>MsgBox() - 消息框返回值</h4>
                        <div class="code">vbOK = <%= vbOK %> (确定按钮)</div>
                        <div class="code">vbCancel = <%= vbCancel %> (取消按钮)</div>
                        <div class="code">vbYes = <%= vbYes %> (是按钮)</div>
                        <div class="code">vbNo = <%= vbNo %> (否按钮)</div>
                    </div>
                </div>

                <div class="test-group">
                    <h3>📌 FormatDateTime 格式常量</h3>
                    <table>
                        <tr>
                            <th>常量</th>
                            <th>值</th>
                            <th>说明</th>
                        </tr>
                        <tr>
                            <td>vbGeneralDate</td>
                            <td>0</td>
                            <td>显示日期和/或时间</td>
                        </tr>
                        <tr>
                            <td>vbLongDate</td>
                            <td>1</td>
                            <td>用区域设置的长日期格式显示</td>
                        </tr>
                        <tr>
                            <td>vbShortDate</td>
                            <td>2</td>
                            <td>用区域设置的短日期格式显示</td>
                        </tr>
                        <tr>
                            <td>vbLongTime</td>
                            <td>3</td>
                            <td>用区域设置的时间格式显示</td>
                        </tr>
                        <tr>
                            <td>vbShortTime</td>
                            <td>4</td>
                            <td>用24小时格式显示时间</td>
                        </tr>
                    </table>
                </div>
            </div>

        </div>
    </div>

    <script>
        function showSection(sectionId) {
            // 隐藏所有 section
            var sections = document.querySelectorAll('.section');
            for (var i = 0; i < sections.length; i++) {
                sections[i].classList.remove('active');
            }

            // 移除所有导航链接的 active 类
            var navLinks = document.querySelectorAll('.nav a');
            for (var i = 0; i < navLinks.length; i++) {
                navLinks[i].classList.remove('active');
            }

            // 显示选中的 section
            document.getElementById(sectionId).classList.add('active');

            // 给当前导航链接添加 active 类
            event.target.classList.add('active');
        }
    </script>
</body>
</html>
