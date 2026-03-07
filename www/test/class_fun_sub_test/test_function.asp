<%@ Language="VBScript" CodePage="65001" %>
<%
Option Explicit
Session.CodePage = 65001
Response.CodePage = 65001
Response.Charset = "UTF-8"

' 测试 Function 过程
' 测试无参数 Function

Function GetVersion()
    GetVersion = "1.0.0"
End Function

' 测试有参数 Function
Function Square(n)
    Square = n * n
End Function

' 测试多参数 Function
Function Add(a, b)
    Add = a + b
End Function

' 测试复杂计算 Function
Function Celsius(fDegrees)
    Celsius = (fDegrees - 32) * 5 / 9
End Function

' 测试字符串处理 Function
Function Concat(str1, str2)
    Concat = str1 & str2
End Function

' 测试布尔返回 Function
Function IsEven(n)
    If n Mod 2 = 0 Then
        IsEven = True
    Else
        IsEven = False
    End If
End Function

' 调用测试
Response.Write("<h3>Function 过程测试</h3>")
Response.Write("1. 无参数 Function: " & GetVersion() & "<br>")
Response.Write("2. 有参数 Function Square(5): " & Square(5) & "<br>")
Response.Write("3. 多参数 Function Add(3, 4): " & Add(3, 4) & "<br>")
Response.Write("4. 华氏转摄氏 Celsius(100): " & Celsius(100) & "<br>")
Response.Write("5. 字符串连接 Concat(""Hello"", ""World""): " & Concat("Hello", "World") & "<br>")
Response.Write("6. 布尔返回 IsEven(4): " & IsEven(4) & "<br>")
Response.Write("7. 布尔返回 IsEven(5): " & IsEven(5) & "<br>")

' 在表达式中使用
Dim result
result = Square(Add(2, 3))
Response.Write("8. 嵌套调用 Square(Add(2, 3)): " & result & "<br>")
%>
