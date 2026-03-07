<%@ Language="VBScript" CodePage="65001" %>
<%
Option Explicit
Session.CodePage = 65001
Response.CodePage = 65001
Response.Charset = "UTF-8"

' 测试 Sub 过程
' 测试无参数 Sub

Sub SayHello()
    Response.Write("Hello, World!<br>")
End Sub

' 测试有参数 Sub
Sub Greet(name)
    Response.Write("Hello, " & name & "!<br>")
End Sub

' 测试多参数 Sub
Sub AddAndPrint(a, b)
    Dim result
    result = a + b
    Response.Write(a & " + " & b & " = " & result & "<br>")
End Sub

' 测试空语句 Sub
Sub EmptySub()
End Sub

' 调用测试
Response.Write("<h3>Sub 过程测试</h3>")
Response.Write("1. 无参数 Sub: ")
Call SayHello()

Response.Write("2. 有参数 Sub: ")
Call Greet("VBScript")

Response.Write("3. 多参数 Sub: ")
Call AddAndPrint(10, 20)

Response.Write("4. 空 Sub: ")
Call EmptySub()
Response.Write("执行完成<br>")

Response.Write("5. 不使用 Call 调用: ")
SayHello

Response.Write("6. 不使用 Call 调用带参数: ")
Greet "Developer"
%>
