<%@ Language="VBScript" CodePage="65001" %>
<%
Option Explicit
Session.CodePage = 65001
Response.CodePage = 65001
Response.Charset = "UTF-8"

' 测试 Call 语句调用 Sub 和 Function

Sub MySub1()
    Response.Write("MySub1 被调用<br>")
End Sub

Sub MySub2(a, b)
    Response.Write("MySub2 被调用, 参数: " & a & ", " & b & "<br>")
End Sub

Function MyFunc1()
    MyFunc1 = "MyFunc1 返回值"
End Function

Function MyFunc2(x, y)
    MyFunc2 = x + y
End Function

' 调用测试
Response.Write("<h3>Call 语句测试</h3>")

Response.Write("1. 使用 Call 调用无参数 Sub:<br>")
Call MySub1()

Response.Write("<br>2. 不使用 Call 调用无参数 Sub:<br>")
MySub1

Response.Write("<br>3. 使用 Call 调用有参数 Sub:<br>")
Call MySub2(10, 20)

Response.Write("<br>4. 不使用 Call 调用有参数 Sub (无括号):<br>")
MySub2 10, 20

Response.Write("<br>5. 使用 Call 调用无参数 Function (丢弃返回值):<br>")
Call MyFunc1()

Response.Write("<br>6. 使用 Call 调用有参数 Function (丢弃返回值):<br>")
Call MyFunc2(5, 10)

Response.Write("<br>7. 不使用 Call 调用 Function (获取返回值):<br>")
Dim result
result = MyFunc2(5, 10)
Response.Write("结果: " & result & "<br>")

Response.Write("<br>8. 在表达式中使用 Function:<br>")
Response.Write("直接输出: " & MyFunc2(100, 200) & "<br>")

Response.Write("<br>9. Function 嵌套调用:<br>")
Response.Write("MyFunc2(MyFunc2(1, 2), 3) = " & MyFunc2(MyFunc2(1, 2), 3) & "<br>")
%>
