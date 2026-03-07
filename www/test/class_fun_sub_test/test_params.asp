<%@ Language="VBScript" CodePage="65001" %>
<%
Option Explicit
Session.CodePage = 65001
Response.CodePage = 65001
Response.Charset = "UTF-8"

' 测试参数传递
' 测试按值传递

Sub TestByVal(ByVal x)
    x = x + 10
    Response.Write("Sub 内部 x = " & x & "<br>")
End Sub

' 测试按引用传递
Sub TestByRef(ByRef y)
    y = y + 10
    Response.Write("Sub 内部 y = " & y & "<br>")
End Sub

' 测试默认传递方式(按引用)
Sub TestDefault(z)
    z = z + 10
    Response.Write("Sub 内部 z = " & z & "<br>")
End Sub

' 测试多个参数混合传递
Sub TestMixed(ByVal a, ByRef b, c)
    a = a + 1
    b = b + 1
    c = c + 1
    Response.Write("Mixed: a=" & a & ", b=" & b & ", c=" & c & "<br>")
End Sub

' 测试数组参数
Sub TestArray(arr)
    Response.Write("数组参数: ")
    Dim i
    For i = 0 To UBound(arr)
        Response.Write(arr(i) & " ")
    Next
    Response.Write("<br>")
End Sub

' 调用测试
Response.Write("<h3>参数传递测试</h3>")

Response.Write("1. ByVal 测试:<br>")
Dim val1
val1 = 5
Response.Write("调用前 val1 = " & val1 & "<br>")
TestByVal val1
Response.Write("调用后 val1 = " & val1 & " (应该不变)<br>")

Response.Write("<br>2. ByRef 测试:<br>")
Dim val2
val2 = 5
Response.Write("调用前 val2 = " & val2 & "<br>")
TestByRef val2
Response.Write("调用后 val2 = " & val2 & " (应该改变)<br>")

Response.Write("<br>3. 默认传递(ByRef)测试:<br>")
Dim val3
val3 = 5
Response.Write("调用前 val3 = " & val3 & "<br>")
TestDefault val3
Response.Write("调用后 val3 = " & val3 & " (应该改变)<br>")

Response.Write("<br>4. 混合参数测试:<br>")
Dim a, b, c
a = 10: b = 20: c = 30
Response.Write("调用前: a=" & a & ", b=" & b & ", c=" & c & "<br>")
TestMixed a, b, c
Response.Write("调用后: a=" & a & " (ByVal不变), b=" & b & " (ByRef改变), c=" & c & " (默认ByRef改变)<br>")

Response.Write("<br>5. 数组参数测试:<br>")
Dim arr(4)
arr(0) = 1: arr(1) = 2: arr(2) = 3: arr(3) = 4: arr(4) = 5
TestArray arr
%>
