<%
' 测试无括号函数调用
Function TestFunc(x)
    TestFunc = x * 2
End Function

' 有括号调用
Response.Write "有括号: " & TestFunc(5) & "<br>"

' 无括号调用
Response.Write "无括号: " & TestFunc 5 & "<br>"

' 多参数无括号调用
Function Add(a, b)
    Add = a + b
End Function

Response.Write "无括号多参数: " & Add 3, 4 & "<br>"
%>
