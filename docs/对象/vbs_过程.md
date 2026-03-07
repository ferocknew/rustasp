在 VBScript 中，过程被分为两类：Sub 过程和 Function 过程。

## [#](https://www.weistock.com/docs/VBA/VBScript/过程.html#sub-过程)Sub 过程

Sub 过程是包含在 Sub 和 End Sub 语句之间的一组 VBScript 语句，执行操作但不返回值。Sub 过程可以使用参数（由调用过程传递的常数、变量或表达式）。如果 Sub 过程无任何参数，则 Sub 语句必须包含空括号 ()。

下面的 Sub 过程使用两个固有的（或内置的）VBScript 函数，即 MsgBox 和 InputBox，来提示用户输入信息。然后显示根据这些信息计算的结果。计算由使用 VBScript 创建的 Function 过程完成。此过程在以下讨论之后演示。

```vb
Sub ConvertTemp()
   temp = InputBox("请输入华氏温度。", 1)
   MsgBox "温度为 " & Celsius(temp) & " 摄氏度。"
End Sub
```

## [#](https://www.weistock.com/docs/VBA/VBScript/过程.html#function-过程)Function 过程

Function 过程是包含在 Function 和 End Function 语句之间的一组 VBScript 语句。Function 过程与 Sub 过程类似，但是 Function 过程可以返回值。Function 过程可以使用参数（由调用过程传递的常数、变量或表达式）。如果 Function 过程无任何参数，则 Function 语句必须包含空括号 ()。Function 过程通过函数名返回一个值，这个值是在过程的语句中赋给函数名的。Function 返回值的数据类型总是 Variant。

在下面的示例中，Celsius 函数将华氏度换算为摄氏度。Sub 过程 ConvertTemp 调用此函数时，包含参数值的变量被传递给函数。换算结果返回到调用过程并显示在消息框中。

```vb
Sub ConvertTemp()
   temp = InputBox("请输入华氏温度。", 1)
   MsgBox "温度为 " & Celsius(temp) & " 摄氏度。"
End Sub

Function Celsius(fDegrees)
   Celsius = (fDegrees - 32) * 5 / 9
End Function
```

## [#](https://www.weistock.com/docs/VBA/VBScript/过程.html#过程的数据进出)过程的数据进出

给过程传递数据的途径是使用参数。参数被作为要传递给过程的数据的占位符。参数名可以是任何有效的变量名。使用 Sub 语句或 Function 语句创建过程时，过程名之后必须紧跟括号。括号中包含所有参数，参数间用逗号分隔。例如，在下面的示例中，fDegrees 是传递给 Celsius 函数的值的占位符：

```vb
Function Celsius(fDegrees)
   Celsius = (fDegrees - 32) * 5 / 9
End Function
```

要从过程获取数据，必须使用 Function 过程。请记住，Function 过程可以返回值；Sub 过程不返回值。

在代码中使用 Sub 和 Function 过程 调用 Function 过程时，函数名必须用在变量赋值语句的右端或表达式中。例如：

```vb
 Temp = Celsius(fDegrees)
```

或

```vb
 MsgBox "温度为 " & Celsius(fDegrees) & " 摄氏度。"
```

调用 Sub 过程时，只需输入过程名及所有参数值，参数值之间使用逗号分隔。不需使用 Call 语句，但如果使用了此语句，则必须将所有参数包含在括号之中。

下面的示例显示了调用 MyProc 过程的两种方式。一种使用 Call 语句；另一种则不使用。两种方式效果相同。

```vb
Call MyProc(firstarg, secondarg)
MyProc firstarg, secondarg
```

请注意当不使用 Call 语句进行调用时，括号被省略。