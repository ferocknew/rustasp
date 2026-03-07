##  Class 对象

使用 Class 语句创建的对象。提供了对类的各种事件的访问

### [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#事件)事件

| 事件                                                         | 描述                                   |
| ------------------------------------------------------------ | -------------------------------------- |
| [Initialize](https://www.weistock.com/docs/VBA/VBScript/事件/Initialize.html) | 当创建相关类的一个实例时将产生此事件。 |
| [Terminate](https://www.weistock.com/docs/VBA/VBScript/事件/Terminate.html) | 当相关类的一个实例结束时将发生此事件。 |

说明

不允许显式地将一个变量声明为 Class 类型。在 VBScript 的上下文中，“类对象”一词指的是用 VBScript Class 语句定义的任何对象。 在使用 Class 语句建立了类定义之后，可以用下面的形式创建类的一个实例：

```vb
Dim X
Set X = New classname
```

由于 VBScript 是一种后期约束型语言，下面的做法是不允许的：

```vb
Dim X as New classname
Dim X
X = New classname
Set X = New Scripting.FileSystemObject
```



## [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#dictionary-对象)Dictionary 对象

保存数据键和项目对的对象。

CreateObject("Scripting.Dictionary")



### [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#属性)属性

| 属性                                                         | 描述                                                         |
| ------------------------------------------------------------ | ------------------------------------------------------------ |
| [CompareMode](https://www.weistock.com/docs/VBA/VBScript/属性/CompareMode.html) | 设置并返回在 Dictionary 对象中比较字符串关键字的比较模式。   |
| [Count](https://www.weistock.com/docs/VBA/VBScript/属性/Count.html) | 返回一个集合或 Dictionary 对象包含的项目数。只读。           |
| [Item](https://www.weistock.com/docs/VBA/VBScript/属性/Item.html) | 设置或返回 Dictionary 对象中指定的 key 对应的 item，或返回集合中基于指定的 key 的 item。可读写。 |
| [Key](https://www.weistock.com/docs/VBA/VBScript/属性/Key.html) | 在 Dictionary 对象中设置 key。                               |

### [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#方法)方法

| 方法                                                         | 描述                                                         |
| ------------------------------------------------------------ | ------------------------------------------------------------ |
| [Add](https://www.weistock.com/docs/VBA/VBScript/方法/Add).html | 向 Dictionary 对象添加键和项目对。                           |
| [Exists](https://www.weistock.com/docs/VBA/VBScript/方法/Exists.html) | 如果在 Dictionary 对象中存在指定键，返回 True；如果不存在，返回 False。 |
| [Items](https://www.weistock.com/docs/VBA/VBScript/方法/Items.html) | 返回一个数组，其中包含有 Dictionary 对象中的所有项目。       |
| [Keys](https://www.weistock.com/docs/VBA/VBScript/方法/Keys.html) | 返回一数组，其中包含有 Dictionary 对象的所有现存键。         |
| [Remove](https://www.weistock.com/docs/VBA/VBScript/方法/Remove.html) | 从 Dictionary 对象中删除键和项目对。                         |
| [RemoveAll](https://www.weistock.com/docs/VBA/VBScript/方法/RemoveAll.html) | RemoveAll 方法删除 Dictionary 对象中的所有键和项目对。       |

说明

Dictionary 对象与 PERL 关联数组是等价的。项目（可以是任何形式的数据）被保存在数组中。每项都与唯一的键相关联。键值用于检索单个项目，通常是整数或字符串，但不能为数组。

示例

如何创建 Dictionary 对象

```vb
Dim d                   '创建一个变量。
Set d = CreateObject("Scripting.Dictionary")
d.Add "a", "Athens"     '添加键和项目。
d.Add "b", "Belgrade"
d.Add "c", "Cairo"
```



## [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#drive-对象)Drive 对象

提供对磁盘驱动器或网络共享的属性的访问。

CreateObject("Scripting.FileSystemObject")



说明

以下代码举例说明如何使用 Drive 对象访问驱动器的属性：

示例

```vb
Function ShowFreeSpace(drvPath)
Dim fso, d, s
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set d = fso.GetDrive(fso.GetDriveName(drvPath))
    s = "驱动器 " & UCase(drvPath) & " - " 
    s = s & d.VolumeName  & "<BR>"
    s = s & "可用空间： " & FormatNumber(d.FreeSpace/1024, 0) 
s = s & "KB"
ShowFreeSpace = s
End Function
```

| 属性                                                         | 描述                                                     |
| ------------------------------------------------------------ | -------------------------------------------------------- |
| [AvailableSpace](https://www.weistock.com/docs/VBA/VBScript/属性/AvailableSpace.html) | 返回指定的驱动器或网络共享对于用户的可用空间大小。       |
| [DriveLetter](https://www.weistock.com/docs/VBA/VBScript/属性/DriveLetter.html) | 返回本地驱动器或网络共享的驱动器号。只读。               |
| [DriveType](https://www.weistock.com/docs/VBA/VBScript/属性/DriveType.html) | 返回一个描述指定驱动器的类型的值。                       |
| [FileSystem](https://www.weistock.com/docs/VBA/VBScript/属性/FileSystem.html) | 返回指定的驱动器使用的文件系统的类型。                   |
| [FreeSpace](https://www.weistock.com/docs/VBA/VBScript/属性/FreeSpace.html) | 返回指定的驱动器或网络共享对于用户的可用空间大小。只读。 |
| [IsReady](https://www.weistock.com/docs/VBA/VBScript/属性/IsReady.html) | 如果指定的驱动器就绪，返回 True；否则返回 False。        |
| [Path](https://www.weistock.com/docs/VBA/VBScript/属性/Path.html) | 返回指定文件、文件夹或驱动器的路径。                     |
| [RootFolder](https://www.weistock.com/docs/VBA/VBScript/属性/RootFolder.html) | 返回一个 Folder 对象，表示指定驱动器的根文件夹。只读。   |
| [SerialNumber](https://www.weistock.com/docs/VBA/VBScript/属性/SerialNumber.html) | 返回十进制序列号，用于唯一标识一个磁盘卷。               |
| [ShareName](https://www.weistock.com/docs/VBA/VBScript/属性/ShareName.html) | 返回指定的驱动器的网络共享名。                           |
| [TotalSize](https://www.weistock.com/docs/VBA/VBScript/属性/TotalSize.html) | 返回驱动器或网络共享的总字节数。                         |
| [VolumeName](https://www.weistock.com/docs/VBA/VBScript/属性/VolumeName.html) | 设置或返回指定驱动器的卷标。可读写。                     |



## [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#drives-集合)Drives 集合

只读所有可用驱动器的集合。

说明 无论是否插入媒体，可移动媒体驱动器都显示在 Drives 集合中。

以下代码举例说明如何获得 Drives 集合并使用 For Each...Next 语句枚举集合成员：

```vb
ShowDriveList 函数
    Dim fso, d, dc, s, n
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set dc = fso.Drives
    For Each d in dc
     n = ""
   s = s & d.DriveLetter & " - " 
        If d.DriveType = Remote Then
            n = d.ShareName
        ElseIf d.IsReady Then
            n = d.VolumeName
        End If
        s = s & n &"<BR>" 
    Next
   ShowDriveList = s
End Function
```

### [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#属性-2)属性

| 属性                                                         | 描述                                                         |
| ------------------------------------------------------------ | ------------------------------------------------------------ |
| [Count](https://www.weistock.com/docs/VBA/VBScript/属性/Count.html) | 返回一个集合或 Dictionary 对象包含的项目数。只读             |
| [Item](https://www.weistock.com/docs/VBA/VBScript/属性/Item.html) | 设置或返回 Dictionary 对象中指定的 key 对应的 item，或返回集合中基于指定的 key 的 item。可读写。 |



## [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#err-对象)Err 对象

**含有关于运行时错误的信息。接受用于生成和清除运行时错误的 Raise 和 Clear 方法。**

说明

Err 对象是一个具有全局范围的固有对象：不必在您的代码中创建它的示例。Err的属性被一个错误的生成器设置：Visual Basic,自动对象，或 VBScript 程序。 Err 对象的默认属性是 number。Err.Number 含有一个整数，且可由 Automation 对象使用以返回 SCODE。 当发生运行时错误时，Err 的属性由标识错误的唯一信息以及可用于处理它的信息填充。要在代码中生成运行时错误，请用 Raise 方法。 Err 对象属性被重新设置为零或零长度字符串 ("")。Clear 方法可被用于显式地重新设置 Err。

### [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#属性-3)属性

| 属性                                                         | 描述                                                   |
| ------------------------------------------------------------ | ------------------------------------------------------ |
| [Description](https://www.weistock.com/docs/VBA/VBScript/属性/Description.html) | 返回或设置与错误相关联的说明性字符串。                 |
| [HelpContext](https://www.weistock.com/docs/VBA/VBScript/属性/HelpContext.html) | 设置或返回帮助文件主题的上下文 ID。                    |
| [HelpFile](https://www.weistock.com/docs/VBA/VBScript/属性/HelpFile.html) | 设置或返回帮助文件的完整有效路径。                     |
| [Number](https://www.weistock.com/docs/VBA/VBScript/属性/Number.html) | 返回或设置数值指定错误。Number 是 Err 对象的默认属性。 |
| [Source](https://www.weistock.com/docs/VBA/VBScript/属性/Source.html) | 返回或设置最初生成错误的对象或应用程序的名称。         |

### [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#方法-2)方法

| 方法                                                         | 描述                          |
| ------------------------------------------------------------ | ----------------------------- |
| [Clear](https://www.weistock.com/docs/VBA/VBScript/方法/Clear.html) | 清除 Err 对象的所有属性设置。 |
| [Raise](https://www.weistock.com/docs/VBA/VBScript/方法/Raise.html) | 产生一个运行时错误。          |

示例

下面的示例说明了 Err 对象的用法：

```vb
On Error Resume Next
Err.Raise 6  '产生溢出错误。
MsgBox ("Error # " & CStr(Err.Number) & " " & Err.Description)
Err.Clear    '清除错误。
```



## [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#file-对象)File 对象

提供对文件的所有属性的访问。

说明

以下代码举例说明如何获得一个 File 对象并查看它的属性：

```vb
 Function ShowDateCreated(filespec)
   Dim fso,f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(filespec)
  ShowDateCreated = f.DateCreated
End Function
```

### [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#属性-4)属性

| 属性                                                         | 描述                                                         |
| ------------------------------------------------------------ | ------------------------------------------------------------ |
| [Attributes](https://www.weistock.com/docs/VBA/VBScript/属性/Attributes.html) | 设置或返回文件或文件夹的属性。可读写或只读（与属性有关）。   |
| [DateCreated](https://www.weistock.com/docs/VBA/VBScript/属性/DateCreated.html) | 返回指定的文件或文件夹的创建日期和时间。只读。               |
| [DateLastAccessed](https://www.weistock.com/docs/VBA/VBScript/属性/DateLastAccessed.html) | 返回指定的文件或文件夹的上次访问日期和时间。只读。           |
| [DateLastModified](https://www.weistock.com/docs/VBA/VBScript/属性/DateLastModified.html) | 返回指定的文件或文件夹的上次修改日期和时间。只读。           |
| [Drive](https://www.weistock.com/docs/VBA/VBScript/属性/Drive.html) | 返回指定的文件或文件夹所在的驱动器的驱动器号。只读。         |
| [Name](https://www.weistock.com/docs/VBA/VBScript/属性/Name.html) | 设置或返回指定的文件或文件夹的名称。可读写。                 |
| [ParentFolder](https://www.weistock.com/docs/VBA/VBScript/属性/ParentFolder.html) | 返回指定文件或文件夹的父文件夹。只读。                       |
| [Path](https://www.weistock.com/docs/VBA/VBScript/属性/Path.html) | 返回指定文件、文件夹或驱动器的路径。                         |
| [ShortName](https://www.weistock.com/docs/VBA/VBScript/属性/ShortName.html) | 返回按照早期 8.3 文件命名约定转换的短文件名。                |
| [ShortPath](https://www.weistock.com/docs/VBA/VBScript/属性/ShortPath.html) | 返回按照 8.3 命名约定转换的短路径名。                        |
| [Size](https://www.weistock.com/docs/VBA/VBScript/属性/Size.html) | 对于文件，返回指定文件的字节数；对于文件夹，返回该文件夹中所有文件和子文件夹的字节数。 |
| [Type](https://www.weistock.com/docs/VBA/VBScript/属性/Type.html) | 返回文件或文件夹的类型信息。例如，对于扩展名为 .TXT 的文件，返回“Text Document”。 |

### [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#方法-3)方法

| 方法                                                         | 描述                                                         |
| ------------------------------------------------------------ | ------------------------------------------------------------ |
| [Copy](https://www.weistock.com/docs/VBA/VBScript/方法/Copy.html) | 将指定的文件或文件夹从某位置复制到另一位置。                 |
| [Delete](https://www.weistock.com/docs/VBA/VBScript/方法/Delete.html) | 删除指定的文件或文件夹。                                     |
| [Move](https://www.weistock.com/docs/VBA/VBScript/方法/Move.html) | 将指定的文件或文件夹从某位置移动到另一位置。                 |
| [OpenAsTextStream](https://www.weistock.com/docs/VBA/VBScript/方法/OpenAsTextStream.html) | 打开指定的文件并返回一个 TextStream 对象，此对象用于对文件进行读、写或追加操作。 |



## [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#files-集合)Files 集合

文件夹中所有 File 对象的集合.

说明 以下代码举例说明如何获得 Folders 集合并使用 For Each...Next 语句枚举集合成员：

```vb
Function ShowFolderList(folderspec)
    Dim fso, f, f1, fc, s
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFolder(folderspec)
    Set fc = f.Files
    For Each f1 in fc
        s = s & f1.name 
        s = s & "<BR>"
    Next
    ShowFolderList = s
End Function
```

### [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#属性-5)属性

| 属性                                                         | 描述                                                         |
| ------------------------------------------------------------ | ------------------------------------------------------------ |
| [Count](https://www.weistock.com/docs/VBA/VBScript/属性/Count.html) | 返回一个集合或 Dictionary 对象包含的项目数。只读             |
| [Item](https://www.weistock.com/docs/VBA/VBScript/属性/Item.html) | 设置或返回 Dictionary 对象中指定的 key 对应的 item，或返回集合中基于指定的 key 的 item。可读写。 |



## [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#filesystemobject-对象)FileSystemObject 对象

提供对计算机文件系统的访问。

[FileSystemObject 基础入门](https://www.weistock.com/docs/VBA/VBScript/FileSystemObject_jc)

### [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#属性-6)属性

| 属性                                                         | 描述                                                |
| ------------------------------------------------------------ | --------------------------------------------------- |
| [Drives](https://www.weistock.com/docs/VBA/VBScript/属性/Drives.html) | 返回由本地机器上所有 Drive 对象组成的 Drives 集合。 |

方法

| 方法                                                         | 描述                                                         |
| ------------------------------------------------------------ | ------------------------------------------------------------ |
| [BuildPath](https://www.weistock.com/docs/VBA/VBScript/方法/BuildPath.html) | 向现有路径后添加名称。                                       |
| [CopyFile](https://www.weistock.com/docs/VBA/VBScript/方法/CopyFile.html) | 将一个或多个文件从某位置复制到另一位置。                     |
| [CopyFolder](https://www.weistock.com/docs/VBA/VBScript/方法/CopyFolder.html) | 将文件夹从某位置递归复制到另一位置。                         |
| [CreateFolder](https://www.weistock.com/docs/VBA/VBScript/方法/CreateFolder.html) | 创建文件夹。                                                 |
| [CreateTextFile](https://www.weistock.com/docs/VBA/VBScript/方法/CreateTextFile.html) | 创建指定文件并返回 TextStream 对象，该对象可用于读或写创建的文件。 |
| [DeleteFile](https://www.weistock.com/docs/VBA/VBScript/方法/DeleteFile.html) | 删除指定的文件。                                             |
| [DeleteFolder](https://www.weistock.com/docs/VBA/VBScript/方法/DeleteFolder.html) | 删除指定的文件夹和其中的内容。                               |
| [DriveExists](https://www.weistock.com/docs/VBA/VBScript/方法/DriveExists.html) | 如果指定的驱动器存在，则返回 True；否则返回 False。          |
| [FileExists](https://www.weistock.com/docs/VBA/VBScript/方法/FileExists.html) | 如果指定的文件存在返回 True；否则返回 False。                |
| [FolderExists](https://www.weistock.com/docs/VBA/VBScript/方法/FolderExists.html) | 如果指定的文件夹存在，则返回 True；否则返回 False。          |
| [GetAbsolutePathname](https://www.weistock.com/docs/VBA/VBScript/方法/GetAbsolutePathname.html) | 从提供的指定路径中返回完整且含义明确的路径。                 |
| [GetBaseName](https://www.weistock.com/docs/VBA/VBScript/方法/GetBaseName.html) | 返回字符串，其中包含文件的基本名 (不带扩展名), 或者提供的路径说明中的文件夹。 |
| [GetDrive](https://www.weistock.com/docs/VBA/VBScript/方法/GetDrive.html) | 返回与指定的路径中驱动器相对应的 Drive 对象。                |
| [GetDriveName](https://www.weistock.com/docs/VBA/VBScript/方法/GetDriveName.html) | 返回包含指定路径中驱动器名的字符串。                         |
| [GetExtensionName](https://www.weistock.com/docs/VBA/VBScript/方法/GetExtensionName.html) | 返回字符串，该字符串包含路径最后一个组成部分的扩展名。       |
| [GetFile](https://www.weistock.com/docs/VBA/VBScript/方法/GetFile.html) | 返回与指定路径中某文件相应的 File 对象。                     |
| [GetFileName](https://www.weistock.com/docs/VBA/VBScript/方法/GetFileName.html) | 返回指定路径（不是指定驱动器路径部分）的最后一个文件或文件夹。 |
| [GetFolder](https://www.weistock.com/docs/VBA/VBScript/方法/GetFolder.html) | 返回与指定的路径中某文件夹相应的 Folder 对象。               |
| [GetParentFolderName](https://www.weistock.com/docs/VBA/VBScript/方法/GetParentFolderName.html) | 返回字符串，该字符串包含指定的路径中最后一个文件或文件夹的父文件夹。 |
| [GetSpecialFolder](https://www.weistock.com/docs/VBA/VBScript/方法/GetSpecialFolder.html) | 返回指定的特殊文件夹。                                       |
| [GetTempName](https://www.weistock.com/docs/VBA/VBScript/方法/GetTempName.html) | 返回随机生成的临时文件或文件夹的名称，用于执行要求临时文件或文件夹的操作。 |
| [MoveFile](https://www.weistock.com/docs/VBA/VBScript/方法/MoveFile.html) | 将一个或多个文件从某位置移动到另一位置。                     |
| [MoveFolder](https://www.weistock.com/docs/VBA/VBScript/方法/MoveFolder.html) | 将一个或多个文件夹从某位置移动到另一位置。                   |
| [OpenTextFile](https://www.weistock.com/docs/VBA/VBScript/方法/OpenTextFile.html) | 打开指定的文件并返回一个 [TextStream 对象](https://www.weistock.com/docs/VBA/VBScript/对象.html#TextStream)，可以读取、写入此对象或将其追加到文件。 |

在该例子代码中，ActiveXObject 对象被赋给 FileSystemObject (fso)。随后 CreateTextFile 方法创建文件 TextStream 对象 (a)，并用 WriteLine 方法将一行文本写入创建的文本文件。Close 方法刷新缓冲区并关闭该文件。

说明 以下代码举例说明如何使用 FileSystemObject 对象返回一个 [TextStream 对象](https://www.weistock.com/docs/VBA/VBScript/对象.html#TextStream)，此对象可以被读取或写入：

```vb
Dim fso, MyFile
Set fso = CreateObject("Scripting.FileSystemObject")
Set MyFile = fso.CreateTextFile("c:\testfile.txt", True)
MyFile.WriteLine("This is a test.")
MyFile.Close
```

在上述代码中，CreateObject 函数返回 FileSystemObject (fso)。然后 CreateTextFile 方法创建一个作为 TextStream 对象 (a) 的文件。之后 WriteLine 方法在所创建的文件中写入一行文本；Close 方法刷新缓冲区，并关闭该文件。



## [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#folder-对象)Folder 对象

提供对文件夹所有属性的访问。

说明 以下代码举例说明如何获得 Folder 对象并查看它的属性：

```vb
Function ShowDateCreated(folderspec)
    Dim fso, f, 
    Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.GetFolder(folderspec)
  ShowDateCreated = f.DateCreated
End Function
```

### [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#属性-7)属性

| 属性                                                         | 描述                                                         |
| ------------------------------------------------------------ | ------------------------------------------------------------ |
| [Attributes](https://www.weistock.com/docs/VBA/VBScript/属性/Attributes.html) | 设置或返回文件或文件夹的属性。可读写或只读（与属性有关）。   |
| [DateCreated](https://www.weistock.com/docs/VBA/VBScript/属性/DateCreated.html) | 返回指定的文件或文件夹的创建日期和时间。只读。               |
| [DateLastAccessed](https://www.weistock.com/docs/VBA/VBScript/属性/DateLastAccessed.html) | 返回指定的文件或文件夹的上次访问日期和时间。只读。           |
| [DateLastModified](https://www.weistock.com/docs/VBA/VBScript/属性/DateLastModified.html) | 返回指定的文件或文件夹的上次修改日期和时间。只读。           |
| [Drive](https://www.weistock.com/docs/VBA/VBScript/属性/Drive.html) | 返回指定的文件或文件夹所在的驱动器的驱动器号。只读。         |
| [Files](https://www.weistock.com/docs/VBA/VBScript/属性/Files.html) | 返回由指定文件夹中所有 File 对象（包括隐藏文件和系统文件）组成的 Files 集合。 |
| [IsRootFolder](https://www.weistock.com/docs/VBA/VBScript/属性/IsRootFolder.html) | 如果指定的文件夹是根文件夹，返回 True；否则返回 False。      |
| [Name](https://www.weistock.com/docs/VBA/VBScript/属性/Name.html) | 设置或返回指定的文件或文件夹的名称。可读写。                 |
| [ParentFolder](https://www.weistock.com/docs/VBA/VBScript/属性/ParentFolder.html) | 返回指定文件或文件夹的父文件夹。只读。                       |
| [Path](https://www.weistock.com/docs/VBA/VBScript/属性/Path.html) | 返回指定文件、文件夹或驱动器的路径。                         |
| [ShortName](https://www.weistock.com/docs/VBA/VBScript/属性/ShortName.html) | 返回按照早期 8.3 文件命名约定转换的短文件名。                |
| [ShortPath](https://www.weistock.com/docs/VBA/VBScript/属性/ShortPath.html) | 返回按照 8.3 命名约定转换的短路径名。                        |
| [Size](https://www.weistock.com/docs/VBA/VBScript/属性/Size.html) | 对于文件，返回指定文件的字节数；对于文件夹，返回该文件夹中所有文件和子文件夹的字节数。 |
| [SubFolders](https://www.weistock.com/docs/VBA/VBScript/属性/SubFolders.html) | 返回由指定文件夹中所有子文件夹（包括隐藏文件夹和系统文件夹）组成的 Folders 集合。 |
| [Type](https://www.weistock.com/docs/VBA/VBScript/属性/Type.html) | 返回文件或文件夹的类型信息。例如，对于扩展名为 .TXT 的文件，返回“Text Document”。 |

### [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#方法-4)方法

| 方法                                                         | 描述                                                         |
| ------------------------------------------------------------ | ------------------------------------------------------------ |
| [Copy](https://www.weistock.com/docs/VBA/VBScript/方法/Copy.html) | 将指定的文件或文件夹从某位置复制到另一位置。                 |
| [Delete](https://www.weistock.com/docs/VBA/VBScript/方法/Delete.html) | 删除指定的文件或文件夹。                                     |
| [Move](https://www.weistock.com/docs/VBA/VBScript/方法/Move.html) | 将指定的文件或文件夹从某位置移动到另一位置。                 |
| [CreateTextFile](https://www.weistock.com/docs/VBA/VBScript/方法/CreateTextFile.html) | 创建指定文件并返回 TextStream 对象，该对象可用于读或写创建的文件。 |



## [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#folders-集合)Folders 集合

包含在一个 Folder 对象中的所有 Folder 对象的集合。

说明 以下代码举例说明如何获得 Folders 集合并使用 For Each...Next 语句枚举集合成员：

```vb
 Function ShowFolderList(folderspec) 
    Dim fso, f, f1, fc, s
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFolder(folderspec)
    Set fc = f.SubFolders
    For Each f1 in fc
        s = s & f1.name 
        s = s & "<BR>"
    Next
    ShowFolderList = s
End Function
```

### [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#属性-8)属性

| 属性                                                         | 描述                                                         |
| ------------------------------------------------------------ | ------------------------------------------------------------ |
| [Count](https://www.weistock.com/docs/VBA/VBScript/属性/Count.html) | 返回一个集合或 Dictionary 对象包含的项目数。只读             |
| [Item](https://www.weistock.com/docs/VBA/VBScript/属性/Item.html) | 设置或返回 Dictionary 对象中指定的 key 对应的 item，或返回集合中基于指定的 key 的 item。可读写。 |

### [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#方法-5)方法

| 方法                                                         | 描述                                               |
| ------------------------------------------------------------ | -------------------------------------------------- |
| [Add](https://www.weistock.com/docs/VBA/VBScript/方法/Add.html) | 包含在一个 Folder 对象中的所有 Folder 对象的集合。 |



## [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#match-对象)Match 对象

提供对一个正则表达式匹配的只读属性的访问途径功能。

说明 Match 对象只能通过 RegExp 对象的 Execute 方法来创建，该方法实际上返回了 Match 对象的集合。所有的 Match 对象属性都是只读的。

在执行正则表达式时，可能产生零个或多个 Match 对象。每个 Match 对象提供了被正则表达式搜索找到的字符串的访问、字符串的长度，以及找到匹配的索引位置等。

下面的代码说明了 Match 对象的用法：

```vb
Function RegExpTest(patrn, strng)
  Dim regEx, Match, Matches         ' 建立变量。
  Set regEx = New RegExp         ' 建立正则表达式。
  regEx.Pattern = patrn         ' 设置模式。
  regEx.IgnoreCase = True         ' 设置是否区分大小写。
  regEx.Global = True            ' 设置全局替换。
  Set Matches = regEx.Execute(strng)      ' 执行搜索。
  For Each Match in Matches         ' 遍历 Matches 集合。
    RetStr = RetStr & "Match " & I & " found at position "
    RetStr = RetStr & Match.FirstIndex & ". Match Value is "'
    RetStr = RetStr & Match.Value & "'." & vbCRLF
  Next
  RegExpTest = RetStr
End Function
MsgBox(RegExpTest("is.", "IS1 is2 IS3 is4"))
```

### [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#属性-9)属性

| 属性                                                         | 描述                                           |
| ------------------------------------------------------------ | ---------------------------------------------- |
| [FirstIndex](https://www.weistock.com/docs/VBA/VBScript/属性/FirstIndex.html) | 返回搜索字符串中找到匹配项的位置。             |
| [Length](https://www.weistock.com/docs/VBA/VBScript/属性/Length.html) | 返回搜索字符串中所找到的匹配的长度。           |
| [Value](https://www.weistock.com/docs/VBA/VBScript/属性/Value.html) | 返回在一个搜索字符串中找到的匹配项的值或文本。 |



## [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#matches-集合)Matches 集合

正则表达式 Match 对象的集合。

说明 Matches 集合中包含若干独立的 Match 对象，只能使用 RegExp 对象的 Execute 方法来创建之。与独立的 Match 对象属性相同，Matches `集合的一个属性是只读的。 在执行正则表达式时，可能产生零个或多个 Match 对象。每个 Match 对象都提供了与正则表达式匹配的字符串的访问入口、字符串的长度，以及标识匹配位置的索引。 下面的代码将说明如何使用正则表达式查找获得 Matches 集合，以及如何循环遍历集合：

```vb
Function RegExpTest(patrn, strng)
  Dim regEx, Match, Matches      ' 创建变量。
  Set regEx = New RegExp         ' 创建正则表达式。
  regEx.Pattern = patrn         ' 设置模式。
  regEx.IgnoreCase = True         ' 设置是否区分大小写。
  regEx.Global = True         ' 设置全程匹配。
  Set Matches = regEx.Execute(strng)   ' 执行搜索。
  For Each Match in Matches      ' 循环遍历Matches集合。
    RetStr = RetStr & "Match found at position "
    RetStr = RetStr & Match.FirstIndex & ". Match Value is '"
    RetStr = RetStr & Match.Value & "'." & vbCRLF
  Next
  RegExpTest = RetStr
End Function
MsgBox(RegExpTest("is.", "IS1 is2 IS3 is4"))
```

### [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#属性-10)属性

| 属性                                                         | 描述                                                         |
| ------------------------------------------------------------ | ------------------------------------------------------------ |
| [Count](https://www.weistock.com/docs/VBA/VBScript/属性/Count.html) | 返回一个集合或 Dictionary 对象包含的项目数。只读             |
| [Item](https://www.weistock.com/docs/VBA/VBScript/属性/Item.html) | 设置或返回 Dictionary 对象中指定的 key 对应的 item，或返回集合中基于指定的 key 的 item。可读写。 |



## [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#regexp-对象)RegExp 对象

提供简单的正则表达式支持功能。

说明 下面的代码说明了RegExp对象的用法：

```vb
Function RegExpTest(patrn, strng)
  Dim regEx, Match, Matches      ' 建立变量。
  Set regEx = New RegExp         ' 建立正则表达式。
  regEx.Pattern = patrn         ' 设置模式。
  regEx.IgnoreCase = True         ' 设置是否区分字符大小写。
  regEx.Global = True         ' 设置全局可用性。
  Set Matches = regEx.Execute(strng)   ' 执行搜索。
  For Each Match in Matches      ' 遍历匹配集合。
    RetStr = RetStr & "Match found at position "
    RetStr = RetStr & Match.FirstIndex & ". Match Value is '"
    RetStr = RetStr & Match.Value & "'." & vbCRLF
  Next
  RegExpTest = RetStr
End Function
MsgBox(RegExpTest("is.", "IS1 is2 IS3 is4"))
```

| 属性                                                         | 描述                                               |
| ------------------------------------------------------------ | -------------------------------------------------- |
| [Global](https://www.weistock.com/docs/VBA/VBScript/属性/Global.html) | 设置或返回一个布尔值。                             |
| [IgnoreCase](https://www.weistock.com/docs/VBA/VBScript/属性/IgnoreCase.html) | 设置或返回一个布尔值，指明模式搜索是否区分大小写。 |
| [Pattern](https://www.weistock.com/docs/VBA/VBScript/属性/Pattern.html) | 设置或返回要被搜索的正则表达式模式。               |

| 方法                                                         | 描述                                   |
| ------------------------------------------------------------ | -------------------------------------- |
| [Execute](https://www.weistock.com/docs/VBA/VBScript/方法/Execute.html) | 对一个指定的字符串进行正则表达式搜索。 |
| [Replace](https://www.weistock.com/docs/VBA/VBScript/方法/Replace.html) | 替换正则表达式搜索中所找到的文本。     |
| [Test](https://www.weistock.com/docs/VBA/VBScript/方法/Test.html) | 对一个指定的字符串进行正则表达式搜索。 |



## [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#submatches-对象)SubMatches 对象

提供对正则表达式子匹配字符串的只读值的访问。

说明

SubMatches 集合包含了单个的子匹配字符串，只能用 RegExp 对象的 Execute 方法创建。SubMatches 集合的属性是只读的。 运行一个正则表达式时，当圆括号中捕捉到子表达式时可以有零个或多个子匹配。SubMatches 集合中的每一项是由正则表达式找到并捕获的的字符串。 下面的代码演示了如何从一个正则表达式获得一个 SubMatches 集合以及如何它的专有成员：

```vb
Function SubMatchTest(inpStr)
  Dim oRe, oMatch, oMatches
  Set oRe = New RegExp
  ' 查找一个电子邮件地址（不是一个理想的 RegExp）
  oRe.Pattern = "(\w+)@(\w+)\.(\w+)"
  ' 得到 Matches 集合
  Set oMatches = oRe.Execute(inpStr)
  ' 得到 Matches 集合中的第一项
  Set oMatch = oMatches(0)
  ' 创建结果字符串。
  ' Match 对象是完整匹配 — dragon@xyzzy.com
  retStr = "电子邮件地址是： " & oMatch & vbNewline
  ' 得到地址的子匹配部分。
  retStr = retStr & "电子邮件别名是： " & oMatch.SubMatches(0)  ' dragon
  retStr = retStr & vbNewline
  retStr = retStr & "组织是： " & oMatch. SubMatches(1)' xyzzy
  SubMatchTest = retStr
End Function

MsgBox(SubMatchTest("请写信到 dragon@xyzzy.com 。 谢谢！"))
```

| 属性                                                         | 描述                                                         |
| ------------------------------------------------------------ | ------------------------------------------------------------ |
| [Count](https://www.weistock.com/docs/VBA/VBScript/属性/Count.html) | 返回一个集合或 Dictionary 对象包含的项目数。只读             |
| [Item](https://www.weistock.com/docs/VBA/VBScript/属性/Item.html) | 设置或返回 Dictionary 对象中指定的 key 对应的 item，或返回集合中基于指定的 key 的 item。可读写。 |



## [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#textstream-对象)TextStream 对象

方便对文件的顺序访问。

TextStream.{property | method( )} property 和 method 参数可以是与 TextStream 对象相连的任何属性和方法。请注意在实际使用时，TextStream 被从 FileSystemObject 返回的代表 TextStream 对象的变量占位符代替。

说明 在下面的代码中，a 是由 FileSystemObject 的 CreateTextFile 方法返回的 TextStream 对象：

```vb
Dim fso, MyFile
Set fso = CreateObject("Scripting.FileSystemObject")
Set MyFile= fso.CreateTextFile("c:\testfile.txt", True)
MyFile.WriteLine("This is a test.")
MyFile.Close
```

WriteLine 和 Close 是 TextStream 对象的两个方法。

### [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#方法-6)方法

| 方法                                                         | 描述                                                         |
| ------------------------------------------------------------ | ------------------------------------------------------------ |
| [Close](https://www.weistock.com/docs/VBA/VBScript/方法/Close.html) | 关闭打开的 TextStream 文件。                                 |
| [Read](https://www.weistock.com/docs/VBA/VBScript/方法/Read.html) | 从TextStream 文件中读取指定数量的字符，并返回由此得到的字符串。 |
| [ReadAll](https://www.weistock.com/docs/VBA/VBScript/方法/ReadAll.html) | 读取 TextStream 文件的全部内容并返回由此得到的字符串。       |
| [ReadLine](https://www.weistock.com/docs/VBA/VBScript/方法/ReadLine.html) | 从TextStream 文件中读取一整行（一直到换行符，但不包括换行符），并返回由此得到的字符串。 |
| [Skip](https://www.weistock.com/docs/VBA/VBScript/方法/Skip.html) | 在读取 TextStream 文件时跳过指定个数的字符。                 |
| [SkipLine](https://www.weistock.com/docs/VBA/VBScript/方法/SkipLine.html) | 在读取TextStream 文件时跳过下一行。                          |
| [Write](https://www.weistock.com/docs/VBA/VBScript/方法/Write.html) | 将给定的字符串写入到一个 TextStream 文件。                   |
| [WriteBlankLines](https://www.weistock.com/docs/VBA/VBScript/方法/WriteBlankLines.html) | 将指定数量的换行符写入到一个 TextStream 文件。               |
| [WriteLine](https://www.weistock.com/docs/VBA/VBScript/方法/WriteLine.html) | 向 TextStream 文件中写入给定的字符串和一个换行符。           |

### [#](https://www.weistock.com/docs/VBA/VBScript/对象.html#属性-11)属性

| 属性                                                         | 描述                                                         |
| ------------------------------------------------------------ | ------------------------------------------------------------ |
| [AtEndOfLine](https://www.weistock.com/docs/VBA/VBScript/属性/AtEndOfLine.html) | 如果文件指针正好位于 TextStream 文件中的行尾符之前，则返回true，否则返回 false。只读。 |
| [AtEndOfStream](https://www.weistock.com/docs/VBA/VBScript/属性/AtEndOfStream.html) | 如果文件指针正好位于 TextStream 文件中的结尾，则返回true，否则返回 false。只读。 |
| [Column](https://www.weistock.com/docs/VBA/VBScript/属性/Column.html) | 为只读属性，返回当前字符在 TextStream 文件中的列号。         |
| [Line](https://www.weistock.com/docs/VBA/VBScript/属性/Line.html) | 只读属性，返回 TextStream 文件中当前的行号。                 |