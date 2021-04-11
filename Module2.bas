Attribute VB_Name = "Filesearch"


Sub GetPath(ByVal FilePath As String, ByVal list As ListBox)
 '获取文件路径
 Dim strX As String
 FilePath = IIf(Right(FilePath, 1) = "\", FilePath, FilePath & "\")
 '获取当前目录内的文件名
 Dim fileName As String
 fileName = Dir(FilePath) '初次使用dir函数需指明路径
 '使用一个循环，遍历当前目录内的文件，并逐一验证其属性
 Do While fileName <> ""
 If Right(fileName, 9) = "Gpoint.ta" Then
 strX = Replace(FilePath & fileName, "\\", "\")
 list.AddItem strX
 ElseIf Right(fileName, 10) = "Routing.la" Then
 strX = Replace(FilePath & fileName, "\\", "\")
 list.AddItem strX
 ElseIf Right(fileName, 11) = "Boundary.la" Then
 strX = Replace(FilePath & fileName, "\\", "\")
 list.AddItem strX
 ElseIf Right(fileName, 9) = "Gpoint.ta" Then
 strX = Replace(FilePath & fileName, "\\", "\")
 list.AddItem strX
 ElseIf Right(fileName, 9) = "Sample.ta" Then
 strX = Replace(FilePath & fileName, "\\", "\")
 list.AddItem strX
 ElseIf Right(fileName, 3) = ".db" Then
 strX = Replace(FilePath & fileName, "\\", "\")
 list.AddItem strX
 End If
 fileName = Dir
 Loop
 '缺少此句只会遍历一级目录
 fileName = LCase(Dir(FilePath, vbDirectory))
 Dim ChildContent() As String
 Dim Count As Integer
 '获取下一级目录
 Do While fileName <> ""
 If fileName <> "." And fileName <> ".." Then
 If GetAttr(FilePath & fileName) And vbDirectory Then
 Count = Count + 1
 ReDim Preserve ChildContent(Count)
 '将下一级目录放入动态数组
 ChildContent(Count) = FilePath & "\" & fileName
 End If
 End If
 fileName = Dir
 DoEvents
 Loop
 '回调自身,获取下一级目录内文件路径
 Dim i As Integer
 For i = 1 To Count
 GetPath ChildContent(i), list
 Next i
End Sub
