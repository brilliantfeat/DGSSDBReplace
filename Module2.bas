Attribute VB_Name = "Filesearch"


Sub GetPath(ByVal FilePath As String, ByVal list As ListBox)
 '��ȡ�ļ�·��
 Dim strX As String
 FilePath = IIf(Right(FilePath, 1) = "\", FilePath, FilePath & "\")
 '��ȡ��ǰĿ¼�ڵ��ļ���
 Dim fileName As String
 fileName = Dir(FilePath) '����ʹ��dir������ָ��·��
 'ʹ��һ��ѭ����������ǰĿ¼�ڵ��ļ�������һ��֤������
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
 'ȱ�ٴ˾�ֻ�����һ��Ŀ¼
 fileName = LCase(Dir(FilePath, vbDirectory))
 Dim ChildContent() As String
 Dim Count As Integer
 '��ȡ��һ��Ŀ¼
 Do While fileName <> ""
 If fileName <> "." And fileName <> ".." Then
 If GetAttr(FilePath & fileName) And vbDirectory Then
 Count = Count + 1
 ReDim Preserve ChildContent(Count)
 '����һ��Ŀ¼���붯̬����
 ChildContent(Count) = FilePath & "\" & fileName
 End If
 End If
 fileName = Dir
 DoEvents
 Loop
 '�ص�����,��ȡ��һ��Ŀ¼���ļ�·��
 Dim i As Integer
 For i = 1 To Count
 GetPath ChildContent(i), list
 Next i
End Sub
