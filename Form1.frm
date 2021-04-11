VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "路线数据替换工具"
   ClientHeight    =   6610
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   10790
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   6610
   ScaleMode       =   0  'User
   ScaleWidth      =   10790
   StartUpPosition =   2  '屏幕中心
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9720
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Cmd6 
      Caption         =   "手动添加文件"
      Height          =   730
      Left            =   8880
      TabIndex        =   11
      Top             =   2880
      Width           =   1810
   End
   Begin VB.CommandButton Cmd5 
      Caption         =   "剔除选择文件"
      Height          =   730
      Left            =   8880
      TabIndex        =   10
      Top             =   3720
      Width           =   1810
   End
   Begin VB.ListBox Filelist 
      Height          =   2380
      ItemData        =   "Form1.frx":ABFE
      Left            =   120
      List            =   "Form1.frx":AC00
      TabIndex        =   9
      Top             =   2160
      Width           =   8650
   End
   Begin VB.CommandButton Cmd4 
      Caption         =   "自动查找文件"
      Height          =   730
      Left            =   8880
      TabIndex        =   8
      Top             =   2040
      Width           =   1810
   End
   Begin VB.CommandButton Cmd3 
      Caption         =   "退出"
      Height          =   730
      Left            =   8880
      TabIndex        =   7
      Top             =   4680
      Width           =   1810
   End
   Begin VB.CommandButton Cmd2 
      Caption         =   "开始替换"
      Height          =   740
      Left            =   7200
      TabIndex        =   6
      Top             =   4680
      Width           =   1570
   End
   Begin VB.TextBox T3 
      Height          =   270
      Left            =   600
      TabIndex        =   5
      Top             =   5160
      Width           =   6500
   End
   Begin VB.TextBox T2 
      Height          =   270
      Left            =   600
      TabIndex        =   4
      Top             =   4800
      Width           =   6500
   End
   Begin VB.TextBox T1 
      Height          =   410
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   8650
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "设置数据库目录"
      Height          =   375
      Left            =   8880
      TabIndex        =   0
      Top             =   1560
      Width           =   1810
   End
   Begin VB.Label Label1 
      Height          =   730
      Left            =   240
      TabIndex        =   13
      Top             =   5760
      Width           =   10450
   End
   Begin VB.Label Tixing 
      Caption         =   "重要提醒："
      Height          =   970
      Left            =   1200
      TabIndex        =   12
      Top             =   360
      Width           =   9610
   End
   Begin VB.Label L2 
      Caption         =   "替换："
      Height          =   260
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Width           =   620
   End
   Begin VB.Label L1 
      Caption         =   "查找："
      Height          =   260
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   620
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim cn As New cConnection
Dim Mypath As String

Sub Replaceword(ByVal list1 As ListBox)
Dim i As Long
Dim DGSSDB As New cConnection
On Error Resume Next
For i = 0 To list1.ListCount - 1
If Right(list1.list(i), 9) = "Gpoint.ta" Then
DGSSDB.OpenDB list1.list(i)
DGSSDB.Execute "UPDATE GeoArea set STRAPHA=REPLACE (STRAPHA ,'" & T2.Text & "','" & T3.Text & "'), STRAPHB=REPLACE (STRAPHB ,'" & T2.Text & "','" & T3.Text & "')"
ElseIf Right(list1.list(i), 10) = "Routing.la" Then
DGSSDB.OpenDB list1.list(i)
DGSSDB.Execute "UPDATE GeoArea set STRAPHA=REPLACE (STRAPHA ,'" & T2.Text & "','" & T3.Text & "');"
ElseIf Right(list1.list(i), 11) = "Boundary.la" Then
DGSSDB.OpenDB list1.list(i)
DGSSDB.Execute "UPDATE GeoArea set RIGHT_BODY=REPLACE (RIGHT_BODY ,'" & T2.Text & "','" & T3.Text & "'),LEFT_BODY=RIGHT_BODY=REPLACE (LEFT_BODY ,'" & T2.Text & "','" & T3.Text & "');"
ElseIf Right(list1.list(i), 11) = "Attitude.ta" Then
DGSSDB.OpenDB list1.list(i)
DGSSDB.Execute "UPDATE GeoArea set STRAPH=REPLACE (STRAPH ,'" & T2.Text & "','" & T3.Text & "');"
ElseIf Right(list1.list(i), 9) = "Sample.ta" Then
DGSSDB.OpenDB list1.list(i)
DGSSDB.Execute "UPDATE GeoArea set GEOUNIT=REPLACE (GEOUNIT ,'" & T2.Text & "','" & T3.Text & "');"
ElseIf Right(list1.list(i), 3) = ".db" Then
DGSSDB.OpenDB list1.list(i)
DGSSDB.Execute "UPDATE BOUNDARY set DESC=REPLACE(DESC ,'" & T2.Text & "','" & T3.Text & "');UPDATE GPOINT set DESC=REPLACE(DESC ,'" & T2.Text & "','" & T3.Text & "');UPDATE ROUTE set DESC=REPLACE(DESC ,'" & T2.Text & "','" & T3.Text & "');UPDATE ROUTING set DESC=REPLACE(DESC ,'" & T2.Text & "','" & T3.Text & "');"
Else
'MsgBox ("似乎存在不支持的文件哦，请检查列表!")
End If
Next
End Sub

Private Sub Cmd1_Click()
Dim str
    str = GetFolder(Me.hWnd, "选择DGSS数据库目录")
    If str <> "" Then
        T1.Text = str
    End If
End Sub

Private Sub Cmd2_Click()
If Filelist.ListCount = 0 Then
MsgBox ("预处理列表似乎是空的哦，请先添加文件!")
ElseIf T2.Text = "" Then
MsgBox ("请先输入查找词!")
ElseIf T3.Text = "" Then
MsgBox ("请先输入替换词!")
Else
Replaceword Filelist
MsgBox ("替换完成，请打开DGSS检查!")
End If
End Sub

Private Sub Cmd3_Click()
End
End Sub
Private Sub SeachFile()
GetPath T1.Text, Filelist
End Sub

Private Sub Cmd4_Click()
If T1.Text = "" Then
MsgBox ("先设置数据库目录!")
Else
SeachFile
End If
End Sub

Private Sub Cmd5_Click()
On Error Resume Next
Filelist.RemoveItem Filelist.ListIndex
End Sub


Private Sub Cmd6_Click()
On Error Resume Next
CommonDialog1.Filter = "DGSS数据库文件|*.db;*.ta;*.la;*.pa|所有文件|*.*"
CommonDialog1.ShowOpen
If CommonDialog1.fileName <> "" Then Filelist.AddItem CommonDialog1.fileName
CommonDialog1.fileName = ""
End Sub

Private Sub Filelist_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 46 Then Filelist.RemoveItem Filelist.ListIndex
End Sub


Private Sub Form_Load()
Label1.Caption = "    请先选择数据库目录，然后进行自动查找功能，对查找到的文件进行增减；一般建议选择“DGSS\...\数字填图\野外手图\”目录为操作目录操作，选完目录请自动查找文件并对文件进行筛查,通常需要修改Gpoint、Routing、Boutary、Attitude、Sample及路线.db文件；如果有漏项可手动添加；对不需要更改文件可按del键或点击“剔除选择文件”进行删除；最后定义需要查找的字段和需要替换为的字段进行查找替换，完成之后到DGSS检查是否符合替换要求！"
Tixing.Caption = "重要提醒：" & "1、数据库操作前务必备份数据；" & "2、该程序只可更改地质点、路线、地质界线、产状、样品的填图单位及其描述内容，一切内容以野外实际情况为准，本程序只是辅助更改；" & "3、本程序引起的所有数据库问题责任由您个人承担；" & "4、版权归@brilliantfeat所有，请勿用于商业用途。"
End Sub

