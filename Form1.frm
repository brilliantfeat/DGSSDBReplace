VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "·�������滻����"
   ClientHeight    =   6610
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   10790
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   6610
   ScaleMode       =   0  'User
   ScaleWidth      =   10790
   StartUpPosition =   2  '��Ļ����
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9720
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Cmd6 
      Caption         =   "�ֶ�����ļ�"
      Height          =   730
      Left            =   8880
      TabIndex        =   11
      Top             =   2880
      Width           =   1810
   End
   Begin VB.CommandButton Cmd5 
      Caption         =   "�޳�ѡ���ļ�"
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
      Caption         =   "�Զ������ļ�"
      Height          =   730
      Left            =   8880
      TabIndex        =   8
      Top             =   2040
      Width           =   1810
   End
   Begin VB.CommandButton Cmd3 
      Caption         =   "�˳�"
      Height          =   730
      Left            =   8880
      TabIndex        =   7
      Top             =   4680
      Width           =   1810
   End
   Begin VB.CommandButton Cmd2 
      Caption         =   "��ʼ�滻"
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
      Caption         =   "�������ݿ�Ŀ¼"
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
      Caption         =   "��Ҫ���ѣ�"
      Height          =   970
      Left            =   1200
      TabIndex        =   12
      Top             =   360
      Width           =   9610
   End
   Begin VB.Label L2 
      Caption         =   "�滻��"
      Height          =   260
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Width           =   620
   End
   Begin VB.Label L1 
      Caption         =   "���ң�"
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
'MsgBox ("�ƺ����ڲ�֧�ֵ��ļ�Ŷ�������б�!")
End If
Next
End Sub

Private Sub Cmd1_Click()
Dim str
    str = GetFolder(Me.hWnd, "ѡ��DGSS���ݿ�Ŀ¼")
    If str <> "" Then
        T1.Text = str
    End If
End Sub

Private Sub Cmd2_Click()
If Filelist.ListCount = 0 Then
MsgBox ("Ԥ�����б��ƺ��ǿյ�Ŷ����������ļ�!")
ElseIf T2.Text = "" Then
MsgBox ("����������Ҵ�!")
ElseIf T3.Text = "" Then
MsgBox ("���������滻��!")
Else
Replaceword Filelist
MsgBox ("�滻��ɣ����DGSS���!")
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
MsgBox ("���������ݿ�Ŀ¼!")
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
CommonDialog1.Filter = "DGSS���ݿ��ļ�|*.db;*.ta;*.la;*.pa|�����ļ�|*.*"
CommonDialog1.ShowOpen
If CommonDialog1.fileName <> "" Then Filelist.AddItem CommonDialog1.fileName
CommonDialog1.fileName = ""
End Sub

Private Sub Filelist_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 46 Then Filelist.RemoveItem Filelist.ListIndex
End Sub


Private Sub Form_Load()
Label1.Caption = "    ����ѡ�����ݿ�Ŀ¼��Ȼ������Զ����ҹ��ܣ��Բ��ҵ����ļ�����������һ�㽨��ѡ��DGSS\...\������ͼ\Ұ����ͼ\��Ŀ¼Ϊ����Ŀ¼������ѡ��Ŀ¼���Զ������ļ������ļ�����ɸ��,ͨ����Ҫ�޸�Gpoint��Routing��Boutary��Attitude��Sample��·��.db�ļ��������©����ֶ���ӣ��Բ���Ҫ�����ļ��ɰ�del���������޳�ѡ���ļ�������ɾ�����������Ҫ���ҵ��ֶκ���Ҫ�滻Ϊ���ֶν��в����滻�����֮��DGSS����Ƿ�����滻Ҫ��"
Tixing.Caption = "��Ҫ���ѣ�" & "1�����ݿ����ǰ��ر������ݣ�" & "2���ó���ֻ�ɸ��ĵ��ʵ㡢·�ߡ����ʽ��ߡ���״����Ʒ����ͼ��λ�����������ݣ�һ��������Ұ��ʵ�����Ϊ׼��������ֻ�Ǹ������ģ�" & "3��������������������ݿ����������������˳е���" & "4����Ȩ��@brilliantfeat���У�����������ҵ��;��"
End Sub

