VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "桂林理工大学-校园网上网客户端3.0"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6555
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   6555
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6000
      Top             =   4200
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1680
      TabIndex        =   15
      Top             =   2640
      Width           =   3135
      Begin VB.OptionButton Option2 
         Caption         =   "屏风校区"
         Height          =   495
         Index           =   1
         Left            =   2040
         TabIndex        =   17
         Top             =   0
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "雁山校区"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "联通"
      Height          =   255
      Index           =   3
      Left            =   4680
      TabIndex        =   7
      Top             =   2400
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "移动"
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   6
      Top             =   2400
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "电信"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "校园网"
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   4
      Top             =   2400
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      Caption         =   "打开时自动登录"
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "关于"
      Height          =   615
      Left            =   3360
      TabIndex        =   3
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "记住账号密码"
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1800
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "登录"
      Height          =   615
      Left            =   1680
      TabIndex        =   2
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   1200
      Width           =   4215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "如果未绑定运营商，请选择校园网"
      ForeColor       =   &H00808000&
      Height          =   180
      Left            =   1920
      TabIndex        =   14
      Top             =   3600
      Width           =   3180
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "桂林理工大学-校园网登录"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   12
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "校园网密码："
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "校园网账号："
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "制作：桂林理工大学-CH"
      Height          =   180
      Left            =   2400
      TabIndex        =   13
      Top             =   4680
      Width           =   1890
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim needUnload As Boolean

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        Check1.Value = 1
    End If
End Sub

Private Sub Command1_Click()
Randomize
If main.login(Text1.text, Text2.text, Val(ReadReg("", "type"))) Then '登录成功的话
    If Check1.Value = 1 Then
        WriteReg "", "zh", Text1.text
        WriteReg "", "mm", Text2.text
    End If
    If Check2.Value = 1 Then
        WriteReg "", "auto", "1"
    Else
        WriteReg "", "auto", "0"
    End If
    tray.show
    tray.Visible = False
    needUnload = True
    Unload Me
Else
    If ReadReg("", "auto") = "1" Then
        Label5.Caption = "自动登录：" & errProxy(main.Error) & "|" & Int(Rnd * 10000)
        Timer1.Interval = 2000
        Timer1.Enabled = True
    End If
    If main.Error = "error5 waitsec <3" Then
        Label5.Caption = "请3秒之后再试！" & Int(Rnd * 100)
    ElseIf main.Error = "bind userid error" Then
        MsgBox "您还未绑定此运营商的账号密码！请先选中校园网登录，进入后设置运营商账号密码即可！", vbInformation, "校园网登录时出错！"
    ElseIf main.Error = "userid error1" Then
        MsgBox "此校园网账号不存在！", vbInformation, "校园网登录时出错！"
    ElseIf main.Error = "userid error2" Then
        MsgBox "密码错误，请重新输入密码！", vbInformation, "校园网登录时出错！"
    ElseIf main.Error = "Oppp error: can't find user." Then
        MsgBox "您选择的运营商还未绑定账号，需先选择校园网登录后再绑定运营商账号！", vbInformation, "校园网登录时出错！"
    Else
        If Val(ReadReg("", "type")) = 0 Then
            MsgBox main.Error, vbInformation, errProxy(main.Error) 'errProxy(main.Error) & Int(Rnd * 100)
        Else
            MsgBox main.Error, vbInformation, errProxy(main.Error)
        End If
    End If
End If

End Sub

Private Sub Command2_Click()
main.about
End Sub


Private Sub Form_Load()
Dim i As Long
WriteReg "", "test", "1"
If ReadReg("", "test") <> "1" Then
    MsgBox "无法设置注册表！请右键使用管理员运行！", vbCritical, "检测"
End If

If App.PrevInstance Then
    'tray.m_NotifyIcon.ShowBubble "多开提醒", "为保障程序运行稳定，请勿运行第二个实例，谢谢！"
    MsgBox "您双开了！如果您觉得没有双开，可以使用任务管理器结束另一个进程！", vbInformation
    End
End If

version = App.Major & "." & App.Minor & "." & App.Revision
Me.Caption = "桂林理工大学-校园网上网 " & version & " ——CH"

If Len(Command) >= 4 Then
    Dim path As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Command = "update" Or Left(Command, 4) = "upin" Then
        On Error Resume Next
        DelayM 100
        Randomize
        If Left(Command, 4) = "upin" Then
            path = App.path & "\" & Mid(Command, 5) & ".exe"
            If Dir(path, vbNormal) <> "" Then '删除旧版本
                Do
                    i = i + 1
                    fso.DeleteFile path, True
                    If Dir(path, vbNormal) = "" Then Exit Do '删除成功退出循环
                    DelayM 100
                Loop Until i > 100
            End If
        End If
        Dim file1 As String, file2 As String
        file1 = App.path & "\" & App.ExeName & ".exe" '随机文件名
        file2 = App.path & "\桂工校园网登录.exe" '规范化文件名
        If Dir(file1, vbNormal) = "" Then
            MsgBox "更新出错！找不到" & file1 & "但新版本已经为你下载好，但程序找不到文件：" & App.path & "\" & App.ExeName & ".exe" & vbCrLf & "请确保没放在盘根目录！"
            End
        End If
        
        fso.CopyFile App.path & "\" & App.ExeName & ".exe", file2
        Set fso = Nothing
        Shell file2 & " upok" & App.ExeName, vbNormalFocus
        End
    End If
    If Left(Command, 4) = "upok" Then
        On Error Resume Next
        DelayM 100
        path = App.path & "\" & Mid(Command, 5) & ".exe"
        If Dir(path, vbNormal) <> "" Then
            Do
                i = i + 1
                fso.DeleteFile path, True
                If Dir(path, vbNormal) = "" Then Exit Do
                DelayM 100
            Loop Until i > 100
        End If
        If ReadReg("", "auto") <> "1" Then
            MsgBox "软件已成功自动更新！此程序路径：" & path & "" & vbCrLf & "点击确认启动！"
        End If
    End If
    If Command = "autorun" Then
        WriteReg "", "auto", "1"
        '自动启动肯定是要自动登录
    End If
End If
Set objScrCtl = CreateObject("MSScriptControl.ScriptControl") '初始化js解释器
objScrCtl.Language = "JavaScript"

Text1.text = ReadReg("", "zh")
Text2.text = ReadReg("", "mm")
Check1.Value = IIf(Text2.text <> "", 1, 0)
Check2.Value = IIf(ReadReg("", "auto") = "1", 1, 0)

Dim xiaoqu As Long
If ReadReg("", "xiaoqu") = "" Then
    xiaoqu = 0
Else
    xiaoqu = Val(ReadReg("", "xiaoqu"))
End If


Dim types As Long
If ReadReg("", "type") = "" Then
    types = 0
Else
    types = Val(ReadReg("", "type"))
End If

Option1(types).Value = True
Option2(xiaoqu).Value = True

'MsgBox md5.md5("1", 32)
Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not needUnload Then End
End Sub

Private Sub Option1_Click(Index As Integer)
    If Option1(Index).Value Then
        WriteReg "", "type", Index
    End If
End Sub
Private Sub Option2_Click(Index As Integer)
    
    If Option2(Index).Value Then
        If Index = 1 Then
            If Option1(2).Value = True Then Option1(0).Value = True
            Option1(2).Enabled = False
        Else
            Option1(2).Enabled = True
        End If
        WriteReg "", "xiaoqu", Index
    End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1_Click
End If
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    If Timer1.Interval <> 2000 Then
        If ReadReg("", "auto") <> "1" Then '没有自动登录，需要检查更新
            main.checkUpdate
        Else
            Do
                If main.checkUpdate() = 0 Then Exit Do
                DelayM 2000
            Loop Until 0
        End If
    End If
    
    
    
    'MsgBox types
    If ReadReg("", "auto") = "1" Then
        Label5.Caption = "自动登录中..." & Rnd
        Command1_Click
    Else
        'If main.check_login Then
            'MsgBox "校园网已登录，请设置您的校园网账号密码，下次打开可快速登录。"
        'End If
    End If
End Sub
