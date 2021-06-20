VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "桂林理工大学校园网上网"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5475
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   5475
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Caption         =   "开机自启动"
      Height          =   495
      Left            =   600
      TabIndex        =   13
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "关于软件"
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Timer Reconnect 
      Interval        =   1000
      Left            =   4200
      Top             =   3000
   End
   Begin VB.CommandButton Command3 
      Caption         =   "设置运营商"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出软件"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "信息框："
      Height          =   2175
      Left            =   3240
      TabIndex        =   6
      Top             =   480
      Width           =   2175
      Begin VB.Label info 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "重连次数：点击查看"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   1620
      End
      Begin VB.Label info 
         AutoSize        =   -1  'True
         Caption         =   "运营商：等待网络"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1440
      End
      Begin VB.Label info 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用户：加载中"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label info 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "月流量：0 MB"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   1080
      End
      Begin VB.Label info 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "月时长：0 分"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1080
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "自动重连"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "注销登录"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label errorTip 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   4200
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "校园网上网状态："
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
If Check1.Value = 0 Then
    If ReadReg("", "auto") = "1" Then
        MsgBox "已取消打开时自动登录！"
        WriteReg "", "auto", "0"
    End If
End If
End Sub

Private Sub Command1_Click()
main.logout
End Sub

Private Sub Command2_Click()
Dim Msg
Msg = MsgBox("确认退出吗？退出后可能将无法上网！", vbOKCancel Or vbExclamation, "确认退出吗？")
If Msg = vbCancel Then Exit Sub
tray.untray
End
End Sub

Private Sub Command3_Click()
chooseNetwork.Show
End Sub

Private Sub Command4_Click()
main.about
End Sub

Private Sub Command5_Click()
Dim Msg
Dim path As String
path = apppath & "\" & App.ExeName & ".exe autorun"
Msg = MsgBox("确定要设置开机自启动吗：" & path & vbCrLf & "可在任务管理器内随时禁用！", vbYesNo)
If Msg <> vbYes Then Exit Sub
On Error GoTo err:
Dim res
res = autorun("glut-network", path)
If res Then
    MsgBox "开机自启动设置成功，下次开机将会自动启动，记得勾选【打开时自动登录】哦！"
Else
    MsgBox "设置失败，请使用管理员权限打开，或杀毒软件拦截！"
End If
Exit Sub
err:
MsgBox "设置开机自启动时出错了！" & vbCrLf & err.Description & vbCrLf & "请使用管理员权限打开，或被杀毒软件拦截！", vbCritical
End Sub

Private Sub Form_Load()
Me.Caption = "桂林理工大学校园网上网 " & version & " ——CH"
Shape1.BackColor = vbGreen
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Hide
tray.m_NotifyIcon.ShowBubble "校园网上网", "当你点击关闭按钮时，将自动为您缩小到托盘！"
Cancel = -1
End Sub
Private Sub info_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 4 Then
        info(4).Caption = "重连次数：" & main.connect_n & "次"
    End If
End Sub

Private Sub info_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 4 Then
        info(4).Caption = "重连次数：点击查看"
    End If
End Sub

Private Sub Reconnect_Timer()
'自动重连
If main.check_login Then
    Shape1.BackColor = vbGreen
Else
    errorTip.Caption = main.Error
    Shape1.BackColor = vbYellow
    If Check1.Value = 1 Then
        errorTip.Caption = "正在重连..."
        If main.login(main.local_zh, main.local_mm, Val(ReadReg("", "type"))) Then
            main.connect_n = main.connect_n + 1
            Shape1.BackColor = vbGreen
            errorTip.Caption = ""
        Else
            Shape1.BackColor = vbRed
            errorTip.Caption = main.Error
        End If
    End If
End If
refresh_info
End Sub
