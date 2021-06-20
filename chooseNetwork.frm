VERSION 5.00
Begin VB.Form chooseNetwork 
   Caption         =   "校园网设置"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8820
   Icon            =   "chooseNetwork.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   5940
   ScaleWidth      =   8820
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6960
      Top             =   4680
   End
   Begin VB.Frame Frame1 
      Caption         =   "运营商设置"
      Enabled         =   0   'False
      Height          =   5055
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   5175
      Begin VB.OptionButton Option1 
         Caption         =   "电信"
         Height          =   615
         Index           =   1
         Left            =   3840
         TabIndex        =   7
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox zh 
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   1
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox mm 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox zh 
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   3
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox mm 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox zh 
         Height          =   375
         Index           =   2
         Left            =   1080
         TabIndex        =   5
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox mm 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   3240
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "确认保存"
         Height          =   495
         Left            =   1440
         TabIndex        =   12
         Top             =   4440
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "显示密码"
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Top             =   3840
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "移动"
         Height          =   615
         Index           =   2
         Left            =   3840
         TabIndex        =   8
         Top             =   1920
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "联通"
         Height          =   615
         Index           =   3
         Left            =   3840
         TabIndex        =   9
         Top             =   2880
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "校园网"
         Height          =   615
         Index           =   0
         Left            =   3840
         TabIndex        =   10
         Top             =   3840
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "账号："
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   19
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "密码："
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   18
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "账号："
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   17
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "密码："
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   16
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "账号："
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   15
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "密码："
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   14
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "请选择一个运营商："
         Height          =   375
         Left            =   3240
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Label Label4 
      Caption         =   "详情："
      Height          =   4815
      Left            =   6240
      TabIndex        =   20
      Top             =   600
      Width           =   2175
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "chooseNetwork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
If Option1(1).Value Then
WriteReg "", "type", "1"
ElseIf Option1(2).Value Then
WriteReg "", "type", "2"
ElseIf Option1(3).Value Then
WriteReg "", "type", "3"
ElseIf Option1(0).Value Then
WriteReg "", "type", "0"
End If
If main.bind(zh(0).text, mm(0).text, zh(1).text, mm(1).text, zh(2).text, mm(2).text) Then
    MsgBox main.Error
    main.logout
Else
    MsgBox main.Error
End If
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
mm(0).PasswordChar = ""
mm(1).PasswordChar = ""
mm(2).PasswordChar = ""
End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mm(0).PasswordChar = "*"
mm(1).PasswordChar = "*"
mm(2).PasswordChar = "*"

End Sub

Private Sub Form_Load()
Option1(Val(ReadReg("", "type"))).Value = True
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Label4.Caption = "正在登录校园网系统，请耐心等待几秒..."
DoEvents
main.login_uss
Label4.Caption = "登录成功，正在取信息..."
DoEvents
main.getbind
zh(0).text = main.bindinfo.data1
mm(0).text = main.bindinfo.data2
zh(1).text = main.bindinfo.data3
mm(1).text = main.bindinfo.data4
zh(2).text = main.bindinfo.data5
mm(2).text = main.bindinfo.data6
DoEvents
Label4.Caption = "登录成功！" & vbCrLf & _
"套餐：" & main.userinfo.userGroupName & vbCrLf & _
"下载流量：" & Round(main.userinfo.internetDownFlow, 3) & " MB" & vbCrLf & _
"上传流量：" & Round(main.userinfo.internetUpFlow, 3) & " MB" & vbCrLf

Frame1.Enabled = True
End Sub
