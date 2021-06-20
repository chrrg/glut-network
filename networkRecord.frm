VERSION 5.00
Begin VB.Form networkRecord 
   Caption         =   "上网详情"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5085
   LinkTopic       =   "Form3"
   ScaleHeight     =   3975
   ScaleWidth      =   5085
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Label2 
      Caption         =   "已登录设备：0台"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "校园网详情："
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
End
Attribute VB_Name = "networkRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

