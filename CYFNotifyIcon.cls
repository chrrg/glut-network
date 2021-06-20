VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CYFNotifyIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================================================================
'文 件 名: CYFNotifyIcon.cls
'当前版本: 1.0
'注    释: 逸风托盘图标类
'公    司: 逸风软件工作室(原随风工作室)
'开发人员: 随风の追梦
'创建时间: 2013-8-12 22:20
'修改时间: ----
'完成时间: 2013-8-12 22:45
'邮    箱: 497759649@QQ.COM
'下    载：http://www.NewXing.com
'版    本: 1.0
'开发时间: 2013-8-12 22:20 ~ 2013-8-12 22:45
'开发人员: 随风の追梦
'描    述: ----
'======================================================================================
Option Explicit
Private Type NOTIFYICONDATA
        cbSize As Long
        hWnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 128
        dwState As Long
        dwStateMask As Long
        szInfo As String * 256
        uTimeout As Long
        szInfoTitle As String * 64
        dwInfoFlags As Long
End Type
Public Enum ICONINFOFLAGS
       NIIF_NONE = &H0
       NIIF_INFO = &H1
       NIIF_WARNING = &H2
       NIIF_ERROR = &H3
       NIIF_GUID = &H4
       NIIF_ICON_MASK = &HF
       NIIF_NOSOUND = &H10
End Enum
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_INFO = &H10
Private Const NIF_MESSAGE = &H1
Private Const WM_MOUSEMOVE = &H200
Private m_stNotifyIcon As NOTIFYICONDATA
'--------------------------------------------------------------------------------------
'函 数 名: AddNotifyIcon
'描    述: 添加托盘图标
'--------------------------------------------------------------------------------------
Public Sub AddNotifyIcon(ByVal hWnd As Long, ByVal hIcon As Long, ByVal strTip As String)
       With m_stNotifyIcon
            .cbSize = Len(m_stNotifyIcon)
            .hIcon = hIcon
            .hWnd = hWnd
            .uID = 1000
            .uCallbackMessage = WM_MOUSEMOVE
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE Or NIF_INFO
            .szTip = strTip & vbNullChar
       End With
       Call Shell_NotifyIcon(NIM_ADD, m_stNotifyIcon)
End Sub
'--------------------------------------------------------------------------------------
'函 数 名: ModifyNotifyIcon
'描    述: 修改托盘图标
'--------------------------------------------------------------------------------------
Public Sub ModifyNotifyIcon(ByVal hWnd As Long, ByVal hIcon As Long, ByVal strTip As String)
       With m_stNotifyIcon
            .cbSize = Len(m_stNotifyIcon)
            .hIcon = hIcon
            .hWnd = hWnd
            .uID = 1000
            .uCallbackMessage = WM_MOUSEMOVE
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE Or NIF_INFO
            .szTip = strTip & vbNullChar
       End With
       Call Shell_NotifyIcon(NIM_MODIFY, m_stNotifyIcon)
End Sub
'--------------------------------------------------------------------------------------
'函 数 名: ShowBubble
'描    述: 显示气泡提示
'--------------------------------------------------------------------------------------
Public Sub ShowBubble(ByVal strTitle As String, ByVal strInfo As String, Optional ByVal eInfoFlags As ICONINFOFLAGS = NIIF_INFO)
       With m_stNotifyIcon
            .szInfoTitle = strTitle & vbNullChar
            .szInfo = strInfo & vbNullChar
            .dwInfoFlags = eInfoFlags
       End With
       Call Shell_NotifyIcon(NIM_MODIFY, m_stNotifyIcon)
End Sub
'--------------------------------------------------------------------------------------
'函 数 名: HideBubble
'描    述: 隐藏气泡提示
'--------------------------------------------------------------------------------------
Public Sub HideBubble()
       With m_stNotifyIcon
            .szInfoTitle = vbNullChar
            .szInfo = vbNullChar
       End With
       Call Shell_NotifyIcon(NIM_MODIFY, m_stNotifyIcon)
End Sub
'--------------------------------------------------------------------------------------
'函 数 名: DeleteNotifyIcon
'描    述: 删除托盘图标
'--------------------------------------------------------------------------------------
Public Sub DeleteNotifyIcon()
       Call Shell_NotifyIcon(NIM_DELETE, m_stNotifyIcon)
End Sub
'--------------------------------------------------------------------------------------
'事 件 名: Class_Terminate
'描    述: ----
'--------------------------------------------------------------------------------------
Private Sub Class_Terminate()
        Call DeleteNotifyIcon
End Sub
