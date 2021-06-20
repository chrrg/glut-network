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
'�� �� ��: CYFNotifyIcon.cls
'��ǰ�汾: 1.0
'ע    ��: �ݷ�����ͼ����
'��    ˾: �ݷ�����������(ԭ��繤����)
'������Ա: ����׷��
'����ʱ��: 2013-8-12 22:20
'�޸�ʱ��: ----
'���ʱ��: 2013-8-12 22:45
'��    ��: 497759649@QQ.COM
'��    �أ�http://www.NewXing.com
'��    ��: 1.0
'����ʱ��: 2013-8-12 22:20 ~ 2013-8-12 22:45
'������Ա: ����׷��
'��    ��: ----
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
'�� �� ��: AddNotifyIcon
'��    ��: ��������ͼ��
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
'�� �� ��: ModifyNotifyIcon
'��    ��: �޸�����ͼ��
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
'�� �� ��: ShowBubble
'��    ��: ��ʾ������ʾ
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
'�� �� ��: HideBubble
'��    ��: ����������ʾ
'--------------------------------------------------------------------------------------
Public Sub HideBubble()
       With m_stNotifyIcon
            .szInfoTitle = vbNullChar
            .szInfo = vbNullChar
       End With
       Call Shell_NotifyIcon(NIM_MODIFY, m_stNotifyIcon)
End Sub
'--------------------------------------------------------------------------------------
'�� �� ��: DeleteNotifyIcon
'��    ��: ɾ������ͼ��
'--------------------------------------------------------------------------------------
Public Sub DeleteNotifyIcon()
       Call Shell_NotifyIcon(NIM_DELETE, m_stNotifyIcon)
End Sub
'--------------------------------------------------------------------------------------
'�� �� ��: Class_Terminate
'��    ��: ----
'--------------------------------------------------------------------------------------
Private Sub Class_Terminate()
        Call DeleteNotifyIcon
End Sub