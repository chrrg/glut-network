Attribute VB_Name = "reg"

Public Const ExeName = "CHglutweb3" '注册表数据保存到32位： HKEY_LOCAL_MACHINE\SOFTWARE\你的应用程序名称  64位：HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\你的应用程序名称

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Const REG_SZ As Long = 1
Const REG_DWORD As Long = 4
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_CURRENT_USER = &H80000001
Const KEY_ALL_ACCESS = &H3F
Const KEY_WOW64_64KEY = &H100 '访问的是64位的注册表。
Const KEY_WOW64_32KEY = &H200 '访问的是32位的注册表。
Private Function QueryValue(ByVal lPredefinedKey As Long, ByVal sKeyName As String, ByVal sValueName As String)
       Dim lRetVal As Long
       Dim hKey As Long
       Dim vValue As Variant
       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey) 'KEY_ALL_ACCESS
       lRetVal = QueryValueEx(hKey, sValueName, vValue)
       QueryValue = vValue
       RegCloseKey (hKey)
End Function
Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, ByRef vValue As Variant) As Long
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String
    On Error GoTo QueryValueExError
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> 0 Then Error 5
    Select Case lType
        Case REG_SZ:
            sValue = String(cch, 0)
            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
            If lrc = 0 Then
                vValue = Left$(sValue, cch)
            Else
                vValue = Empty
            End If
        Case REG_DWORD:
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
            If lrc = 0 Then vValue = lValue
        Case Else
            lrc = -1
    End Select
QueryValueExExit:
    QueryValueEx = lrc
    Exit Function
QueryValueExError:
    Resume QueryValueExExit
End Function

Public Function ReadReg(ByVal Program As String, ByVal Value As String) As String
    If Program = "" Then Program = ExeName
    ReadReg = GetSetting("CHSoftware", Program, Value, "")
End Function
Public Sub WriteReg(ByVal Program As String, ByVal Value As String, ByVal str As String)
    If Program = "" Then Program = ExeName
    SaveSetting "CHSoftware", Program, Value, str
End Sub

Public Function autorun(ByVal Program As String, ByVal url As String) As Boolean
If QueryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", Program) <> url Then
    Set W = CreateObject("wscript.shell")
    W.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\" & Program, url
    Set W = Nothing
    autorun = Replace(QueryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", Program), Chr(0), "") = url
    Exit Function
End If
autorun = True
End Function
