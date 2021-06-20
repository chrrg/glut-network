Attribute VB_Name = "main"
Option Explicit
Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Integer, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Integer, ByVal lpfnCB As Integer) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public domain As String
Private Type userinfoType
    userRealName As String
    userGroupName As String
    internetDownFlow As String
    internetUpFlow As String
    Time As String
    flow As String
End Type
Private Type bindType
    csrf As String
    data1 As String
    data2 As String
    data3 As String
    data4 As String
    data5 As String
    data6 As String
End Type
Public userinfo As userinfoType
Public bindinfo As bindType
Public Error As String
Public version As String
Dim cookie As String
Dim checkcode As String
Dim csrf As String
Public local_zh As String
Public local_mm As String
Public objScrCtl As Object
Public connect_n As Long
Function errProxy(s As String)
Dim text As String
text = "|SESSION已过期,请重新登录|no errcode|AC认证失败|Authentication Fail ErrCode=04|上网时长/流量已到上限|Authentication Fail ErrCode=05|您的账号已停机，造成停机的可能原因： 1、用户欠费停机 2、用户报停 需要了解具体原因，请访问自助服务系统。|Authentication Fail ErrCode=09|本账号费用超支，禁止使用|Authentication Fail ErrCode=11|不允许Radius登录|Authentication Fail ErrCode=80|接入服务器不存在|Authentication Fail ErrCode=81|LDAP认证失败|Authentication Fail ErrCode=85|账号正在使用|Authentication Fail ErrCode=86|绑定IP或MAC失败|Authentication Fail ErrCode=88|IP地址冲突|Authentication Fail ErrCode=94|接入服务器并发超限|err(2)|请在指定的登录源地址范围内登录|err(3)|请在指定的IP登录|err(7)|请在指定的登录源VLAN范围登录|err(10)|请在指定的Vlan登录|err(11)|请在指定的MAC登录|err(17)|请在指定的设备端口登录|userid error1|账号不存在|userid error2|密码错误|userid error3|密码错误|auth error4|用户使用量超出限制|auth error5|账号已停机|auth error9|时长流量超支|auth error80|本时段禁止上网|auth error99|用户名或密码错误|" & _
"auth err198|用户名或密码错误|auth error199|用户名或密码错误|auth error258|账号只能在指定区域使用|auth error|用户验证失败|set_onlinet error|用户数超过限制|In use|登录超过人数限制|port err|上课时间不允许上网|can not use static ip|不允许使用静态IP|[01], 本帐号只能在指定VLANID使用(0.4095)|本帐号只能在指定VLANID使用|Mac, IP, NASip, PORT err(6)!|本帐号只能在指定VLANID使用|wuxian OLno|VLAN范围控制账号的接入数量超出限制|Oppp error: 1|运营商账号密码错误，错误码为：1|Oppp error: 5|运营商账号在线，错误码为：5|Oppp error: 18|运营商账号密码错误，错误码为：18|Oppp error: 21|运营商账号在线，错误码为：21|Oppp error: 26|运营商账号被绑定，错误码为：26|Oppp error: 29|运营商账号锁定的用户端口NAS-Port-Id错误，错误码为：29|Oppp error: userid inuse|运营商账号已被使用|Oppp error: can't find user|运营商账号无法获取或不存在|bind userid error|绑定运营商账号失败|Oppp error: TOO MANY CONNECTIONS|运营商账号在线|Oppp error: Timeout|运营商账号状态异常(欠费等)|Oppp error: User dial-in so soon|运营商账号刚下线|Oppp error: " & _
"SERVICE SUSPENDED|欠费暂停服务|Oppp error: open vpn session fail!|运营商账号已欠费,请充值|Oppp error: INVALID LOCATION.|运营商锁定的用户端口错误|Oppp error: 99|帐号绑定域名错误，请联系运营商检查或解绑。|error5 waitsec <3|登录过于频繁，请等候重新登录。"
Dim arr
arr = Split(text, "|")
Dim l As Long, i As Long
l = (UBound(arr) + 1) / 2
For i = 0 To l - 1
    If s = arr(i * 2) Then
        errProxy = arr(i * 2 + 1)
        Exit Function
    End If
Next
errProxy = "登录异常"
End Function
Function getUserInfo() As Boolean
On Error Resume Next:
If cookie = "" Then
    Exit Function
End If
Dim WinHttp
Set WinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
'设置参数
WinHttp.SetTimeouts 60000, 60000, 60000, 5000 '设置操作超时时间
WinHttp.Option(4) = 13056 '忽略错误标志
WinHttp.Option(6) = True '为 True 时，当请求页面重定向跳转时自动跳转，False 不自动跳转，截取服务端返回的302状态。
WinHttp.Open "GET", "http://uss.glut.edu.cn/Self/dashboard", True 'GET 或 POST, Url, False 同步方式；True 异步方式
WinHttp.SetRequestHeader "Host", "uss.glut.edu.cn"
WinHttp.SetRequestHeader "Connection", "keep-alive"
WinHttp.SetRequestHeader "Cookie", "JSESSIONID=" & cookie
WinHttp.SetRequestHeader "DNT", "1"
WinHttp.SetRequestHeader "Cache-Control", "max-age=0"
WinHttp.SetRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3"
WinHttp.SetRequestHeader "Origin", "http://uss.glut.edu.cn" '接受数据类型
WinHttp.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2987.98 Safari/537.36 LBBROWSER" '用户浏览器信息
WinHttp.SetRequestHeader "Accept-Language", "zh-CN,zh;q=0.8"
WinHttp.SetRequestHeader "Upgrade-Insecure-Requests", "1"
WinHttp.send
WinHttp.WaitForResponse
Dim result As String, userRealName As String
result = WinHttp.ResponseText
If InStr(1, result, "userRealName"":""") Then
    userinfo.userRealName = Split(Split(result, "userRealName"":""")(1), """")(0)
Else
    Exit Function
End If
If InStr(1, result, "userGroupName"":""") Then
    userinfo.userGroupName = Split(Split(result, "userGroupName"":""")(1), """")(0)
Else
    Exit Function
End If
If InStr(1, result, "internetDownFlow"":") Then
    userinfo.internetDownFlow = Trim(Split(Split(result, "internetDownFlow"":")(1), ",")(0))
Else
    Exit Function
End If

If InStr(1, result, "internetUpFlow"":") Then
    userinfo.internetUpFlow = Trim(Split(Split(result, "internetUpFlow"":")(1), ",")(0))
Else
    Exit Function
End If

getUserInfo = True
End Function

Function login(zh As String, mm As String, Optional loginType = 0) As Boolean
On Error Resume Next:
If zh = "" Or mm = "" Then
    Exit Function
End If
If Val(ReadReg("", "xiaoqu")) = 0 Then
    main.domain = "172.16.2.2"
Else
    main.domain = "202.193.80.124"
End If
DoEvents
Dim WinHttp
Set WinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
'设置参数
WinHttp.SetTimeouts 60000, 60000, 60000, 5000 '设置操作超时时间
WinHttp.Option(4) = 13056 '忽略错误标志
WinHttp.Option(6) = True '为 True 时，当请求页面重定向跳转时自动跳转，False 不自动跳转，截取服务端返回的302状态。
WinHttp.Open "GET", "http://" & main.domain & "/drcom/login?callback=dr1004&DDDDD=" & zh & "&upass=" & mm & "&0MKKey=123456&R1=0&R3=" & loginType & "&R6=0&para=00&v6ip=&v=8239", True 'GET 或 POST, Url, False 同步方式；True 异步方式
WinHttp.SetRequestHeader "Host", main.domain
WinHttp.SetRequestHeader "Connection", "keep-alive"
WinHttp.SetRequestHeader "DNT", "1"
WinHttp.SetRequestHeader "Cache-Control", "max-age=0"
WinHttp.SetRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3"
WinHttp.SetRequestHeader "Origin", "http://" & main.domain '接受数据类型
WinHttp.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2987.98 Safari/537.36 LBBROWSER" '用户浏览器信息
WinHttp.SetRequestHeader "Accept-Language", "zh-CN,zh;q=0.8"
WinHttp.SetRequestHeader "Upgrade-Insecure-Requests", "1"
WinHttp.send
WinHttp.WaitForResponse
Dim result As String
result = BytesToBstr(WinHttp.ResponseBody, "GB2312")
If result = "" Then
Exit Function
End If
If InStr(1, result, "dr1004({") Then
    Dim js As String
    js = "{" & Split(Split(result, "dr1004({")(1), "})")(0) & "}"
    'MsgBox js
    If objScrCtl.Eval("(" & js & ").result") = "1" Then '登录成功！
        local_zh = zh
        local_mm = mm
        login = True
    Else
        Error = objScrCtl.Eval("(" & js & ").msga")
    End If
End If

End Function
Function check_login()
On Error Resume Next:
Dim WinHttp
Set WinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
'设置参数
WinHttp.SetTimeouts 60000, 60000, 60000, 5000 '设置操作超时时间
WinHttp.Option(4) = 13056 '忽略错误标志
WinHttp.Option(6) = True '为 True 时，当请求页面重定向跳转时自动跳转，False 不自动跳转，截取服务端返回的302状态。
WinHttp.Open "GET", "http://" & main.domain, True 'GET 或 POST, Url, False 同步方式；True 异步方式
WinHttp.SetRequestHeader "Host", main.domain
WinHttp.SetRequestHeader "Connection", "keep-alive"
WinHttp.SetRequestHeader "DNT", "1"
WinHttp.SetRequestHeader "Cache-Control", "max-age=0"
WinHttp.SetRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3"
WinHttp.SetRequestHeader "Origin", "http://" & main.domain '接受数据类型
WinHttp.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2987.98 Safari/537.36 LBBROWSER" '用户浏览器信息
WinHttp.SetRequestHeader "Accept-Language", "zh-CN,zh;q=0.8"
WinHttp.SetRequestHeader "Upgrade-Insecure-Requests", "1"
WinHttp.send
WinHttp.WaitForResponse

Dim result As String
result = BytesToBstr(WinHttp.ResponseBody, "GB2312")
'result = WinHttp.ResponseText
If InStr(1, result, "NID='") Then
    
    userinfo.userRealName = Trim(Split(Split(result, "NID='")(1), "'")(0))
    userinfo.Time = Trim(Split(Split(result, "time='")(1), "'")(0))
    userinfo.flow = Round(Trim(Split(Split(result, "flow='")(1), "'")(0)) / 1024, 3)
    check_login = True
    
End If

End Function
Function get_yys() As String
    get_yys = ReadReg("", "type")
    If get_yys = "" Then
        get_yys = "0"
    End If
End Function
Function getYYS(yysType As String) As String
If yysType = "" Or yysType = "0" Then
    getYYS = "校园网"
    Exit Function
ElseIf yysType = "1" Then
    getYYS = "电信"
    Exit Function
ElseIf yysType = "2" Then
    getYYS = "移动"
    Exit Function
ElseIf yysType = "3" Then
    getYYS = "联通"
    Exit Function
End If
End Function
Function refresh_info()
Form2.info(0).Caption = "运营商：" & getYYS(ReadReg("", "type"))
Form2.info(1).Caption = "用户：" & userinfo.userRealName
Form2.info(2).Caption = "月时长：" & userinfo.Time & " 分"
Form2.info(3).Caption = "月流量：" & userinfo.flow & " MB"
'Form2.info(4).Caption = "设备：" & 0 & " 台"
End Function
Function login_uss() As Boolean '用uss登录校园网
On Error Resume Next:
    Dim zh As String, mm As String
    zh = local_zh
    mm = local_mm
    If Not getcookie() Then
        Error = "校园网暂时不稳定或失效，请稍后重试！"
        Exit Function
    End If
    Dim DAT As String
    DAT = "foo=&bar=&checkcode=" & checkcode & "&account=" & zh & "&password=" & md5.md5(mm, 32) & "&code="
    'MsgBox DAT
    
    Dim WinHttp
    
    Set WinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    '设置参数
    WinHttp.SetTimeouts 60000, 60000, 60000, 5000 '设置操作超时时间
    WinHttp.Option(4) = 13056 '忽略错误标志
    WinHttp.Option(6) = True '为 True 时，当请求页面重定向跳转时自动跳转，False 不自动跳转，截取服务端返回的302状态。
    WinHttp.Open "GET", "http://uss.glut.edu.cn/Self/login/randomCode", True 'GET 或 POST, Url, False 同步方式；True 异步方式
    WinHttp.SetRequestHeader "Host", "uss.glut.edu.cn"
    WinHttp.SetRequestHeader "Connection", "keep-alive"
    WinHttp.SetRequestHeader "Cookie", "JSESSIONID=" & cookie
    WinHttp.SetRequestHeader "DNT", "1"
    WinHttp.SetRequestHeader "Cache-Control", "max-age=0"
    WinHttp.SetRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3"
    WinHttp.SetRequestHeader "Origin", "http://uss.glut.edu.cn" '接受数据类型
    WinHttp.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2987.98 Safari/537.36 LBBROWSER" '用户浏览器信息
    WinHttp.SetRequestHeader "Accept-Language", "zh-CN,zh;q=0.8"
    WinHttp.SetRequestHeader "Upgrade-Insecure-Requests", "1"
    WinHttp.send
    WinHttp.WaitForResponse
    
    Set WinHttp = CreateObject("WinHttp.WinHttpRequest.5.1") '设置参数
    WinHttp.SetTimeouts 60000, 60000, 60000, 5000 '设置操作超时时间
    WinHttp.Option(4) = 13056 '忽略错误标志
    WinHttp.Option(6) = True '为 True 时，当请求页面重定向跳转时自动跳转，False 不自动跳转，截取服务端返回的302状态。
    WinHttp.Open "POST", "http://uss.glut.edu.cn/Self/login/verify", True 'GET 或 POST, Url, False 同步方式；True 异步方式
    
    WinHttp.SetRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3"
    WinHttp.SetRequestHeader "Accept-Language", "zh-CN,zh;q=0.9"
    WinHttp.SetRequestHeader "Cache-Control", "max-age=0"
    WinHttp.SetRequestHeader "Connection", "keep-alive"
    WinHttp.SetRequestHeader "Content-Length", Len(DAT)
    WinHttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    WinHttp.SetRequestHeader "Cookie", "JSESSIONID=" & cookie
    WinHttp.SetRequestHeader "Host", "uss.glut.edu.cn"
    WinHttp.SetRequestHeader "Origin", "http://uss.glut.edu.cn"
    WinHttp.SetRequestHeader "Referer", "http://uss.glut.edu.cn/Self/login/"
    WinHttp.SetRequestHeader "Upgrade-Insecure-Requests", "1"
    WinHttp.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2987.98 Safari/537.36 LBBROWSER" '用户浏览器信息

    WinHttp.send (DAT)
    WinHttp.WaitForResponse
    Dim result
    result = WinHttp.ResponseText

    If InStr(1, result, "csrftoken") Then
        csrf = Trim(Split(Split(result, "csrftoken: '")(1), "'")(0))
        'MsgBox csrf
        login_uss = True
    End If
    If InStr(1, result, "})('") Then
        Error = Split(Split(result, "})('")(1), "'")(0) '登录失败时的提示！
    End If
    getUserInfo
End Function

Function getcookie() As Boolean '得到一个新的cookie
On Error Resume Next:
Dim url As String
url = "http://uss.glut.edu.cn/Self/login/"

Dim WinHttp
Set WinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
'设置参数
WinHttp.SetTimeouts 60000, 60000, 60000, 5000 '设置操作超时时间
WinHttp.Option(4) = 13056 '忽略错误标志
WinHttp.Option(6) = True '为 True 时，当请求页面重定向跳转时自动跳转，False 不自动跳转，截取服务端返回的302状态。
WinHttp.Open "GET", url, True 'GET 或 POST, Url, False 同步方式；True 异步方式
WinHttp.SetRequestHeader "Host", "uss.glut.edu.cn"
WinHttp.SetRequestHeader "Connection", "keep-alive"
WinHttp.SetRequestHeader "DNT", "1"
WinHttp.SetRequestHeader "Cache-Control", "max-age=0"
WinHttp.SetRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3"
WinHttp.SetRequestHeader "Origin", "http://uss.glut.edu.cn" '接受数据类型
WinHttp.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2987.98 Safari/537.36 LBBROWSER" '用户浏览器信息
WinHttp.SetRequestHeader "Accept-Language", "zh-CN,zh;q=0.8"
WinHttp.SetRequestHeader "Upgrade-Insecure-Requests", "1"
WinHttp.send
WinHttp.WaitForResponse
Dim result
result = WinHttp.ResponseText

cookie = WinHttp.getResponseHeader("set-cookie")

If InStr(1, cookie, "JSESSIONID=") Then
    cookie = Trim(Split(Split(cookie, "JSESSIONID=")(1), ";")(0))
Else
    getcookie = False
    Exit Function
End If
If InStr(1, result, "name=""checkcode"" value=""") Then
    checkcode = Split(Split(result, "name=""checkcode"" value=""")(1), """")(0)
Else
    getcookie = False
    Exit Function
End If
getcookie = True
End Function

Function getbind() '得到绑定的运营商
On Error Resume Next:
Dim WinHttp
    
Set WinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
'设置参数
WinHttp.SetTimeouts 60000, 60000, 60000, 5000 '设置操作超时时间
WinHttp.Option(4) = 13056 '忽略错误标志
WinHttp.Option(6) = True '为 True 时，当请求页面重定向跳转时自动跳转，False 不自动跳转，截取服务端返回的302状态。
WinHttp.Open "GET", "http://uss.glut.edu.cn/Self/service/operatorId", True 'GET 或 POST, Url, False 同步方式；True 异步方式
WinHttp.SetRequestHeader "Host", "uss.glut.edu.cn"
WinHttp.SetRequestHeader "Connection", "keep-alive"
WinHttp.SetRequestHeader "Cookie", "JSESSIONID=" & cookie
WinHttp.SetRequestHeader "DNT", "1"
WinHttp.SetRequestHeader "Cache-Control", "max-age=0"
WinHttp.SetRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3"
WinHttp.SetRequestHeader "Origin", "http://uss.glut.edu.cn" '接受数据类型
WinHttp.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2987.98 Safari/537.36 LBBROWSER" '用户浏览器信息
WinHttp.SetRequestHeader "Accept-Language", "zh-CN,zh;q=0.8"
WinHttp.SetRequestHeader "Referer", "http://uss.glut.edu.cn/Self/service"
WinHttp.SetRequestHeader "Upgrade-Insecure-Requests", "1"

WinHttp.send
WinHttp.WaitForResponse
Dim result
result = WinHttp.ResponseText
'Clipboard.Clear
'Clipboard.SetText result
If InStr(1, result, "value=""") Then
Dim valuearr
valuearr = Split(result, "value=""")

csrf = Split(valuearr(1), """")(0)
Dim data1, data2, data3, data4, data5, data6
data1 = Split(valuearr(2), """")(0)
data2 = Split(valuearr(3), """")(0)
data3 = Split(valuearr(4), """")(0)
data4 = Split(valuearr(5), """")(0)
data5 = Split(valuearr(6), """")(0)
data6 = Split(valuearr(7), """")(0)
bindinfo.csrf = csrf
bindinfo.data1 = data1
bindinfo.data2 = data2
bindinfo.data3 = data3
bindinfo.data4 = data4
bindinfo.data5 = data5
bindinfo.data6 = data6
Else
    MsgBox "信息加载失败！"
End If
End Function
Function bind(data1, data2, data3, data4, data5, data6) As Boolean
On Error Resume Next:
If bindinfo.csrf = "" Then
    getbind
End If

Dim DAT As String
DAT = "csrftoken=" & bindinfo.csrf & "&FLDEXTRA1=" & data1 & "&FLDEXTRA2=" & data2 & "&FLDEXTRA3=" & data3 & "&FLDEXTRA4=" & data4 & "&FLDEXTRA5=" & data5 & "&FLDEXTRA6=" & data6

Dim WinHttp As Object
Set WinHttp = CreateObject("WinHttp.WinHttpRequest.5.1") '设置参数
WinHttp.SetTimeouts 60000, 60000, 60000, 5000 '设置操作超时时间
WinHttp.Option(4) = 13056 '忽略错误标志
WinHttp.Option(6) = True '为 True 时，当请求页面重定向跳转时自动跳转，False 不自动跳转，截取服务端返回的302状态。
WinHttp.Open "POST", "http://uss.glut.edu.cn/Self/service/bind-operator", True 'GET 或 POST, Url, False 同步方式；True 异步方式

WinHttp.SetRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3"
WinHttp.SetRequestHeader "Accept-Language", "zh-CN,zh;q=0.9"
WinHttp.SetRequestHeader "Cache-Control", "max-age=0"
WinHttp.SetRequestHeader "Connection", "keep-alive"
WinHttp.SetRequestHeader "Content-Length", Len(DAT)
WinHttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
WinHttp.SetRequestHeader "Cookie", "JSESSIONID=" & cookie
WinHttp.SetRequestHeader "Host", "uss.glut.edu.cn"
WinHttp.SetRequestHeader "Origin", "http://uss.glut.edu.cn"
WinHttp.SetRequestHeader "Referer", "http://uss.glut.edu.cn/Self/service/operatorId"
WinHttp.SetRequestHeader "Upgrade-Insecure-Requests", "1"
WinHttp.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2987.98 Safari/537.36 LBBROWSER" '用户浏览器信息

WinHttp.send (DAT)
WinHttp.WaitForResponse

Dim result
result = WinHttp.ResponseText
'Clipboard.Clear
'Clipboard.SetText result
bindinfo.csrf = ""
If InStr(1, result, "})('") Then
    Dim text
    text = Split(Split(result, "})('")(1), "'")(0) '登录失败时的提示！
    
    If InStr(1, text, "成功") Then
        Error = Trim(Replace(text, "\n", ""))
        bind = True
    Else
        Error = Trim(Replace(text, "\n", "")) '没有输入任何东西的时候
        bind = False
    End If
    If Error = "" Then
        bind = True
        Error = "无需保存！"
    End If
End If
End Function
Function checkUpdateFromHTTP(DAT As String) As String
    On Error Resume Next
    Dim WinHttp
    Set WinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    '设置参数
    WinHttp.SetTimeouts 60000, 60000, 60000, 5000 '设置操作超时时间
    WinHttp.Option(4) = 13056 '忽略错误标志
    WinHttp.Option(6) = True '为 True 时，当请求页面重定向跳转时自动跳转，False 不自动跳转，截取服务端返回的302状态。
    WinHttp.Open "POST", "http://yiban.glut.edu.cn/xyw/update.php", True 'GET 或 POST, Url, False 同步方式；True 异步方式
    WinHttp.SetRequestHeader "Host", "yiban.glut.edu.cn"
    WinHttp.SetRequestHeader "Connection", "keep-alive"
    WinHttp.SetRequestHeader "Content-Length", Len(DAT)
    WinHttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    WinHttp.SetRequestHeader "Cache-Control", "max-age=0"
    WinHttp.SetRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3"
    WinHttp.SetRequestHeader "Origin", "http://yiban.glut.edu.cn" '接受数据类型
    WinHttp.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2987.98 Safari/537.36 LBBROWSER" '用户浏览器信息
    WinHttp.SetRequestHeader "Accept-Language", "zh-CN,zh;q=0.8"
    WinHttp.SetRequestHeader "Upgrade-Insecure-Requests", "1"
    WinHttp.send (DAT)
    WinHttp.WaitForResponse
    checkUpdateFromHTTP = WinHttp.ResponseText
End Function
Function checkUpdateFromHTTPS(DAT As String) As String
    On Error Resume Next
    Dim WinHttp
    Set WinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    '设置参数
    WinHttp.SetTimeouts 60000, 60000, 60000, 5000 '设置操作超时时间
    WinHttp.Option(4) = 13056 '忽略错误标志
    WinHttp.Option(6) = True '为 True 时，当请求页面重定向跳转时自动跳转，False 不自动跳转，截取服务端返回的302状态。
    WinHttp.Open "POST", "https://yiban.glut.edu.cn/xyw/update.php", True 'GET 或 POST, Url, False 同步方式；True 异步方式
    WinHttp.SetRequestHeader "Host", "yiban.glut.edu.cn"
    WinHttp.SetRequestHeader "Connection", "keep-alive"
    WinHttp.SetRequestHeader "Content-Length", Len(DAT)
    WinHttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    WinHttp.SetRequestHeader "Cache-Control", "max-age=0"
    WinHttp.SetRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3"
    WinHttp.SetRequestHeader "Origin", "https://yiban.glut.edu.cn" '接受数据类型
    WinHttp.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2987.98 Safari/537.36 LBBROWSER" '用户浏览器信息
    WinHttp.SetRequestHeader "Accept-Language", "zh-CN,zh;q=0.8"
    WinHttp.SetRequestHeader "Upgrade-Insecure-Requests", "1"
    WinHttp.send (DAT)
    WinHttp.WaitForResponse
    checkUpdateFromHTTPS = WinHttp.ResponseText
End Function
Function checkUpdate() As Long '检查更新，软件是否可用，提示语，是否在校园网段内
On Error Resume Next

Dim DAT As String
DAT = "checkupdate={""version"":""" & App.Major * 10000 + App.Minor * 100 + App.Revision & """}"

Dim result
result = checkUpdateFromHTTPS(DAT) '优先HTTPS
If result = "" Then '尝试http
    result = checkUpdateFromHTTP(DAT)
End If
If result = "" Then '尝试https
    If ReadReg("", "auto") <> "1" Then
        DelayM 2000
        If checkUpdate() = 0 Then Exit Function
        MsgBox "暂时无法连接校园网或校园网不稳定，请连接校园网后重试！", vbCritical
        End
    Else
        checkUpdate = 1
        Exit Function
    End If
End If
'Clipboard.Clear
'Clipboard.SetText result
If Left(result, 1) <> "{" Then
    Exit Function
End If
If ReadReg("", "auto") <> "1" Then
    If objScrCtl.Eval("(" & result & ").text") <> "" Then
        MsgBox objScrCtl.Eval("(" & result & ").text"), vbInformation, "重要提醒！"
    End If
End If
If objScrCtl.Eval("(" & result & ").res") <> 100 Then
    WriteReg "", "auto", "0"
    End '不是100就退出软件
End If
If objScrCtl.Eval("(" & result & ").update") = "True" Then
    If ReadReg("", "auto") <> "1" Then
        MsgBox objScrCtl.Eval("(" & result & ").tip"), vbInformation '弹出更新提示！
    End If
    Dim paths As String
    paths = App.path & "\up" & Int(Rnd * 100000) & ".exe"
    Call URLDownloadToFile(0, objScrCtl.Eval("(" & result & ").url"), paths, 0, 0)
    If Dir(paths, vbNormal) <> "" Then
        Shell paths & " upin" & App.ExeName
        End
    Else
        MsgBox "自动更新失败，请手动更新，或下载" & objScrCtl.Eval("(" & result & ").url") & "！"
        End
    End If
End If
checkUpdate = 0
'MsgBox result

End Function
Function logout()
On Error Resume Next:
Dim WinHttp
Set WinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
'设置参数
WinHttp.SetTimeouts 60000, 60000, 60000, 5000 '设置操作超时时间
WinHttp.Option(4) = 13056 '忽略错误标志
WinHttp.Option(6) = True '为 True 时，当请求页面重定向跳转时自动跳转，False 不自动跳转，截取服务端返回的302状态。
WinHttp.Open "GET", "http://" & main.domain & "/drcom/logout?callback=dr1003&v=5023", True 'GET 或 POST, Url, False 同步方式；True 异步方式
WinHttp.SetRequestHeader "Host", main.domain
WinHttp.SetRequestHeader "Connection", "keep-alive"
WinHttp.SetRequestHeader "DNT", "1"
WinHttp.SetRequestHeader "Cache-Control", "max-age=0"
WinHttp.SetRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3"
WinHttp.SetRequestHeader "Origin", "http://" & main.domain '接受数据类型
WinHttp.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2987.98 Safari/537.36 LBBROWSER" '用户浏览器信息
WinHttp.SetRequestHeader "Accept-Language", "zh-CN,zh;q=0.8"
WinHttp.SetRequestHeader "Upgrade-Insecure-Requests", "1"
WinHttp.send
WinHttp.WaitForResponse
Dim result As String
result = BytesToBstr(WinHttp.ResponseBody, "GB2312")
If result = "" Then
Exit Function
End If
If InStr(1, result, "dr1003({") Then
    Dim js As String
    js = "{" & Split(Split(result, "dr1003({")(1), "})")(0) & "}"
    'MsgBox js
    If objScrCtl.Eval("(" & js & ").result") = "1" Then '注销成功！
        logout = True
    Else
        'MsgBox objScrCtl.Eval("(" & js & ").msga")
    End If
End If

End Function
Function ToUnixTime(strTime, intTimeZone)
On Error Resume Next
    If IsEmpty(strTime) Or Not IsDate(strTime) Then strTime = Now
    If IsEmpty(intTimeZone) Or Not IsNumeric(intTimeZone) Then intTimeZone = 0
     ToUnixTime = DateAdd("h", -intTimeZone, strTime)
     ToUnixTime = DateDiff("s", "1970-1-1 0:0:0", ToUnixTime)
End Function
Function about()
On Error Resume Next:
MsgBox "本软件由：桂林理工大学-软件工程-CH制作--------2019-12-09" & vbCrLf & _
"当前版本号：" & version & vbCrLf & _
"本软件免费使用！" & vbCrLf & _
"基于校园网上网服务！" & vbCrLf & _
"所有校园网系统操作均在本机执行，密码仅当您选择了保存密码后保存在本机！" & vbCrLf & _
"本软件不会主动收集个人隐私信息，不会上传账号密码等信息，请放心使用。" & vbCrLf & _
"但您所有的上网记录将会在学校的校园网系统中保存，这跟是否使用此软件无关。" & vbCrLf & _
"如若您使用此软件进行违法操作的后果与开发者无关！" & vbCrLf & _
"如不接受以上所有条款，请立即停止使用！" & vbCrLf & _
"使用方法：连接glut_web的wifi或连接校园网的网线，打开软件输入账号密码即可。" & vbCrLf & _
"谢谢您的使用！桂工学习交流群：60913498"
End Function
Public Sub DelayM(Msec As Long)
On Error Resume Next:
On Error Resume Next
    Dim EndTime As Long
    EndTime = Int(ToUnixTime(Now, 8)) + -Int(-Msec / 1000)
    Do
        Sleep 1
        DoEvents
    Loop While Int(ToUnixTime(Now, 8)) < EndTime
End Sub
Public Function apppath()
apppath = App.path
If Right(apppath, 1) = "\" Then apppath = Left(apppath, Len(apppath) - 1)
End Function
Public Function BytesToBstr(strBody, CodeBase)
On Error Resume Next:
Dim ObjStream
Set ObjStream = CreateObject("Adodb.Stream")
With ObjStream
.Type = 1
.Mode = 3
.Open
.Write strBody
.Position = 0
.Type = 2
.Charset = CodeBase
BytesToBstr = .ReadText
.Close
End With
Set ObjStream = Nothing
End Function
