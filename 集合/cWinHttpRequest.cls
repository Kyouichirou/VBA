VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cWinHttpRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'WinHttpRequestOption_UserAgentString = 0;  '设置或检索包含用户代理字符串的VARIANT。
'WinHttpRequestOption_URL = 1;               '检索包含资源的URL的VARIANT。此值为只读;您不能使用此属性设置URL。直到打开方法被调用才能读取URL。此选项在完成发送方法以验证是否发生任何重定向后对URL进行检查非常有用。
'WinHttpRequestOption_URLCodePage = 2;        '设置或检索标识URL字符串的代码页的VARIANT。默认值为UTF-8代码页。代码页用于将以打开方法传递的Unicode URL字符串转换为单字节字符串表示形式。
'WinHttpRequestOption_EscapePercentInURL = 3;  '设置或检索VARIANT，指示URL字符串中的百分比字符是否转换为转义序列。此选项的默认值为VARIANT_TRUE，它指定所有不安全的美国国家标准学院（ANSI）字符，除了将百分比符号转换为转义序列。
'WinHttpRequestOption_SslErrorIgnoreFlags = 4; WinHTTP版本5.1中此选项的默认值为零，这样就不会忽略错误。在早期版本的WinHTTP中，默认设置为0x3300(13056)，导致默认情况下忽略所有服务器证书错误。 = 忽略错误标志
'WinHttpRequestOption_SelectCertificate = 5; ' '设置一个VARIANT，指定发送到服务器进行认证的客户端证书。此选项指示用反斜杠分隔的客户端证书的位置，证书库和主题。有关选择客户端证书的更多信息，请参阅WinHTTP中的SSL.
'WinHttpRequestOption_EnableRedirects = 6; '为True时,当请求的页面中有跳转时,抓取跳转页面信息.False相反不抓取
'WinHttpRequestOption_UrlEscapeDisable = 7;            设置或检索VARIANT，指示URL的路径和查询组件中的不安全字符是否转换为转义序列。此选项的默认值为VARIANT_TRUE，它指定路径和查询中的字符将被转换。
'WinHttpRequestOption_UrlEscapeDisableQuery = 8;       '设置或检索一个VARIANT，指示URL的查询组件中的不安全字符是否转换为转义序列。此选项的默认值为VARIANT_TRUE，它指定查询中的字符将被转换。
'WinHttpRequestOption_SecureProtocols = 9;             '设置或检索指示可以使用哪些安全协议的VARIANT。此选项选择客户端可接受的协议。该协议在安全套接字层（SSL）握手期间协商。这可以是一个或多个以下标志的组合。
'WinHttpRequestOption_EnableTracing = 10;              '设置或检索一个VARIANT，指示是否启用跟踪。有关Microsoft Windows HTTP服务（WinHTTP）中的跟踪工具的更多信息，请参阅WinHTTP跟踪设备.
'WinHttpRequestOption_RevertImpersonationOverSsl = 11;  '控制在SSL证书身份验证操作期间WinHttpRequest对象是否临时还原客户端模拟。WinHttpRequest对象的默认设置为TRUE。将此选项设置为FALSE以在执行证书身份验证操作时保持模拟。
'WinHttpRequestOption_EnableHttpsToHttpRedirects = 12; '控制WinHTTP是否允许重定向。默认情况下，将自动遵循所有重定向，但从安全（https）URL转移到非安全（http）URL的用户除外。将此选项设置为TRUE以启用HTTPS到HTTP重定向。
'WinHttpRequestOption_EnablePassportAuthentication = 13; '启用或禁用对Passport身份验证的支持。默认情况下，Passport身份验证的自动支持被禁用;将此选项设置为TRUE以启用Passport身份验证支持。
'WinHttpRequestOption_MaxAutomaticRedirects = 14; '设置或检索WinHTTP遵循的最大重定向数量;默认值为10。此限制可防止未经授权的网站在大量重定向后使WinHTTP客户端停顿。
'WinHttpRequestOption_MaxResponseHeaderSize = 15;  '在服务器响应的头部分的最大大小上设置或检索绑定集。此绑定保护客户端免受尝试通过发送具有无限量的头数据的响应来阻止客户端的恶意服务器。默认值为64 KB。
'WinHttpRequestOption_MaxResponseDrainSize = 16;  '在SSL协商期间启用服务器证书吊销检查。当服务器提供证书时，执行检查以确定该证书是否已由其颁发者撤销。如果证书确实被撤销，或撤销检查失败，因为证书吊销列表（CRL）无法下载，则请求失败;这种撤销错误不能被抑
'WinHttpRequestOption_EnableHttp1_1 = 17;  '设置或检索一个布尔值，指示是否应使用HTTP / 1.1或HTTP / 1.0。默认值为TRUE，默认使用HTTP / 1.1。
'WinHttpRequestOption_EnableCertificateRevocationCheck = 18 '在SSL协商期间启用服务器证书吊销检查。当服务器提供证书时，执行检查以确定该证书是否已由其颁发者撤销。如果证书确实被撤销，或撤销检查失败，因为证书吊销列表（CRL）无法下载，则请求失败;这种撤销错误不能被抑
'-----------------------------------------------------------Option setup
Private oWinHttpRQ As Object
Private oStream As Object
Private isPost As Boolean     '请求方式
Private tCharset As String    '字符集
Private sResult As String     '执行的结果
Private sFilePath As String   '文件保存位置
Private BasePath As String    '工作薄的工作位置
Private Const adTypeBinary As Byte = 1
Private Const adTypeText As Byte = 2
Private Const adModeRead As Byte = 1
Private Const adModeWrite As Byte = 2
Private Const adModeReadWrite As Byte = 3
Private Const commEType As String = "application/x-www-form-urlencoded" '常用用于接受的数据类型
Private Const commJType As String = "application/json"
Private Const FidderProxy As String = "localhost:8888" 'fiddler本地代理设置
'------------------------------------------安全协议
Public Enum SecurityProtocol
    SSL2 = &H8        '默认情况下启用 SSL 2.0 '此协议已经处于不安全的状态
    SSL3 = &H20       '默认情况下启用 SSL 3.0 '此协议存在较大的安全隐患
    TLS1 = &H80       '默认情况下启用 TLS 1.0
    TLS2 = &H200      '默认情况下启用 TLS 1.1
    TLS3 = &H800      '默认情况下启用 TLS 1.2
End Enum

Public Enum AutoLogonPolicy '自动登陆配置
    Always = 0
    OnlyIfBypassProxy = 1
    Never = 2
End Enum

Public Enum mRequest '请求的类型, 值得注意的HEAD, 用于请求头部资源
    mGET = 1
    mPost = 2
    mHead = 3
    mTrace = 4
    mOptions = 5
    mConnect = 6
End Enum

'未知的认证机构（CA）或不受信任的根证书
'使用错误
'通用名称（CN）无效
'无效的日期或证书已过期
Public Enum SslErrorIgnoreFlags
    AllErr = &H3300
    UnCA_Root = &H100
    wUsage = &H200
    InvalidCN = &H1000
    InvalidExpired = &H2000
End Enum
'CLSID_WinHttpRequest
'CreateObject("new:{2087C2F4-2CEF-4953-A8AB-66779B670495}")
Public ProxyUsername As String '代理设置
Public ProxyServer As String
Public ProxyPassword As String
Public ProxyBypassList As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function timeGetTime Lib "winmm.dll" () As Long '时间 -单词训练

Friend Sub SetProxyConf(ByVal sProxy As String, Optional ByVal nProxy As String, Optional ByVal pProxy As String, Optional ByRef byProxy As Variant) '代理服务器配置
    Dim strTemp As String
    
    strTemp = Trim$(sProxy)
    If LenB(strTemp) = 0 Then Exit Sub
    If strTemp <> FidderProxy Then
        If InStr(1, strTemp, ChrW$(46), vbBinaryCompare) = 0 Then Exit Sub    '需要增加错误提醒
        If UBound(Split(strTemp, ChrW$(46), -1, vbBinaryCompare)) <> 3 Then Exit Sub '可以改成正则校检是否满足 xxx.xxx.xxx.xxx地址形式
    End If
    If InStr(1, strTemp, ChrW$(58), vbBinaryCompare) = 0 Then Exit Sub '端口
    If IsArray(byProxy) = True Then
        Dim i As Byte
        Dim k As Byte, m As Byte
        i = UBound(byProxy)
        k = LBound(byProxy)
        For m = k To i
            strTemp = strTemp & Trim$(byProxy(m)) & ChrW$(44) '逗号
        Next
        strTemp = Left$(strTemp, Len(strTemp) - 1)
    Else
        strTemp = Trim$(byProxy)
    End If
    With Me
        .ProxyServer = sProxy
        .ProxyUsername = nProxy
        .ProxyPassword = pProxy
        .ProxyBypassList = strTemp
    End With
End Sub

Friend Sub SetProxy(Optional ByVal isCancel As Boolean)   '代理服务器设置     '2表示默认仅在request中起作用
    If isCancel = True Then
        oWinHttpRQ.SetProxy 0
    Else
        If LenB(Trim$(ProxyServer)) > 0 Then
            oWinHttpRQ.SetProxy 2, sProxy
            If Len(Trim$(ProxyUsername)) > 0 Then
                If Len(Trim$(ProxyUsername)) > 0 And Len(Trim$(ProxyPassword)) > 0 Then SetCredentials ProxyUsername, ProxyPassword, True
            End If
        End If
    End If
End Sub
'-------------------------------------------------------------代理设置

Friend Property Get Get_NowUrl() As String '获取现在的Url, 假如出现跳转等情况
    Get_NowUrl = oWinHttpRQ.Option(1)      '此项不能设置访问url, 只能获取值
End Property

Friend Sub CacheTime(ByVal isWait As Boolean, Optional ByVal WaitTime As Long) '等待时间
If isWait = True Then
If WaitTime = 0 Then WaitTime = 200
End If
End Sub


Friend Property Get Get_Cookie(Optional ByVal isAll As Boolean = False) As String '获取cookie
    Dim sHead As String
    
    sHead = oWinHttpRQ.getAllResponseHeaders
    If InStr(1, sHead, "Set-Cookie", vbTextCompare) > 0 Then
        If isAll = False Then
            Get_Cookie = oWinHttpRQ.getResponseHeader("Set-Cookie") '假如没有设置cookie那么将会出错, 单个cookie
        Else
            Dim xCookie As Variant
            Dim i As Byte, k As Byte
            xCookie = Split(strTemp, "Set-Cookie:", -1, vbTextCompare)
            i = UBound(xCookie)
            sHead = ""
            For k = 1 To i
                If InStr(1, xCookie(k), ChrW$(59), vbBinaryCompare) > 0 Then sHead = sHead & Trim$(Split(xCookie(k), ";", -1, vbBinaryCompare)(0)) & "; " '拼接在一起
            Next
            sHead = Trim$(sHead)
            Get_Cookie = Left$(sHead, Len(sHead) - 1)
        End If
    End If
End Property

Friend Property Get Get_Charset() As String '获取字符集
    Dim sHead As String
    Dim strTemp As String
    Dim sCharset As String

    sHead = oWinHttpRQ.getAllResponseHeaders              '发出请求后
    If InStr(1, sHead, "charset=", vbTextCompare) > 0 Then
        strTemp = Split(strTemp, "charset=")(1)
        strTemp = LCase$(Left$(strTemp, 7))
        If InStr(1, strTemp, "gbk", vbBinaryCompare) > 0 Then
            sCharset = "gb2312"
        ElseIf InStr(1, strTemp, "gb2312", vbBinaryCompare) > 0 Then
            sCharset = "gb2312"
        ElseIf InStr(1, strTemp, "unicode", vbBinaryCompare) > 0 Then
            sCharset = "unicode"
        ElseIf InStr(1, strTemp, "utf-8", vbBinaryCompare) > 0 Then
            sCharset = "utf-8"
        Else
            sCharset = "utf-8"
        End If
        Get_Charset = sCharset
    End If
End Property

Friend Property Get Get_Result() As String '获取返回的数据
    Get_Result = sResult
End Property
'-------------------------------------------------------------------获取属性

'---------------------------------------------------------重要的伪装, 用于反爬
Friend Property Let RefererUrl(ByVal refUrl As String)
    oWinHttpRQ.setRequestHeader "Referer", refUrl    '伪装从特定的url而来
End Property

Friend Property Let Cookie(ByVal sCookie As String)
    oWinHttpRQ.setRequestHeader "Cookie", sCookie    'Cookie
End Property

Friend Property Let Origin(ByVal oRig As String)
    oWinHttpRQ.setRequestHeader "Origin", oRig       '源起
End Property

Friend Sub UserAgent(Optional ByVal sCustomUA As String, Optional ByVal isMobile As Boolean, Optional ByVal isForceIE As Boolean) '随机浏览器
    If LenB(Trim$(sCustomUA)) > 0 Then
        oWinHttpRQ.setRequestHeader "User-Agent", sCustomUA '自定义
    Else
        oWinHttpRQ.setRequestHeader "User-Agent", Random_UserAgent(isMobile, isForceIE) '从列表中选取
    End If
End Sub
'------------------------------------------------------------伪装

'-------------------------------------------------------------------------------------默认执行的项
Friend Sub CharSet(Optional ByVal isCheck As Boolean, Optional ByVal sCharset As String = "utf-8") '设置请求数据的方法 ,ischeck用于指示自动检查网站的编码, 这里就将使用到Head请求数据的方式
    If isCheck = False Then
        tCharset = sCharset
    Else
        tCharset = Get_Charset
        If LenB(tCharset) = 0 Then tCharset = sCharset
    End If
End Sub

Friend Sub Redirect(Optional ByVal isRedirect As Boolean = False) '重定向, 即访问a, 被跳转b, 如网站的反爬, 或者错误重定向, 如果不允许跳转将直接返回错误302
    oWinHttpRQ.Option(6) = isRedirect
End Sub

Friend Sub TimeOut(Optional ByVal rsTimeOut As Long = 3000, Optional ByVal cTimeOut As Long = 3000, Optional ByVal sTimeOut As Long = 5000, Optional ByVal rcTimeOut As Long = 3000) '设置请求数据的方法
    oWinHttpRQ.setTimeouts rsTimeOut, cTimeOut, sTimeOut, rcTimeOut '-----ResolveTimeout, ConnectTimeout, SendTimeout, ReceiveTimeout
End Sub
'----------------------------------默认执行项

Friend Property Let mVerb(ByVal mPost As Boolean) '设置请求数据的方法
    isPost = mPost
End Property

Private Function Request_Mode(ByVal intmode As Byte) As String '选择请求的方式
    Dim sMode As String
    Select Case intmode
        Case 1: sMode = "GET"
        Case 2: sMode = "POST"
        Case 3: sMode = "HEAD"
        Case 4: sMode = "TRACE"
        Case 5: sMode = "OPTIONS"
        Case 6: sMode = "CONNECT"
    End Select
    Request_Mode = sMode
End Function

Friend Sub Send(Optional ByVal sPostdata As String, Optional ByVal isFilesave As Boolean = False)
    Dim bResult() As Byte
    If isPost = False Then
        oWinHttpRQ.Send
    Else
        oWinHttpRQ.Send (sPostdata)
    End If
    If oWinHttpRQ.Status <> 200 Then Exit Sub
    If sVerb = "HEAD" Then sResult = oWinHttpRQ.getAllResponseHeaders: Exit Sub
    bResult = oWinHttpRQ.responseBody
    If isFilesave = False Then
        Byte2String bResult, tCharset
    Else
        Byte2File bResult
    End If
End Sub

Friend Sub Request(ByVal sUrl As String, ByVal iMode As mRequest, Optional ByVal isAynch As Boolean = False) '最先执行, set-header部分才能设置
    sUrl = LCase$(Trim$(sUrl))
    If InStr(1, sUrl, "http", vbBinaryCompare) <> 1 Then Exit Sub  '检查url是否满足标砖
    If isPost = False Then
        oWinHttpRQ.Open Request_Mode(iMode), sUrl, isAynch
    Else
        oWinHttpRQ.Open "POST", sUrl, False    'get, post, head, 三者独立处理, 其他的混合
    End If
End Sub

'-------------------------------------------------------------------------------------------------安全设置
Friend Property Let Authorization(ByVal Username As String, ByVal Password As String)
    oWinHttpRQ.setRequestHeader "Authorization", "Basic " + Base64(Username + ":" + Password) '需要引入base64
End Property

' const HTTPREQUEST_SETCREDENTIALS_FLAGS  HTTPREQUEST_SETCREDENTIALS_FOR_SERVER   = 0x00000000;
' const HTTPREQUEST_SETCREDENTIALS_FLAGS  HTTPREQUEST_SETCREDENTIALS_FOR_PROXY    = 0x00000001;
Friend Sub SetCredentials(ByVal Username As String, ByVal Password As String, Optional ByVal isProxy As Boolean)
    Const HTTPREQUEST_SETCREDENTIALS_FOR_SERVER As Byte = 0
    Const HTTPREQUEST_SETCREDENTIALS_FOR_PROXY  As Byte = 1
    Dim sType As Byte
    If isProxy = False Then sType = 0 Else sType = 1
    oWinHttpRQ.SetCredentials Username, Password, sType
End Sub

Friend Sub Https2Http_Redirect(Optional ByVal isRedirect As Boolean = False) 'https跳转http
    oWinHttpRQ.Option(12) = isRedirect
End Sub

Friend Sub EnableCertificateCheck(Optional ByVal isChech As Boolean = False) '证书吊销检查'Revocation
    oWinHttpRQ.Option(18) = isChech
End Sub

Friend Sub SSLProtocol(ByVal intProtocol As SecurityProtocol) '设置安全协议, SSL3.0, TLS 1.x
    oWinHttpRQ.Option(9) = intProtocol
End Sub

Friend Property Let SslErrorIgnore(ByVal IngoreTye As SslErrorIgnoreFlags) '忽视证书错误
    If isIgnore = True Then oWinHttpRQ.Option(4) = IngoreTye
End Property

'"LOCAL_MACHINE\Personal\My Middle-Tier Certificate", 证书路径
'The string specified in the ClientCertificate parameter consists of the certificate location, certificate store, and subject name delimited by backslashes.
'bstrClientCertificate参数中指定的字符串由以反斜杠分隔的证书位置，证书存储和主题名称组成。有关证书字符串组件的更多信息，请参阅客户证书.
'证书库名称和位置是可选的。但是，如果指定证书存储，则还必须指定该证书存储的位置。默认位置为CURRENT_USER，默认证书库为“MY”。空白主题表示应使用证书存储中的第一个证书。
'在致电发送发送请求之前，请调用SetClientCertificate选择证书。
Friend Property Let SetClientCertificate(ByVal sClient As String) '证书设置
    If Len(Trim$(sClient)) = 0 Then Exit Property
    oWinHttpRQ.SetClientCertificate sClient
End Property
'----------------------------------------------------------------------------------------安全设置

'-----------------------------------------------------------------------------------一般属性设置
Friend Property Let CustomHeader(ByVal sHeader As String, ByVal sParameter As String) '自定义请求头, 如:Cache-Control, no-cache
    If LenB(Trim$(sHeader)) = 0 Then Exit Property
    If LenB(Trim$(sParameter)) = 0 Then Exit Property
    oWinHttpRQ.setRequestHeader sHeader, sParameter
End Property

Friend Property Let URLCodePage(ByVal iCodePage As CodePage) '设置或检索标识URL字符串的代码页的VARIANT。默认值为UTF-8代码页
    oWinHttpRQ.Option(2) = iCodePage
End Property

Friend Property Let AutoLogonPolicy(ByVal iPolicy As AutoLogonPolicy) '自动登陆, 注意: 可能需要和密码授权等安全设置搭配使用
    oWinHttpRQ.SetAutoLogonPolicy(2) = iPolicy
End Property

Friend Property Let Accept_Encoding(ByVal acEncode As String) '如果发送，会返回 gzip, deflate 压缩过的编码
    oWinHttpRQ.setRequestHeader "Accept-Encoding", acEncode
End Property

Friend Property Let Conten_Length(ByVal cLength As Integer) '内容长度
    oWinHttpRQ.setRequestHeader "Content-Length", cLength
End Property

Friend Property Let Accept_Language(ByVal acLang As String) '接受语言
    oWinHttpRQ.setRequestHeader "Accept-Language", acLang
End Property

Friend Property Let Host(ByVal cHost As String)            '主机
    oWinHttpRQ.setRequestHeader "Host", cHost
End Property

Friend Property Let Accept_Type(ByVal acType As String)    '接受类型
    oWinHttpRQ.setRequestHeader "Accept", acType
End Property

Friend Property Let Requested_With(ByVal xReqw As String)
    oWinHttpRQ.setRequestHeader "X-Requested-With", xReqw
End Property

Friend Property Let Content_Type(ByVal cType As String)   '接受编码
    oWinHttpRQ.setRequestHeader "Content-Type", cType
End Property

'--------------------------------------------------------------------------------------------辅助项
'主站, 上一级, 访问的url
Private Function AutoCreateRef(ByVal sUrl As String, Optional ByVal iMode As Byte = 1) As String '自动生成referURL
    Dim strTemp As String, sHttp As String
    Dim x As Variant
    
    If iMode = 0 Then                                        '访问的url
        AutoCreateRef = sUrl
    Else
        x = Split(sUrl, "//", -1, vbBinaryCompare)
        sHttp = x(0)
        strTemp = x(1)
        If InStr(1, strTemp, "/", vbBinaryCompare) = 0 Then  '没有斜杠
            AutoCreateRef = sUrl
        Else
            x = Split(strTemp, "/", -1, vbBinaryCompare)
            If iMode = 1 Then                                '主站
                sUrl = sHttp & "//" & x(0) & "/"
            Else                                             '上一级
                sUrl = Left$(sUrl, Len(sUrl) - Len(x(UBound(x))))
            End If
            AutoCreateRef = sUrl
        End If
    End If
End Function

Private Function Random_UserAgent(ByVal isMobile As Boolean, Optional ByVal ForceIE As Boolean = False) As String '随机浏览器伪装/手机-PC
    Dim i As Byte
    Dim UA As String

    If ForceIE = True Then '使用ie
        UA = "Mozilla/5.0 (Windows NT 10.0; WOW64; Trident/7.0; rv:11.0) like Gecko" 'Mozilla/5.0(compatible;MSIE9.0;WindowsNT6.1;Trident/5.0)
    Else
        i = Random_Num(10)
        If isMobile = True Then
            Select Case i
            Case 0: UA = "UCWEB/2.0 (MIDP-2.0; U; Adr 9.0.0) UCBrowser U2/1.0.0 Gecko/63.0 Firefox/63.0 iPhone/7.1 SearchCraft/2.8.2 baiduboxapp/3.2.5.10 BingWeb/9.1 ALiSearchApp/2.4"
            Case 1: UA = "Mozilla/5.0 (Linux; Android 7.0; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/48.0.2564.116 Mobile Safari/537.36 T7/10.3 SearchCraft/2.6.2 (Baidu; P1 7.0)"
            Case 2: UA = "MQQBrowser/26 Mozilla/5.0 (Linux; U; Android 2.3.7; zh-cn; MB200 Build/GRJ22; CyanogenMod-7) AppleWebKit/533.1 (KHTML, like Gecko) Version/4.0 Mobile Safari/533.1"
            Case 3: UA = "Mozilla/5.0 (Linux; U; Android 8.1.0; zh-cn; BLA-AL00 Build/HUAWEIBLA-AL00) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/57.0.2987.132 MQQBrowser/8.9 Mobile Safari/537.36"
            Case 4: UA = "Mozilla/5.0 (Linux; Android 6.0.1; OPPO A57 Build/MMB29M; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/63.0.3239.83 Mobile Safari/537.36 T7/10.13 baiduboxapp/10.13.0.10 (Baidu; P1 6.0.1)"
            Case 5: UA = "Mozilla/5.0 (Linux; Android 8.0; MHA-AL00 Build/HUAWEIMHA-AL00; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/57.0.2987.132 MQQBrowser/6.2 TBS/044304 Mobile Safari/537.36 MicroMessenger/6.7.3.1360(0x26070333) NetType/4G Language/zh_CN Process/tools"
            Case 6: UA = "Mozilla/5.0 (iPhone; CPU iPhone OS 12_1_4 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) FxiOS/15.0b13894 Mobile/16D57 Safari/605.1.15"
            Case 7: UA = "Mozilla/5.0 (iPhone; CPU iPhone OS 12_1_4 like Mac OS X; zh-cn) AppleWebKit/601.1.46 (KHTML, like Gecko) Mobile/16D57 Quark/3.0.6.926 Mobile"
            Case 8: UA = "Mozilla/5.0 (iPhone; CPU iPhone OS 12_1_4 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/10.0 Mobile/16D57 Safari/602.1 MXiOS/5.2.20.508"
            Case 9: UA = "Mozilla/5.0 (iPhone; CPU iPhone OS 12_1_4 like Mac OS X) AppleWebKit/606.4.5 (KHTML, like Gecko) Mobile/16D57 QHBrowser/317 QihooBrowser/4.0.10"
            Case 10: UA = "Mozilla/5.0 (iPhone; CPU iPhone OS 12_1_4 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Mobile/16D57 unknown BingWeb/6.9.8.1"
            End Select
        Else
            Select Case i
                Case 0: UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:61.0) Gecko/20100101 Firefox/61.0"
                Case 1: UA = "Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2987.133 Safari/537.36"
                Case 2: UA = "Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.119 Safari/537.36"
                Case 3: UA = "Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/42.0.2311.152 Safari/537.36"
                Case 4: UA = "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.169 Safari/537.36 OPR/44.0.2213.246"
                Case 5: UA = "Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.103 Safari/537.36"
                Case 6: UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36"
                Case 7: UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100"
                Case 8: UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.149 Safari/537.36"
                Case 9: UA = "Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.131 Safari/537.36"
                Case 10: UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.77 Safari/537.36"
            End Select
        End If
    End If
    'UA = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/62.0.3202.9 Safari/537.36"
    Random_UserAgent = UA
End Function

'------------------------------------------------------------------------------数据处理
Friend Property Let File_SavePath(ByVal FilePath As String) '文件保存路径
    sFilePath = FilePath
End Property

Private Function AutoCreatePath(ByVal fileExtension As String) As String '自动生成文件保存路径 'StrReverse$("temp1591931900490")

End Function

Private Sub Byte2String(ByRef bContent() As Byte, ByVal sCharset As String) '生成字符串
    With oStream
        .Open
        .type = adTypeBinary
        .Write bContent
        .Position = 0
        .type = adTypeText
        .CharSet = sCharset
         sResult = .ReadText
        .Close
    End With
End Sub

Private Function Byte2File(ByRef bContent() As Byte) As Boolean '生成文件
    On Error GoTo ErrHandle
    If LenB(sFilePath) = "" Then sFilePath = BasePath & Format(Now, "yyyymmddhhmmss") & Random_Num(9999, 1000)
    Byte2File = True
    With oStream
        .Open
        .type = adTypeBinary
        .Write bContent
        .SaveToFile sFilePath, 2
        .Close
    End With
    Exit Function
ErrHandle:
    Byte2File = False
End Function

'--------------------------------------------------------------输入字典Object, 参数采用字典结构输入
Private Function String2Json(ByVal oDict As Object) As String '将字符串转为json结构的数据, {"User":"ab","PW":123}, {"User":"ab","Detail":{"coutry":"cn","language":"sp"}}
    Dim sKey As Variant, Result As String, sValue As String
    
    Result = "{"
    For Each sKey In oDict.Keys
        Result = Result & IIf(Len(Result) > 1, ",", "")
        If TypeName(oDict(sKey)) = "Dictionary" Then
            Value = String2Json(oDict(sKey))
            String2Json = sValue
        Else
            If IsNumeric(oDict(sKey)) = False Then
                Value = """" & oDict(sKey) & """"
            Else
                Value = oDict(sKey)
            End If
        End If
        Result = Result & """" & sKey & """:" & sValue & ""
    Next
    Result = Result & "}"
    String2Json = Result
End Function

Private Function Random_Num(ByVal iEnd As Long, Optional ByVal iStart As Long = 0) As Long '随机数, 生成特定范围的随机数, 不重复的随机数 '
    Randomize (Timer)                                    '初始化随机数种子
    Random_Num = Int((iEnd - iStart + 1) * Rnd + iStart) 'Int 和 Fix 函数如何返回数字的整数部分。 如果为负数参数，则 Int 函数将返回小于或等于该数的第一个负整数； Fix 函数返回大于或等于该数的第一个负整数
End Function

'-----------------------------类
Private Sub Class_Initialize()
    Set oWinHttpRQ = CreateObject("winhttp.winhttprequest.5.1")
    Set oStream = CreateObject("adodb.stream")
    BasePath = ThisWorkbook.Path & ChrW$(92)
End Sub

Private Sub Class_Terminate()
    Set oWinHttpRQ = Nothing
    Set oStream = Nothing
    sResult = "": sFilePath = "": isPost = False: tCharset = ""
    BasePath = vbNullString
End Sub

Private Function Time_sTamp() As String     '生成13位的时间戳, now不行, 只能生成10位的, tiemgettime这个函数是自系统启动开始的计时(49天重置)
    Time_sTamp = ((Round(CDbl(Now), 5) - 70 * 365 - 19) * 86400 - 8 * 3600) * 1000
End Function

Private Function CheckRname(ByVal cText As String) As String '将文件名中的非法字符替换掉
    Dim strTemp As String
    Dim Codex As Variant
    Dim rText As String
    Dim i As Byte, k As Byte
    '------------------------------------------------------------------其他的涉及到文件命名的也可以调用这个模块
   
    strTemp = StrConv(cText, vbFromUnicode) '清除掉Unicode字符
    strTemp = StrConv(cText, vbUnicode)
    Codex = Array(58, 124, 92, 60, 62, 34, 39, 42, 63, 47)
    k = UBound(Codex)
    rText = ChrW$(45)
    For i = 0 To k
        strTemp = Replace$(strTemp, ChrW$(Codex(i)), rText, 1, , vbBinaryCompare)
    Next
    CheckRname = strTemp
End Function





