VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cWinHttpRQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents winHttpRq As winhttp.WinHttpRequest
Attribute winHttpRq.VB_VarHelpID = -1
Private oHtmlDom As Object
Private tUrl As String
Private arrResult() As String
Private intx As Integer
Private eFlag As Boolean
Private sFlag As Boolean
Private irst As Byte

Private Sub Class_Initialize()
    eFlag = False
    sFlag = False
    irst = 0
    Set winHttpRq = CreateObject("winhttp.winhttprequest.5.1")
End Sub

Private Sub Class_Terminate()
    Set winHttpRq = Nothing
    Erase arrResult
End Sub

Friend Property Let url(ByVal sUrl As String)
    tUrl = sUrl
End Property

Friend Property Let Index(ByVal i As Integer)
    intx = i
End Property

Friend Property Get isOK() As Boolean '成功执行
    isOK = sFlag
End Property

Friend Property Get IsErr() As Boolean '成功执行
    IsErr = eFlag
End Property

Friend Property Get Result() As String() '获取返回的结果
    If eFlag = True Or sFlag = False Then Exit Property
    Dim i As Integer, k As Integer
    i = UBound(arrResult, 1) '一维数组的上界
    k = UBound(arrResult, 2) '二维数组的上界
    ReDim Result(i, k)
    Result = arrResult
End Property
'-----------------------------------------------------------class


'abort,You can abort both asynchronous and synchronous Send methods. To abort a synchronous Send method, you must call Abort from within an IWinHttpRequestEvents event.
'abort, 限于异步请求
Private Sub winHttpRq_OnResponseStart(ByVal Status As Long, ByVal ContentType As String)

End Sub

Private Sub winHttpRq_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
   Restart
End Sub

Private Sub winHttpRq_OnResponseFinished()
    DataTreat
End Sub

Friend Sub StartRe(Optional ByVal Cookie As String)
    HTTP_GetData "GET", tUrl, "http://data.10jqka.com.cn/funds/ggzjl/", sCookie:=Cookie, sCharset:="gb2312"
End Sub

Private Function DataTreat()
    Dim bResult() As Byte
    Dim sResult As String
    Dim sCharset As String
    Dim arr() As String
    Dim item As Object, itemx As Object
    Dim oTable As Object
    Dim i As Integer, k As Integer
    sCharset = "gb2312"
    bResult = winHttpRq.responseBody '按照指定的字符编码显示
    sResult = Byte2String(bResult, sCharset)
    WriteHtml sResult
    On Error Resume Next
    Set oTable = oHtmlDom.getElementsByclassName("m-table J-ajax-table")
    If oTable Is Nothing Then Set oHtmlDom = Nothing: Set oTable = Nothing: Restart
    If oTable.Length = 0 Then Set oHtmlDom = Nothing: Set oTable = Nothing: Restart
    ReDim arr(49, 10)
    ReDim arrResult(49, 10)
    For Each item In oTable.item(0).Children.item(1).Children
        k = 0
        For Each itemx In item.Children
            arr(i, k) = itemx.innertext
            k = k + 1
        Next
        i = i + 1
    Next
    arrResult = arr
    sFlag = True
    Set oTable = Nothing
    Set oHtmlDom = Nothing
End Function

Private Sub Restart() '重试
    If irst < 3 Then
        irst = irst + 1
        winHttpRq.abort
        StartRe '失败重试
    Else
        eFlag = True
    End If
End Sub

'-------------------------------------------------------------------------------------------------------------------------------------------通用
Private Function HTTP_GetData(ByVal sVerb As String, ByVal sUrl As String, Optional ByVal refUrl As String = "https://www.baidu.com", _
Optional ByVal sProxy As String, Optional ByVal sCharset As String = "utf-8", Optional ByVal sPostdata As String = "", _
Optional ByVal cType As String = "application/x-www-form-urlencoded", Optional sCookie As String = "", _
Optional ByVal acType As String, Optional ByVal cHost As String, Optional ByVal oRig As String, _
Optional ByVal acLang As String, Optional ByVal acEncode As String, _
Optional ByVal rsTimeOut As Long = 3000, Optional ByVal cTimeOut As Long = 3000, Optional ByVal sTimeOut As Long = 5000, Optional ByVal rcTimeOut As Long = 3000, Optional ByVal IsSave As Boolean) As String
    '--------------------------sVerb为发送的Html请求的方法,sUrl为具体的网址,sCharset为网址对应的字符集编码,sPostData为Post方法对应的发送body
    '- <form method="post" action="http://www.yuedu88.com/zb_system/cmd.php?act=search"><input type="text" name="q" id="edtSearch" size="12" /><input type="submit" value="搜索" name="btnPost" id="btnPost" /></form
    '----------------------https://blog.csdn.net/tylm22733367/article/details/52596990
    '------------------------------https://msdn.microsoft.com/en-us/library/windows/desktop/aa384106(v=vs.85).aspx
    '------------------------https://docs.microsoft.com/en-us/windows/win32/winhttp/iwinhttprequest-interface
    On Error GoTo ErrHandle
    With winHttpRq
        .Option(6) = 0 '为 True 时，当请求页面重定向跳转时自动跳转，False 不自动跳转，截取服务端返回的302状态
        '--------------如果不设置禁用重定向,如有道词典无法有效处理post的数据,将会跳转有道翻译的首页,返回不必要的数据
        Select Case sVerb
        Case "GET"
            .Open "GET", sUrl, True '---url, This must be an absolute URL.
        Case "POST"
            .Open "POST", sUrl, False
            .setRequestHeader "Content-Type", cType
        End Select
        If Len(sProxy) > 0 Then '检测格式是否满足要求
            If LCase(sProxy) <> "localhost:8888" Then
            '-------------------注意fiddler无法直接抓取whq的请求, 需要将代理设置为localhost:8888端口
                If InStr(sProxy, ":") > 0 And InStr(sProxy, ".") > 0 Then
                    If UBound(Split(sProxy, ".")) = 4 Then .SetProxy 2, sProxy 'localhost:8888----代理服务器/需要增加错误判断(并不是每一个代理都可用)
                End If
            Else
                .SetProxy 2, sProxy
            End If
        End If
        If Len(acEncode) > 0 Then .setRequestHeader "Accept-Encoding", acEncode
        If Len(acLang) > 0 Then .setRequestHeader "Accept-Language", acLang
        If Len(acType) > 0 Then .setRequestHeader "Accept", acType
        If Len(cHost) > 0 Then .setRequestHeader "Host", cHost
        If Len(oRig) > 0 Then .setRequestHeader "Origin", oRig
        If Len(sCookie) > 0 Then .setRequestHeader "cookie", sCookie
        .setRequestHeader "Referer", refUrl '伪装从特定的url而来
        .setRequestHeader "User-Agent", Random_UserAgent '伪造浏览器的ua
        If Len(sPostdata) > 0 Then
            .Send (sPostdata)
        Else
            .Send
        End If
        If .waitForResponse(6) = False Then eFlag = True: Set winHttpRq = Nothing: Exit Function
        If .Status <> 200 Then eFlag = True: Set winHttpRq = Nothing: Exit Function
    End With
    Exit Function
ErrHandle:
    If Err.Number = -2147012867 Then MsgBox "无法链接服务器", vbCritical, "Warning!"
    eFlag = True
    Set winHttpRq = Nothing
End Function
'---------------------------------------https://www.w3school.com.cn/ado/index.asp
Private Function Byte2String(ByRef bContent() As Byte, ByVal sCharset As String, Optional ByVal IsSave As Boolean) As String
    Const adTypeBinary As Byte = 1
    Const adTypeText As Byte = 2
    Const adModeRead As Byte = 1
    Const adModeWrite As Byte = 2
    Const adModeReadWrite As Byte = 3
    Dim oStream As Object
    '----------------------利用adodb将字节转为字符串
    Set oStream = CreateObject("ADODB.Stream")
    With oStream
        .Open
        .type = adTypeBinary
        .Write bContent
        If IsSave = True Then '获取发音
            .SaveToFile ThisWorkbook.Path & "\voice.mp3", 2
            .Close
            Set oStream = Nothing
        End If
        .Position = 0
        .type = adTypeText
        .CharSet = sCharset
         Byte2String = .ReadText
        .Close
    End With
    Set oStream = Nothing
End Function

Private Sub WriteHtml(ByVal sHtml As String) '将页面信息写到html file
    'https://www.w3.org/TR/DOM-Level-2-HTML/html
    'https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa752574%28v%3dvs.85%29
    '----------------------------------------https://ken3memo.hatenablog.com/entry/20090904/1252025888
    'https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa752573(v=vs.85)
    Set oHtmlDom = CreateObject("htmlfile")
    With oHtmlDom
        .DesignMode = "on" ' 开启编辑模式(不要直接使用.body.innerhtml=shtml,这样会导致IE浏览器打开)
        .Write sHtml ' 写入数据
    End With
End Sub

Private Function Random_IP() As String '在代理ip列表中随机挑选ip/还需要增加判断ip是否可用
    Dim i As Integer
    Dim arr() As String
    
    isReady = True
    arr = Proxy_IP
    If isReady = False Then Random_IP = "127.0.0.1:8888": Exit Function '本机/使用fiddler
    i = UBound(arr)
    i = RandNumx(i)
    If i = 0 Then i = 1
    Random_IP = arr(i, 1) & ":" & arr(i, 2)
    Set oHtmlDom = Nothing
    Erase arr
End Function
'--------------代理(proxy)设置
'----------https://docs.microsoft.com/zh-cn/windows/win32/winhttp/iwinhttprequest-setproxy
'HTTPREQUEST_PROXYSETTING_DEFAULT （0）：Default proxy setting. Equivalent to HTTPREQUEST_PROXYSETTING_PRECONFIG.
'HTTPREQUEST_PROXYSETTING_PRECONFIG（0）：Indicates that the proxy settings should be obtained from the registry.
'This assumes that Proxycfg.exe has been run. If Proxycfg.exe has not been run and HTTPREQUEST_PROXYSETTING_PRECONFIG is specified, then the behavior is equivalent to HTTPREQUEST_PROXYSETTING_DIRECT.
'HTTPREQUEST_PROXYSETTING_DIRECT（1）：Indicates that all HTTP and HTTPS servers should be accessed directly.
'Use this command if there is no proxy server.
'HTTPREQUEST_PROXYSETTING_PROXY（2）：When HTTPREQUEST_PROXYSETTING_PROXY is specified, varProxyServer should be set to a proxy server string
'and varBypassList should be set to a domain bypass list string. This proxy configuration applies only to the current instance of the WinHttpRequest object.
Private Function Proxy_IP() As String() '爬取http代理ip地址列表
    Dim sResult As String
    Dim oHtml As Object
    Dim objList As Object
    Dim arr() As String
    Dim list_item As Object, item As Object, itemx As Object
    Dim i As Integer, k As Integer
    
    On Error Resume Next
    sResult = HTTP_GetData("GET", "https://www.xicidaili.com/wn/") '此站点具有较为敏感的反爬虫防护(直接xmlhttp访问会出现503错误返回)
    WriteHtml sResult
    Set objList = oHtmlDom.getElementById("ip_list")
    i = objList.Children.Length
    If i = 0 Then isReady = False: Exit Function
    Set oHtml = objList.Children.item(i - 1)
    Set list_item = oHtml.getElementsByTagName("tr")
    i = 0
    i = list_item.Length
    If i = 0 Then isReady = False: Exit Function
    For Each item In list_item
        If k = 0 Then
            k = item.Children.Length
            If k = 0 Then isReady = False: Exit Function
            ReDim arr(i - 1, k - 1)
            ReDim Proxy_IP(i - 1, k - 1)
            k = 0: i = 0
        End If
        k = 0
        For Each itemx In item.Children
            arr(i, k) = itemx.innertext '1,ip, 2,port, 3,地址, 4,匿名, 5, http/https
            k = k + 1
        Next
        i = i + 1
    Next
    Proxy_IP = arr
    Set objList = Nothing
    Set oHtml = Nothing
    Set list_item = Nothing
    Erase arr
End Function

Private Function Random_UserAgent() As String '随机浏览器伪装
    Dim i As Byte
    Dim UA As String
    i = RandNumx(10)
    Select Case i
        Case 0: UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:61.0) Gecko/20100101 Firefox/61.0"
        Case 1: UA = "Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2987.133 Safari/537.36"
        Case 2: UA = "Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.119 Safari/537.36"
        Case 3: UA = "Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/42.0.2311.152 Safari/537.36"
        Case 4: UA = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/62.0.3202.9 Safari/537.36"
        Case 5: UA = "Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.103 Safari/537.36"
        Case 6: UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36"
        Case 7: UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100"
        Case 8: UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:62.0) Gecko/20100101 Firefox/62.0"
        Case 9: UA = "Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.131 Safari/537.36"
        Case 10: UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.77 Safari/537.36"
    End Select
    Random_UserAgent = UA
End Function

Private Function sUnicode2Character(strText As String) As String '\u30a2\u30e1\u30ea\u30ab\u5927\u7d71\u9818\u9078\u6319\u304c\u307e\u3082\u306a\u304f\u59cb\u307e\u308b
    With CreateObject("MSScriptControl.ScriptControl")
        .Language = "javascript"
        Convert = .eval("('" & strText & "').replace(/&#\d+;/g,function(b){return String.fromCharCode(b.slice(2,b.length-1))});")
    End With
End Function

Private Function Unicode2Character(ByVal strText As String) '将Unicode转为文字
    'https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa752599(v=vs.85)
    With CreateObject("htmlfile")
        .Write "<script></script>"
        '--------------------------https://www.w3school.com.cn/jsref/jsref_unescape.asp
        '该函数的工作原理是这样的：通过找到形式为 %xx 和 %uxxxx 的字符序列（x 表示十六进制的数字），用 Unicode 字符 \u00xx 和 \uxxxx 替换这样的字符序列进行解码
        'ECMAScript v3 已从标准中删除了 unescape() 函数，并反对使用它，因此应该用 decodeURI() 和 decodeURIComponent() 取而代之。
        Unicode2Character = .parentwindow.unescape(Replace(strText, "\u", "%u"))
    End With
End Function

Private Function oEncodeUrl(ByVal strText As String) As String '将字符串进行编码,需要注意的是各类符号的处理
    'http://www.ruanyifeng.com/blog/2010/02/url_encoding.html
    'https://baike.baidu.com/item/URL%E7%BC%96%E7%A0%81/3703727?fr=aladdin
    With CreateObject("msscriptcontrol.scriptcontrol")
        .Language = "JavaScript"
        oEncodeUrl = .eval("encodeURI('" & Replace(strText, "'", "\'") & "');")
    End With
End Function
