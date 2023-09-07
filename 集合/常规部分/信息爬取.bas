Attribute VB_Name = "信息爬取"
'These constants and corresponding values indicate HTTP status codes returned by servers on the Internet.
'HTTP_STATUS_CONTINUE        '网页相应状态参数
'100
'The request can be continued.
'HTTP_STATUS_SWITCH_PROTOCOLS
'101
'The server has switched protocols in an upgrade header.
'HTTP_STATUS_OK
'200
'The request completed successfully.
'HTTP_STATUS_CREATED
'201
'The request has been fulfilled and resulted in the creation of a new resource.
'HTTP_STATUS_ACCEPTED
'202
'The request has been accepted for processing, but the processing has not been completed.
'HTTP_STATUS_PARTIAL
'203
'The returned meta information in the entity-header is not the definitive set available from the originating server.
'HTTP_STATUS_NO_CONTENT
'204
'The server has fulfilled the request, but there is no new information to send back.
'HTTP_STATUS_RESET_CONTENT
'205
'The request has been completed, and the client program should reset the document view that caused the request to be sent to allow the user to easily initiate another input action.
'HTTP_STATUS_PARTIAL_CONTENT
'206
'The server has fulfilled the partial GET request for the resource.
'HTTP_STATUS_WEBDAV_MULTI_STATUS
'207
'During a World Wide Web Distributed Authoring and Versioning (WebDAV) operation, this indicates multiple status codes for a single response. The response body contains Extensible Markup Language (XML) that describes the status codes. For more information, see HTTP Extensions for Distributed Authoring.
'HTTP_STATUS_AMBIGUOUS
'300
'The requested resource is available at one or more locations.
'HTTP_STATUS_MOVED
'301
'The requested resource has been assigned to a new permanent Uniform Resource Identifier (URI), and any future references to this resource should be done using one of the returned URIs.
'HTTP_STATUS_REDIRECT
'302
'The requested resource resides temporarily under a different URI.
'HTTP_STATUS_REDIRECT_METHOD
'303
'The response to the request can be found under a different URI and should be retrieved using a GET HTTP verb on that resource.
'HTTP_STATUS_NOT_MODIFIED
'304
'The requested resource has not been modified.
'HTTP_STATUS_USE_PROXY
'305
'The requested resource must be accessed through the proxy given by the location field.
'HTTP_STATUS_REDIRECT_KEEP_VERB
'307
'The redirected request keeps the same HTTP verb. HTTP/1.1 behavior.
'HTTP_STATUS_BAD_REQUEST
'400
'The request could not be processed by the server due to invalid syntax.
'HTTP_STATUS_DENIED
'401
'The requested resource requires user authentication.
'HTTP_STATUS_PAYMENT_REQ
'402
'Not implemented in the HTTP protocol.
'HTTP_STATUS_FORBIDDEN
'403
'The server understood the request, but cannot fulfill it.
'HTTP_STATUS_NOT_FOUND
'404
'The server has not found anything that matches the requested URI.
'HTTP_STATUS_BAD_METHOD
'405
'The HTTP verb used is not allowed.
'HTTP_STATUS_NONE_ACCEPTABLE
'406
'No responses acceptable to the client were found.
'HTTP_STATUS_PROXY_AUTH_REQ
'407
'Proxy authentication required.
'HTTP_STATUS_REQUEST_TIMEOUT
'408
'The server timed out waiting for the request.
'HTTP_STATUS_CONFLICT
'409
'The request could not be completed due to a conflict with the current state of the resource. The user should resubmit with more information.
'HTTP_STATUS_GONE
'410
'The requested resource is no longer available at the server, and no forwarding address is known.
'HTTP_STATUS_LENGTH_REQUIRED
'411
'The server cannot accept the request without a defined content length.
'HTTP_STATUS_PRECOND_FAILED
'412
'The precondition given in one or more of the request header fields evaluated to false when it was tested on the server.
'HTTP_STATUS_REQUEST_TOO_LARGE
'413
'The server cannot process the request because the request entity is larger than the server is able to process.
'HTTP_STATUS_URI_TOO_LONG
'414
'The server cannot service the request because the request URI is longer than the server can interpret.
'HTTP_STATUS_UNSUPPORTED_MEDIA
'415
'The server cannot service the request because the entity of the request is in a format not supported by the requested resource for the requested method.
'HTTP_STATUS_RETRY_WITH
'449
'The request should be retried after doing the appropriate action.
'HTTP_STATUS_SERVER_ERROR
'500
'The server encountered an unexpected condition that prevented it from fulfilling the request.
'HTTP_STATUS_NOT_SUPPORTED
'501
'The server does not support the functionality required to fulfill the request.
'HTTP_STATUS_BAD_GATEWAY
'502
'The server, while acting as a gateway or proxy, received an invalid response from the upstream server it accessed in attempting to fulfill the request.
'HTTP_STATUS_SERVICE_UNAVAIL
'503
'The service is temporarily overloaded.
'HTTP_STATUS_GATEWAY_TIMEOUT
'504
'The request was timed out waiting for a gateway.
'HTTP_STATUS_VERSION_NOT_SUP
'505
'The server does not support the HTTP protocol version that was used in the request message.
'--https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualbasic.calltype?view=netframework-4.8 -callbyname
'-https://docs.microsoft.com/zh-cn/office/vba/Language/Reference/user-interface-help/callbyname-function
'1.访问时间过长
'2.返回的数据乱码
'3.处理返回的数据,正则,htmlfile,json
'4.代理ip访问
'5.异步
'6.多线程
'post json类型的数据, 也来越多的站点采用json数据作为postdata
'WEBSERVICE 函数- https://support.microsoft.com/zh-cn/office/webservice-%E5%87%BD%E6%95%B0-0546a35a-ecc6-4739-aed7-c0b7ce1562c4
'ENCODEURL 函数- https://support.microsoft.com/zh-cn/office/encodeurl-%E5%87%BD%E6%95%B0-07c7fb90-7c60-4bff-8687-fac50fe33d0e
'FILTERXML 函数- https://support.microsoft.com/zh-cn/office/filterxml-%E5%87%BD%E6%95%B0-4df72efc-11ec-4951-86f5-c1374812f5b7?ui=zh-cn&rs=zh-cn&ad=cn
'http://excel880.com/blog/archives/3527
'https://www.bilibili.com/video/av75931500
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function timeGetTime Lib "winmm.dll" () As Long '时间 -单词训练
Private Declare Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Dim oHtmlDom As Object  '文档
Dim isReady As Boolean '判断执行的情况
Public Pagexs As Integer '返回页码
Dim regEx As Object '正则表达式
Dim oTli As Object '获取对象属性
Dim xmCookie As String
Dim tkCookie As String '虾米cookie

 '-----------------------获取小说
Function ObtainPage_Info(ByVal strText As String, ByVal cmCode As Byte) As String() '获取页面信息(搜索结果-章节目录-页面数)
    Dim sVerb As String, strx As String
    Dim sUrl As String
    Dim sCharset As String
    Dim sPostdata As String
    Dim arr() As String
    Dim sResult As String, Urlx As String
    Const searchUrl As String = "http://www.yuedu88.com/search.php?q="
    
    If Len(strText) = 0 Then Exit Function
    ThisWorkbook.Application.ScreenUpdating = False
    isReady = True
    sVerb = "GET"
    If cmCode = 1 Then
        strx = ThisWorkbook.Application.EncodeUrl(strText) '转码
        Urlx = searchUrl & strx
    Else
        Urlx = strText
    End If
    sResult = HTTP_GetData(sVerb, Urlx)
    If Len(sResult) = 0 Then Exit Function
    WriteHtml sResult
    If cmCode = 1 Then
        arr = searchData '搜索结果
    Else
        Pagexs = 0
        arr = ObtainLists
        Pagexs = ObtainPages
    End If
    ObtainPage_Info = arr
    ThisWorkbook.Application.ScreenUpdating = True
    Set oHtmlDom = Nothing
End Function

Private Function searchData() As String() '处理搜索返回的内容
    Dim oHtml As Object
    Dim oA As Object
    Dim i As Integer, k As Integer
    Dim arr() As String
    Dim oCrumbs As Object
    
    On Error GoTo ErrHandle
    With oHtmlDom
        Set oHtml = .getElementById("searchText").Children(1) '此站点的搜索结果存放的位置
        If oHtml.Children.Length = 0 Then GoTo ErrHandle '没有搜索到内容
        Set oA = oHtml.getElementsByTagName("a") '超链接
        k = oA.Length - 1
        ReDim arr(k, 1)
        ReDim searchData(k, 1)
        For i = 0 To k
            arr(i, 0) = oA(i).Text '搜索结果的显示内容
            arr(i, 1) = oA(i).href '结果对应的超链接 'https://www.w3school.com.cn/tags/att_a_href.asp
        Next
    End With
    searchData = arr
ErrHandle:
    Set oHtmlDom = Nothing
    Set oHtml = Nothing
    Set oA = Nothing
End Function

Private Function ObtainLists() As String() '返回页面章节目录
    Dim oHtml As Object
    Dim item As Object
    Dim i As Integer, j As Integer
    Dim oA As Object
    Dim arr() As String
    
    Set oHtml = oHtmlDom.getElementsByTagName("li") '获取列表信息
    j = oHtml.Length
    If j > 0 Then '获取到有效的信息
        ReDim arr(j - 1, 1)
        ReDim ObtainLists(j - 1, 1)
        j = 0
        For Each item In oHtml
            Set oA = item.getElementsByTagName("a")
            i = oA.Length
            If i > 0 Then
                i = i - 1
                arr(j, 0) = oA(i).Text '搜索结果的显示内容
                arr(j, 1) = oA(i).href '结果对应的超链接
                j = j + 1
            End If
        Next
    End If
    ObtainLists = arr
    Set oA = Nothing
    Set oHtml = Nothing
End Function

Private Function ObtainPages() As Integer '获取总共有多少页
    Dim oHtml As Object
    Dim i As Integer, j As Integer
    Dim oA As Object
    Dim strTemp As String
    
    Set oHtml = oHtmlDom.getElementsByclassName("pagebar") '获取页码
    i = oHtml.Length
    If i > 0 Then
        i = i - 1
        j = oHtml.item(i).Children.Length - 1
        strTemp = oHtml.item(i).Children.item(j).href '最后页面的链接
        strTemp = Left(strTemp, Len(strTemp) - 1)
        strTemp = Right(strTemp, Len(strTemp) - InStrRev(strTemp, "/"))
        j = Int(strTemp)
    Else
        j = -1
    End If
    ObtainPages = j
    Set oHtml = Nothing
    Set oA = Nothing
End Function

Private Function MoreDetail() As String() '获取更多书籍的信息
    Dim oHtml As Object
    Dim item As Object
    Dim oA As Object
    Dim arr(4) As String
    
    With oHtmlDom
        For Each item In oHtmlDom.all
            arr(0) = item.Children.item(0).Children.item("description").Content: Exit For '内容简介/信息位于第一个item
        Next
        Set oHtml = .getElementsByclassName("bookinfo")
        Set oA = oHtml.getElementsByTagName("em")
        arr(1) = oA.innertext '--------------------------作者
        Set oHtml = .getElementsByclassName("stats")
        Set oA = oHtml.getElementsByTagName("a")
        arr(2) = oA(0).href '---------------最新章节链接
        arr(3) = oA(0).innertext '----------章节名称
        Set oHtml = .getElementsByclassName("intro") '封面图片
        Set oA = oHtml.getElementsByTagName("img")
        arr(4) = oA(0).href
    End With
    ReDim MoreDetail(4)
    MoreDetail = arr
    Set oA = Nothing
    Set oHtml = Nothing
End Function
'------------------------------------获取小说

'-----------------------------------------------------金山/有道词典-翻译
Function sGet_Translation_Youdao(ByVal strText As String, Optional ByVal iType As Byte = 0) As String
    Dim strType As String
    Dim strPost As String
    Dim sResult As String
    Dim strTemp As String
    Dim xError As Integer
    Dim js As Object
    Dim i As Integer, k As Integer             '---------'"http://fanyi.youdao.com/translate_o?smartresult=dict&smartresult=rule" '如果需要使用此接口, 需要规避有道词典的反爬虫限制
    '---------------------------------m.youdao.com没有反爬机制,但是没有数据接口,需要解析html来获取信息
    Const Urlx As String = "http://fanyi.youdao.com/translate?smartresult=dict&smartresult=rule&smartresult=ugc&sessionFrom=null"
    Const cUrl As String = "http://fanyi.youdao.com/"
    '有道sign的计算方法
    'tsStr=timestamp '时间戳13位
    'saltStr=tsStr & randnumx(9) '加上一位(0-9)的随机数-14位
    '常量aStr="fanyideskweb"
    '要翻译的内容 strText
    '另一个变化很慢的常量 Youdao_Const
    'sign=getmd5hash_string(astr & strText & salt & Youdao_Const) '等到32位的hash值,就是有道的sign
    '--------------------------------------------------------------------------------------------------但是却还是用不了,还是返回50 Error
    strType = "&type=AUTO" '& Translate_Type(iType) '有道的翻译类型经常出现问题(只能选auto了)
    strPost = "&i=" & Replace(Replace(Replace(oEncodeUrl(strText), "%", "%C2%"), "%C2%E", "%C3%A"), "%C2%20", " ") & strType  ' 'Application.EncodeURL(strText)
    strPost = strPost & "&doctype=json"
    DeleteUrlCacheEntry cUrl
    sResult = HTTP_GetData("POST", Urlx, refUrl:=cUrl, cType:="application/x-www-form-urlencoded; charset=UTF-8", acType:="application/json, text/javascript, */*; q=0.01", _
    cHost:="fanyi.youdao.com", oRig:=cUrl, sPostdata:=strPost)
    If Len(sResult) = 0 Then Exit Function
    Set js = CreateObject("scriptcontrol")
    With js
        .Language = "jscript"
        .addcode "HLA=" & sResult
        Debug.Print sResult
        xError = .eval("HLA.errorCode")
        If xError = 50 Then Set js = Nothing: Exit Function '获取执行的错误代码(50为无返回值)
        k = .eval("HLA.translateResult[0].length")
        For i = 1 To k
            strTemp = strTemp & .eval("HLA.translateResult[0][" & i - 1 & "].tgt")
            '------------------------------{"type":"FR2ZH_CN","errorCode":0,"elapsedTime":2,"translateResult":[[{"src":"Garde contre les uv","tgt":"紫外警告"}]]}
        Next
    End With
    sGet_Translation_Youdao = strTemp
    Set js = Nothing
 End Function
'有道的常量的变化很小， bv的值并不是指定,可以随意指定一个md5值
Function Get_Youdao_Translate(ByVal strText As String, Optional ByVal iType As Byte) As String
    Const cAssist As String = "&bv=3aabbc1a31e864bb89725aa04c217a5c&doctype=json&version=2.1&keyfrom=fanyi.web&action=FY_BY_REALTlME"
    Const aSign As String = "fanyideskweb"
    Const sAssist As String = "&smartresult=dict&client=fanyideskweb&salt="
    Const cSign As String = "Nw(nmmbP%A-r6U3EUn]Aj"
    Const rUrl As String = "http://fanyi.youdao.com/"
    Const sUrl As String = "http://fanyi.youdao.com/translate_o?smartresult=dict&smartresult=rule"
    Const cCookie As String = "OUTFOX_SEARCH_USER_ID=325673277@27.38.20.19"
    Dim sResult As String
    Dim sign As String
    Dim Salt As String
    Dim sTime As String
    Dim strTemp As String
    Dim eStr As String
    Dim PostData As String
    Dim sFrom As String
    Dim sTo As String
    Dim x
    Dim cReg As New cRegex
    
    x = Split(Translate_Type(True, iType), "2", , vbBinaryCompare)
    sFrom = x(0)
    sTo = x(1)
    eStr = ThisWorkbook.Application.EncodeUrl(strText) '对搜索内容进行编码
    sTime = Get_Timestamp '时间戳 '---尽量让时间戳的时间延后生成
    Salt = sTime & Random_Ten '时间戳+0-9的随机数
    strTemp = aSign & strText & Salt & cSign
    sign = LCase(GetMD5Hash_String(strTemp)) '生成签名
    '------------------------------------------------------破解有道的反爬
    PostData = "i=" & eStr & "&from=" & sFrom & "&to=" & sTo & sAssist & Salt & "&sign=" & sign & "&ts=" & sTime & cAssist
    isReady = True
    sResult = HTTP_GetData("POST", sUrl, rUrl, sCookie:=cCookie, sPostdata:=PostData)
    If isReady = False Then Exit Function
    strTemp = cReg.sMatch(sResult, Chr(34) & "errorCode" & Chr(34) & ":\d{1,}")
    If Split(strTemp, ":")(1) = "50" Then Set cReg = Nothing: Exit Function '50表示返回的数据有误
    strTemp = cReg.sMatch(sResult, Chr(34) & "tgt" & Chr(34) & ":.*?,")
    strTemp = Split(strTemp, ":", , vbBinaryCompare)(1)
    strTemp = Mid$(strTemp, 2, Len(strTemp) - 3)
    Get_Youdao_Translate = strTemp
    Set cReg = Nothing
End Function

Private Function Translate_Type(ByVal IorY As Boolean, Optional ByVal iType As Byte = 0, Optional ByVal nType As Byte = 0) As String '选择翻译的类型
    Dim sType As String
    Dim fType As String
    Dim tType As String
    If IorY = True Then
        Select Case iType
            Case 1: sType = "en2zh-CHS"
            Case 2: sType = "zh-CHS2en"
            Case 3: sType = "ja2zh-CHS"
            Case 4: sType = "zh-CHS2ja" '"&from=zh-CHS&to=ja" '
            Case Else: sType = "AUTO2AUTO" '让网页自动判别语言的类型
        End Select
    Else
        Select Case iType
            Case 1: fType = "zh"
            Case 2: fType = "en"
            Case 3: fType = "ja"
            Case 4: fType = "ko" '"&from=zh-CHS&to=ja" '
            Case 5: fType = "fr"
            Case Else: fType = "auto": iType = 0 '让网页自动判别语言的类型
        End Select
        If iType = 0 Then
            tType = "auto"
        ElseIf iType = 1 Then
            Select Case nType
                Case 1: tType = "en"
                Case 2: tType = "ja"
                Case 3: tType = "ko"
                Case 4: tType = "kr"
                Case 5: tType = "fr"
                Case Else: tType = "en"
            End Select
        Else
            tType = "zh"
        End If
        sType = "&f=" & fType & "&t=" & tType
    End If
    Translate_Type = sType
End Function

Private Function Youdao_Const() As String '获取用于计算有道sign的常量,该数据位于min.js当中
    Dim sResult As String
    Dim oReg As Object
    Dim Matches As Object
    Dim match As Object
    Const tUrl As String = "http://shared.ydstatic.com/fanyi/newweb/v1.0.25/scripts/newweb/fanyi.min.js"
    
    sResult = HTTP_GetData("GET", tUrl, "http://www.youdao.com/")
    Set oReg = CreateObject("VBScript.RegExp")
    With oReg
        '--------https://www.cnblogs.com/shuai1993/p/10235577.html
        '--------------http://blog.sina.com.cn/s/blog_16698a9a40102zwtb.html
        '----------------------------------匹配此类型的数据n.md5("fanyideskweb"+e+i+"Nw(nmmbP%A-r6U3EUn]Aj")}}
        .Pattern = "n\.md5\(\" & Chr(34) & ".*\)\}\}\;" '符号\转意
        .Global = True '不区分大小写
        .IgnoreCase = True
        Set Matches = .Execute(sResult)
        For Each match In Matches
            sResult = match.Value: Exit For
        Next
    End With
    Set oReg = Nothing
    Set Matches = Nothing
    Youdao_Const = Split(sResult, Chr(34))(3)
End Function
'----------------优先使用金山词霸作为翻译的首选,金山词霸没有有道词典这么多的问题,支持多语种,不会因为符号等问题造成数据无法正常传输等-Google也可以(有个没有堵上的漏洞)
'----------------------------------------------------------------------google translation
'source_code_name:[{code:'auto',name:'检测语言'},{code:'sq',name:'阿尔巴尼亚语'},{code:'ar',name:'阿拉伯语'},{code:'am',name:'阿姆哈拉语'},{code:'az',name:'阿塞拜疆语'},{code:'ga',name:'爱尔兰语'},{code:'et',name:'爱沙尼亚语'},
'{code:'or',name:'奥里亚语(奥里亚文)'},{code:'eu',name:'巴斯克语'},{code:'be',name:'白俄罗斯语'},{code:'bg',name:'保加利亚语'},{code:'is',name:'冰岛语'},{code:'pl',name:'波兰语'},{code:'bs',name:'波斯尼亚语'},
'{code:'fa',name:'波斯语'},{code:'af',name:'布尔语(南非荷兰语)'},{code:'tt',name:'鞑靼语'},{code:'da',name:'丹麦语'},{code:'de',name:'德语'},{code:'ru',name:'俄语'},{code:'fr',name:'法语'},{code:'tl',name:'菲律宾语'},
'{code:'fi',name:'芬兰语'},{code:'fy',name:'弗里西语'},{code:'km',name:'高棉语'},{code:'ka',name:'格鲁吉亚语'},{code:'gu',name:'古吉拉特语'},{code:'kk',name:'哈萨克语'},{code:'ht',name:'海地克里奥尔语'},{code:'ko',name:'韩语'},
'{code:'ha',name:'豪萨语'},{code:'nl',name:'荷兰语'},{code:'ky',name:'吉尔吉斯语'},{code:'gl',name:'加利西亚语'},{code:'ca',name:'加泰罗尼亚语'},{code:'cs',name:'捷克语'},{code:'kn',name:'卡纳达语'},{code:'co',name:'科西嘉语'},
'{code:'hr',name:'克罗地亚语'},{code:'ku',name:'库尔德语'},{code:'la',name:'拉丁语'},{code:'lv',name:'拉脱维亚语'},{code:'lo',name:'老挝语'},{code:'lt',name:'立陶宛语'},{code:'lb',name:'卢森堡语'},{code:'rw',name:'卢旺达语'},
'{code:'ro',name:'罗马尼亚语'},{code:'mg',name:'马尔加什语'},{code:'mt',name:'马耳他语'},{code:'mr',name:'马拉地语'},{code:'ml',name:'马拉雅拉姆语'},{code:'ms',name:'马来语'},{code:'mk',name:'马其顿语'},{code:'mi',name:'毛利语'},
'{code:'mn',name:'蒙古语'},{code:'bn',name:'孟加拉语'},{code:'my',name:'缅甸语'},{code:'hmn',name:'苗语'},{code:'xh',name:'南非科萨语'},{code:'zu',name:'南非祖鲁语'},{code:'ne',name:'尼泊尔语'},{code:'no',name:'挪威语'},
'{code:'pa',name:'旁遮普语'},{code:'pt',name:'葡萄牙语'},{code:'ps',name:'普什图语'},{code:'ny',name:'齐切瓦语'},{code:'ja',name:'日语'},{code:'sv',name:'瑞典语'},{code:'sm',name:'萨摩亚语'},{code:'sr',name:'塞尔维亚语'},
'{code:'st',name:'塞索托语'},{code:'si',name:'僧伽罗语'},{code:'eo',name:'世界语'},{code:'sk',name:'斯洛伐克语'},{code:'sl',name:'斯洛文尼亚语'},{code:'sw',name:'斯瓦希里语'},{code:'gd',name:'苏格兰盖尔语'},
'{code:'ceb',name:'宿务语'},{code:'so',name:'索马里语'},{code:'tg',name:'塔吉克语'},{code:'te',name:'泰卢固语'},{code:'ta',name:'泰米尔语'},{code:'th',name:'泰语'},{code:'tr',name:'土耳其语'},{code:'tk',name:'土库曼语'},
'{code:'cy',name:'威尔士语'},{code:'ug',name:'维吾尔语'},{code:'ur',name:'乌尔都语'},{code:'uk',name:'乌克兰语'},{code:'uz',name:'乌兹别克语'},{code:'es',name:'西班牙语'},{code:'iw',name:'希伯来语'},{code:'el',name:'希腊语'},
'{code:'haw',name:'夏威夷语'},{code:'sd',name:'信德语'},{code:'hu',name:'匈牙利语'},{code:'sn',name:'修纳语'},{code:'hy',name:'亚美尼亚语'},{code:'ig',name:'伊博语'},{code:'it',name:'意大利语'},{code:'yi',name:'意第绪语'},
'{code:'hi',name:'印地语'},{code:'su',name:'印尼巽他语'},{code:'id',name:'印尼语'},{code:'jw',name:'印尼爪哇语'},{code:'en',name:'英语'},{code:'yo',name:'约鲁巴语'},{code:'vi',name:'越南语'},{code:'zh-CN',name:'中文'}]

Function Google_Translation(ByVal strText As String) As String '用上面的代码替换掉(sl=auto)(要翻译的内容) auto or (tl=en)en(翻译的目标语言)
    Const tUrl As String = "http://translate.google.cn/translate_a/single?client=gtx&dt=t&ie=UTF-8&oe=UTF-8&sl=auto&tl=zh-CN&q="
    Dim sResult As String
    Dim strTemp As String
    Dim xResult As Variant
    Dim i As Integer, k As Integer
    
    If Len(strText) > 1024 Then MsgShow "内容长度超出范围", "Tips", 1200: Exit Function
    sResult = HTTP_GetData("GET", tUrl & ThisWorkbook.Application.EncodeUrl(strText), "http://translate.google.cn/")
    If Len(sResult) = 0 Then Exit Function
    xResult = Split(sResult, ",[")
    i = UBound(xResult) - 3
    For k = 0 To i
        If k Mod 2 = 0 Then
            If InStr(xResult(k), Chr(34)) > 0 Then strTemp = strTemp & Split(xResult(k), Chr(34))(1)
        End If
    Next
    Google_Translation = strTemp
End Function
'-------------------------使用TK的方法太慢,需要多次请求网络,而且需要计算值,执行速度偏慢
Function Get_Google_Translation_TK(ByVal strText As String) '获取TK值
    Dim sResult As String
    Dim sTKK As String
    Dim aTkk As String
    Dim bTkk As String
    Dim oStream As Object
    Dim oHtml As Object
    Dim FilePath As String
    
    sTKK = Get_Google_Translation_TKK
    aTkk = Split(sTKK, ".")(0)
    bTkk = Split(sTKK, ".")(1)
    Set regEx = CreateObject("VBScript.RegExp")
    FilePath = ThisWorkbook.Path & "\googletks.html"
    Set oStream = New ADODB.Stream
    With oStream
        .Mode = adModeReadWrite
        .type = adTypeText
        .CharSet = "gb2312"
        .Open
        .LoadFromFile FilePath
        sResult = .ReadText()
        sResult = ReplaceText(sResult, "[var b=]+([1-9]{6,6})[;]", "var b = " & aTkk & ";")
        sResult = ReplaceText(sResult, "[var b1=]+([1-9]{9,9})[;]", "var b1 = " & bTkk & ";")
        .Position = 0 '必须使用position以实现对文本的覆盖写入
        .WriteText sResult
        .Flush
        .SaveToFile FilePath, 2 '将数据重新写入进去
        .Close
    End With
    Set oStream = Nothing
    Set regEx = Nothing
    '---------------------https://www.cnblogs.com/qinshou/p/5932274.html
    '需要注意的是,需要在htmlfile上注释上:!-- saved from url=(0013)about:internet --,否则会出现警告提示
    Set oHtml = CreateObject(FilePath)
    '如果将内容写入htmlfile,无法调用js函数
    Get_Google_Translation_TK = CallByName(oHtml.parentwindow, "TKValue", VbMethod, strText)
    Set oHtml = Nothing
End Function

Private Function Get_Google_Translation_TKK() As String
    Dim iCookie As String
    Dim asType As String
    Dim item As Object
    Dim sResult As String
    asType = "text/html, application/xhtml+xml, */*"
    iCookie = "NID=202=L52Shg2Vi_YUD4QM8DYSbBfsoF0knvnllb5uxmkmHOLv_h4tfNOASO_Eef-Kw6wKg4_MJe0TpwhWlh42242Be-Rv9NEKpL_IlYh0RPdYhoAniG0OSVo_Hn8CjKcN53OZ2XKWGM_O8M5GiQyjVYZF6Bycp8ktOEmz7M5qBY6ITkM"
    sResult = HTTP_GetData("GET", "https://translate.google.cn/", "https://translate.google.cn/", acType:=asType, cHost:="translate.google.cn", sCookie:=iCookie)
    WriteHtml sResult
    For Each item In oHtmlDom.all
        sResult = item.innertext: Exit For '数据位于第一个item中
    Next
    Set oHtmlDom = Nothing
    Get_Google_Translation_TKK = Google_Tkk(sResult)
End Function

Private Function Google_Tkk(ByVal strText As String) As String '获取google_translation的tkk值
    Dim oReg As Object
    Dim Matches As Object
    Dim match As Object
    Set oReg = CreateObject("VBScript.RegExp")
    With oReg
        .Pattern = "([0-9]{6,})+(\.[0-9]{9,})" '匹配长度为6位数+小数点+9位数的组合,\转意, \. 表示标点符号"."
        .Global = True '不区分大小写
        .IgnoreCase = True
        Set Matches = .Execute(strText)
        For Each match In Matches
            Google_Tkk = match.Value: Exit For
        Next
    End With
    Set oReg = Nothing
    Set Matches = Nothing
End Function
'由于要分开请求的内容用起来较为麻烦,不建议过长的内容
'total //标准为100个字符为一组 然后total是总共多少组
'idx//下标,第几组
'textlen //所有文字的计数,就是这个字符串的长度
'q //合成的文字内容,中文的话需要转编码utf-8最为合适
'tl //语言，跟识别的语言一致
Sub Online_Voice_Google(ByVal strText As String) '有长度的限制
    Dim i As Byte, j As Byte, k As Integer
    Dim tUrl As String, strText As String
    Const gUrl As String = "https://translate.google.cn/translate_tts?ie=UTF-8&client=tw-ob&ttsspeed=1"
    i = 0
    j = 1
    k = Len(strText)
    If k > 144 Then MsgShow "长度超出范围", "Tips", 1200: Exit Sub
    tUrl = gUrl & "&total=" & CStr(i) & "&idx=" & CStr(j) & "&textlen=" & CStr(k) & "&q=" & ThisWorkbook.Application.EncodeUrl(strText) & "&tl=en"
    HTTP_GetData "GET", strx, "https://translate.google.cn/", IsSave:=True
    isReady = True
    HTTP_GetData "GET", tUrl, IsSave:=True
    If isReady = False Then Exit Sub
    Dim cwm As New cWMP
    cwm.wmOpen ThisWorkbook.Path & "\voice.mp3"
    cwm.wmPlay
    Do Until cwm.Status = "stopped"
        DoEvents
    Loop
    Set cwm = Nothing
End Sub

Sub Online_Voice(ByVal strText As String, ByVal iType As Byte, Optional ByVal vYoudao As Boolean = False, Optional IsUK As Boolean = False) '中英日/ 美式英式
    Const usUrl As String = "https://fanyi.baidu.com/gettts?lan=en&text=" '语音-百度
    Const ukUrl As String = "https://fanyi.baidu.com/gettts?lan=uk&text="
    Const zhUrl As String = "https://fanyi.baidu.com/gettts?lan=zh&text="
    Const jaUrl As String = "https://fanyi.baidu.com/gettts?lan=jp&text="
    '----------------baidu
    Const yUrl As String = "http://dict.youdao.com/dictvoice?audio=" '有道
    Dim tUrl As String
    Dim strx As String
    
    strText = Trim(strText)
    If Len(strText) = 0 Then Exit Sub
    strText = ThisWorkbook.Application.EncodeUrl(strText)
    If vYoudao = False Then '优先选择百度作为发音源
        Select Case iType
            Case 1: tUrl = zhUrl
            Case 2: tUrl = jsurl
            Case Else:
            tUrl = IIf(IsUK = True, ukUrl, usUrl)
        End Select
        tUrl = tUrl & strText & "&spd=3&source=web" 'baidu
    Else
        Select Case iType
            Case 1: strx = "&le=zh"
            Case 2: strx = "&le=ja"
            Case Else:
            strx = IIf(IsUK = True, "1", "2")
            strx = "&type" & strx
            strText = Replace(strText, "%20", "+")
        End Select
        tUrl = yUrl & strxtext & strx
    End If
    isReady = True
    HTTP_GetData "GET", tUrl, IsSave:=True
    If isReady = False Then Exit Sub
    Dim cwm As New cWMP
    cwm.wmOpen ThisWorkbook.Path & "\voice.mp3"
    cwm.wmPlay
    Do Until cwm.Status = "stopped"
        DoEvents
    Loop
    Set cwm = Nothing
End Sub

'url为直接转码 , 对于特定字符不作处理
'https://fanyi.qq.com/api/tts?platform=PC_Website&lang=zh&text=%E5%A5%B9%E6%98%AF%E4%B8%AA%E5%B0%8F%E7%BE%8E%E5%A5%B3&guid=700b30a9-8bef-4f3f-9b13-238ca7f51f9a
'后面只使用guid进行限制, guid来源于cookie
Sub Get_Voice_fromTencent(ByVal strText As String, ByVal cCookie As String, Optional ByVal iType As Boolean) '从腾讯翻译君获取语音
    Const rUrl As String = "https://fanyi.qq.com/"
    Const vUrl As String = "https://fanyi.qq.com/api/tts?platform=PC_Website&lang"
    Dim sUrl As String
    Dim cReg As New cRegex
    Dim sGuid As String
    If iType = False Then
        sUrl = vUrl & "zh&text="
    Else
        sUrl = vUrl & "en&text="
    End If
    sGuid = cReg.sMatch(cCookie, "guid.*?;")
    If Len(sGuid) = 0 Then Exit Sub
    sGuid = Split("=")(0)
    sGuid = Replace$(sGuid, ";", "", 1, , vbBinaryCompare)
    sUrl = sUrl & ThisWorkbook.Application.EncodeUrl(strText) & "&guid" & sGuid
    HTTP_GetData "GET", sUrl, rUrl, IsSave:=True, sCookie:=cCookie
    Set cReg = Nothing
End Sub

Function Get_Translation_iCiba(ByVal strText As String, Optional ByVal iType As Byte = 0, Optional ByVal nType As Byte = 0) As String
    Dim oJS As Object
    Dim sPost As String
    Dim xError As Integer, k As Integer
    Dim sResult As String
    Dim strType As String, strTemp As String
    Const xUrl As String = "http://fy.iciba.com/ajax.php?a=fy"
    Const ref As String = "http://www.iciba.com"
    Const rType As String = "application/json, text/javascript, */*; q=0.01"
    
    If InStr(strText, "\u") > 0 Then MsgBox "请检查输入的内容是否存在有误", vbInformation, "Tips": Exit Function
    isReady = True
    strType = Translate_Type(False, iType, nType) & "&w="
    sPost = strType & ThisWorkbook.Application.EncodeUrl(strText)
    sResult = HTTP_GetData("POST", xUrl, ref, acType:=rType, sPostdata:=sPost) '"&f=zh&t=ja&w="
    If isReady = False Then Exit Function
    Set oJS = CreateObject("scriptcontrol")
    With oJS
        .Language = "jscript"
        .addcode "HLA=" & sResult
        xError = .eval("HLA.content.err_no") '错误代码
        If xError = 50 Then Set oJS = Nothing: Exit Function
        strTemp = Trim(.eval("HLA.content.out"))
        If Left$(strTemp, 2) = "\u" Then strTemp = Unicode2Character(strTemp) '返回的数据有unicode字符（如：日文）/明文返回
    End With
    Set oJS = Nothing
    Get_Translation_iCiba = strTemp
End Function
'-----------------------------------------------------------------------------翻译

'------------------------------------------------集思录ETF
Sub GetETF_Lists() '获取ETF列表信息
    Dim Urlx As String
    Dim sResult As String
    Dim sVerb As String
    Dim strT As String
    Const tUrl As String = "https://www.jisilu.cn/data/etf/etf_list/?___jsl=LST___t="
    
    DisEvents
    sVerb = "GET"
    strT = TimeStamp '时间戳
    Urlx = tUrl & strT & "&rp=25&page=1"
    sResult = HTTP_GetData(sVerb, Urlx)
    ETF_Lists sResult, strT
    Set oHtmlDom = Nothing
    EnEvents
End Sub

Private Function cETF_Lists(ByVal strText As String) As String '通过类模块来实现ETF信息获取
    Dim cjs As New cJSON
    Dim objdic As Object
    Dim objrow As Object
    Dim objcell As Object
    Dim objtemp As Object
    Dim k As Integer, i As Integer, p As Integer
    Dim item
    Dim arr() As String
    '如果使用jsonconvertor的话, 和cjson模块的使用上有所区别,在获取具体的值,jsonconvertor无法直接返回对应的值,但是可以直接返回值的数组(variant)w = b.Items
    Set objdic = cjs.parse(strText)
    Set objrow = objdic("rows")
    k = objrow.Count
    If k = 0 Then Exit Sub
    For i = 1 To k
        Set objtemp = objrow(i)
        Set objcell = objtemp("cell")
        If i = 1 Then p = objcell.Count: ReDim arr(1 To k, 1 To p): ReDim cETF_Lists(1 To k, 1 To p): p = 1
        For Each item In objcell
            arr(i, p) = objcell(item)
            p = p + 1
        Next
        p = 1
    Next
    Set objrow = Nothing
    Set objcell = Nothing
    Set objtemp = Nothing
    Set objdic = Nothing
    Set cjs = Nothing
End Function

Private Sub ETF_Lists(ByVal strText As String, ByVal filen As String) '爬取集思录上的ETF信息
    Dim objSc As Object
    Dim strTemp As String
    Dim objItem As Object
    Dim objJson As Object
    Dim arrFund() As Variant
    Dim p As Integer
    Dim i As Integer, k As Integer, j As Byte, objtemp As Object, objcell As Object, arrt() As String
    Dim objinfo As Object
    Dim wb As Workbook, FilePath As String, arrTemp() As Variant
    Dim im As Integer
    
    Set regEx = CreateObject("VBScript.RegExp")             ' 建立正则表达式。
    Set objSc = CreateObject("ScriptControl")
    objSc.Language = "JScript"
    strTemp = ReplaceText(strText, "cell", "icell") '替换掉有冲突的关键字
    strTemp = ReplaceText(strTemp, "rows", "irows")
    strTemp = ReplaceText(strTemp, "page", "ipage")
    strTemp = ReplaceText(strTemp, "null", "-")
    Set objJson = objSc.eval("s=" & strTemp)
    Set objtemp = objJson.irows
    '----------https://docs.microsoft.com/en-us/previous-versions/cc175542(v=vs.90)
    '----------https://docs.microsoft.com/en-us/previous-versions/cc177247%28v%3dvs.90%29
    Set oTli = CreateObject("TLI.TLIApplication")
    '-----------------------此功能不直接适用于x64,需要手动注册dll
    Set objinfo = oTli.InterfaceInfoFromObject(objtemp)
    im = objinfo.Members.Count
    '------------将数据写入到新的工作簿中
    arrTemp = Array("代码", "名称", "现价", "涨幅(%)", "成交额(万元)", "指数", "指数PE", "指数PB", "指数涨幅(%)", "估值", "净值", "净值日期", "溢价率(%)", "最小申赎(万份)", "管托费(%)", "份额(万份)", "规模变化(亿元)", "规模(亿元)", "基金公司", "Index ID")
    Set wb = Workbooks.Add
    With wb
        .Sheets(1).Name = "ETFLists"
        With .Sheets(1)
            .Cells(3, 1).Resize(1, 20) = arrTemp '表头
            .Range("a3:t3").HorizontalAlignment = xlCenter '居中
            .Range("a3:t3").Font.Bold = True
            '------------------添加超链接
            .Hyperlinks.Add Anchor:=.Cells(1, 2), address:= _
            "https://www.jisilu.cn/data/etf/", TextToDisplay:="集思录"
            .Cells(1, 1) = "数据来源:"
            .Cells(1, 1).Font.Bold = True
            .Cells(2, 1) = "数据更新时间:"
            .Cells(2, 1).Font.Bold = True
            .Cells(2, 2) = filen
            .Cells(2, 2).NumberFormatLocal = "000000"
            .Cells(2, 3) = "(时间戳)"
            For Each objItem In objJson.irows
                Set objcell = objItem.icell
                If j = 0 Then arrt = ObtainObjInfo(objcell): j = 1: k = UBound(arrt): ReDim arrFund(1 To im, 1 To k): i = 1 '数组不从0开始,方便后续将数据直接放进表格
                For p = 1 To k
                    If p <> 13 And p <> 14 And p <> 17 And p <> 25 Then
                        '-------'防止出现null的情况
                        'If IsNull(CallByName(objcell, arrT(p), VbGet)) = False Then arrFund(i, p) = CallByName(objcell, arrT(p), VbGet)
                        arrFund(i, p) = CallByName(objcell, arrt(p), VbGet)
                        '-----------------https://docs.microsoft.com/zh-cn/office/vba/Language/Reference/user-interface-help/callbyname-function
                    End If
                Next
                i = i + 1
            Next
            '----------------将数据重新进行排列,调整对应位置
            For j = 1 To i
                k = ColumnChoice(j)
                If k > 0 Then
                    '-----------------https://docs.microsoft.com/zh-cn/office/vba/api/Excel.WorksheetFunction.Index
                    arrTemp = ThisWorkbook.Application.Index(arrFund, , j) '将数组的部分(某一列)单独提取写入表格
                    WriteList wb, k, im, arrTemp
                End If
            Next
            For j = 1 To im '将基金的链接格式调整为超链接/显示基金公司的名称
                .Hyperlinks.Add Anchor:=.Cells(j + 3, "s"), address:= _
                arrFund(j, 7), TextToDisplay:=arrFund(j, 6)
            Next
            '----------------------调整显示数字的格式
            im = im + 3
            .Range("e4:e" & im).NumberFormatLocal = "0.00_ "
            .Range("d4:d" & im).NumberFormatLocal = "0.00_ "
            .Range("c4:c" & im).NumberFormatLocal = "0.000_ "
            .Range("g4:h" & im).NumberFormatLocal = "0.000_ "
            .Range("i4:i" & im).NumberFormatLocal = "0.00_ "
            .Range("j4:k" & im).NumberFormatLocal = "0.0000_ "
            .Range("m4:m" & im).NumberFormatLocal = "0.00_ "
            .Range("o4:o" & im).NumberFormatLocal = "0.00_ "
            .Range("q4:r" & im).NumberFormatLocal = "0.00_ "
            .Range("t4:t" & im).NumberFormatLocal = "000000"
            .Columns.AutoFit '调整列宽
        End With
        FilePath = ThisWorkbook.Path & "\" & "ETF" & filen & ".xlsx"
        .SaveAs FilePath
    End With
    Set objinfo = Nothing
    Set objtemp = Nothing
    Set objcell = Nothing
    Set objSc = Nothing
    Set regEx = Nothing
    Set wb = Nothing
    Erase arrFund
End Sub

Private Function ColumnChoice(ByVal intx As Byte) As Byte '将数据放入对应的列
    Dim i As Byte
    Select Case intx
        Case 1: i = 1
        Case 2: i = 2
        Case 3: i = 20
        Case 4: i = 15
        Case 5: i = 14
        Case 8: i = 16
        Case 9: i = 18
        Case 10: i = 17
        Case 11: i = 3
        Case 12: i = 5
        Case 15: i = 4
        Case 16: i = 10
        Case 18: i = 13
        Case 19: i = 11
        Case 20: i = 12
        Case 21: i = 6
        Case 22: i = 9
        Case 23: i = 7
        Case 24: i = 8
        Case Else: i = 0
    End Select
    ColumnChoice = i
End Function

Private Sub WriteList(ByVal wbx As Workbook, ByVal Indexi As Byte, ByVal p As Integer, ByRef arrx()) '将数据写入表格
    With wbx.Sheets(1)
        .Cells(4, Indexi).Resize(p, 1) = arrx
    End With
End Sub

Private Function ObtainObjInfo(ByVal objx As Object) As String() '获取对象属性名称
    Dim i As Integer, k As Integer
    Dim arr() As String
    Dim objinfo As Object
    
    Set objinfo = oTli.InterfaceInfoFromObject(objx)
    k = objinfo.Members.Count
    ReDim arr(1 To k)
    ReDim ObtainObjInfo(1 To k)
    With objinfo
        For i = 1 To k
            arr(i) = .Members(i).Name
        Next
    End With
    ObtainObjInfo = arr
    Set objinfo = Nothing
End Function

Private Function ReplaceText(ByVal strText As String, ByVal strFind As String, ByVal rpText As String) As String '正则替换
    With regEx
        .Global = True '需要开启全局,才能将全部找到的值替换掉
        .Pattern = strFind
        .IgnoreCase = True
        ReplaceText = .Replace(strText, rpText)
    End With
End Function
'---------------------------------------------------------------------------------------------------------------------集思录ETF

'-------------------------------------------------------------------------------------天气
Function Code_City_Weather(ByVal cName As String) As String() '获取天气搜索列表
    Dim Urlx As String
    Dim sResult As String
    Dim xV As Variant, xva As Variant
    Dim strTemp As String
    Dim arr() As String
    Dim i As Integer, k As Integer
    '--------------------------------得到搜索的列表的城市的地名和对应的编号,返回的数据较少,可以用简单的文本处理方式来实现数据的提取
    strTemp = ThisWorkbook.Application.EncodeUrl(cName)
    Urlx = "http://toy1.weather.com.cn/search?cityname=" & strTemp & "&callback=success_jsonpCallback&_=" & TimeStamp
    sResult = HTTP_GetData("GET", Urlx)
    strTemp = Replace$(sResult, "success_jsonpCallback([", "")
    sResult = Trim$(Replace$(strTemp, "])", ""))
    If Len(sResult) = 0 Then Exit Function
    xV = Split(sResult, ",")
    k = UBound(xV)
    ReDim arr(k, 1)
    ReDim Code_City_Weather(k, 2)
    If k = 0 Then Exit Function
    For i = 0 To k
        strTemp = xV(i)
        xva = Split(strTemp, "~")
        strTemp = xva(0) '所在地区编码
        arr(i, 0) = Trim$(Right$(strTemp, Len(strTemp) - 8))
        arr(i, 1) = Trim$(xva(2)) '所在地名称
        strTemp = xva(UBound(xva))
        arr(i, 2) = Left$(strTemp, Len(strTemp) - 1) '所在地所属区域
    Next
    Code_City_Weather = arr
End Function
'-----------------------由于中国天气网没有相关方便获取数据的端口,不从这个网站获取数据
'Function Get_Weather(ByVal cityCode As String, Optional ByVal dType As Byte) As String
'Dim urlx As String
'Dim xtype As String
'If dType = 1 Then
'xtype = "weather" '7天
'Else
'xtype = "weather1d" '1天
'End If
'urlx = "http://www.weather.com.cn/" & xtype & "/" & cityCode & ".shtml"
'End Function
'----------------------------------------------------------------------
Function Get_Weather_fromAPI(ByVal cityCode As String) As String() '获取天气的信息 /15天
    Dim Urlx As String
    Dim sResult As String
    Dim cjs As New cJSON
    Dim objdic As Object
    Dim objdata As Object
    Dim objinfo As Object
    Dim objtemp As Object
    Dim item As Variant
    Dim arr() As String
    Dim i As Byte, j As Byte, p As Byte, m As Byte, k As Byte
    Const urlapi As String = "http://t.weather.sojson.com/api/weather/city/"
    '-------------------------------https://www.sojson.com/api/weather.html /2000/day
    Urlx = urlapi & cityCode
    sResult = HTTP_GetData("GET", Urlx)
    Set objdic = cjs.parse(sResult)
    Set objdata = objdic("data")
    Set objinfo = objdata("forecast")
    k = objinfo.Count
    Set objtemp = objinfo(1)
    p = objtemp.Count
    Set objtemp = objinfo(k)
    j = objtemp.Count
    m = IIf(p > j, p, j)
    For i = 1 To k    '比较第一组和最后一组数据 15,12/11
        Set objtemp = objinfo(i)
        j = objtemp.Count
        If i = 1 Then ReDim arr(1 To k, 1 To m): ReDim Get_Weather_fromAPI(1 To k, 1 To m): p = 1 '数据不是完全一致
        For Each item In objtemp
            arr(i, p) = objtemp(item)
            p = p + 1
            If j < m Then
                If p = 8 Then p = 9: arr(i, 8) = "-" '空气质量部分缺失
            End If
        Next
        p = 1
    Next
    Get_Weather_fromAPI = arr
    Set cjs = Nothing
    Set objtemp = Nothing
    Set objdata = Nothing
    Set objinfo = Nothing
    Set objdic = Nothing
    Erase arr
End Function

Function Get_IP_CityCode() As String() '获取所在ip对应的地区的天气id
    Dim sResult As String
    Dim arr() As String
    Dim temp As Variant
    Dim i As Byte, k As Byte
    Const cUrl As String = "http://wgeo.weather.com.cn/ip/?_="
    Const ref As String = "http://www.weather.com.cn/forecast/"
    
    sResult = HTTP_GetData("GET", cUrl & TimeStamp, ref)
    sResult = Left$(sResult, Len(sResult) - 2)
    temp = Split(sResult, ";")
    i = UBound(temp)
    ReDim arr(i)
    ReDim Get_IP_CityCode(i)
    For k = 0 To i
        arr(k) = Trim$(Replace$(Split(temp(k), "=")(1), Chr(34), "")) 'ip地址,城市代码, 所在地区
    Next
    Get_IP_CityCode = arr
End Function
'--------------------------------------------------------------------------------------------------------天气

'---------------------------------豆瓣
'需要注意的是豆瓣的信息爬取,直接爬html的信息会出现反爬的信息
'但是可以直接获取对应id元素的信息,却可以绕过豆瓣的反爬措施
Sub sGet_doubanBook_Tag()
    Dim sTag As Variant
    Dim itemx
    Dim i As Integer, k As Integer
    
    Set sTag = Get_doubanBook_Tag
    For Each itemx In sTag
        Debug.Print itemx   '大的分类
        i = sTag(itemx).Count
        For k = 1 To i
            Debug.Print sTag(itemx)(k) '子Tag
        Next
    Next
    Set sTag = Nothing
End Sub

Function Get_doubanBook_Tag() As Variant '获取豆瓣书籍分类Tag
    Const tUrl As String = "https://book.douban.com/tag/"
    Dim sResult As String
    Dim arr() As String
    Dim i As Integer, j As Integer
    Dim oTag As Object
    Dim oA As Object
    Dim oTitle As Object
    Dim item As Object
    Dim strTemp As String
    Dim dic As Object
    
    sResult = HTTP_GetData("GET", tUrl)
    WriteHtml sResult
    Set oTag = oHtmlDom.getElementsByclassName("article")
    Set oA = oTag(0).getElementsByTagName("a")
    Set oTitle = oTag(0).getElementsByclassName("tag-title-wrapper")
    j = oTitle.Length
    ReDim arr(j)
    j = 0
    '--https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/dictionary-object
    Set dic = CreateObject("Scripting.Dictionary")
    For Each item In oTitle
        arr(j) = item.Name
        '-----------https://excelmacromastery.com/excel-vba-collections/
        '------https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/collection-object
        Set dic(arr(j)) = New Collection
        j = j + 1
    Next
    j = -1
    For Each item In oA
        If i > 0 Then
            If Len(item.Name) > 0 Then
                j = j + 1
            Else
                dic(arr(j)).Add item.innertext
            End If
        Else
            i = 1
        End If
    Next
    Set Get_doubanBook_Tag = dic
    Set oA = Nothing
    Set oTag = Nothing
    Set oHtmlDom = Nothing
    Set dic = Nothing
End Function

Function Get_doubanBook_Tag_Rank(ByVal tName As String, Optional ByVal rType As Byte = 1, Optional ByVal Pages As Byte = 5) '获取豆瓣标签书籍排名
    Dim tUrl As String
    Dim oitem As Object
    Dim oA As Object
    Dim i As Integer, p As Integer, k As Byte
    Dim arr() As String
    Dim item As Object
    Dim ospan As Object
    Dim sResult As String
    Dim sType As String
    Dim opage As Object
    '-------------------每页20条,抓取前5页(有可能少于20)
    If rType = 1 Then
        sType = "S" '按照评分
    ElseIf rType = 2 Then
        sType = "R" '按照出版日期
    Else
        sType = "" '综合排序
    End If
    i = 0
    tUrl = "https://book.douban.com/tag/" & ThisWorkbook.Application.EncodeUrl(tName) & "?start=" & CStr(i * 20) & "&type=" & sType
    sResult = HTTP_GetData("GET", tUrl, "https://www.douban.com")
    WriteHtml sResult
    Set opage = oHtmlDom.getElementsByclassName("paginator")
    Set oA = opage(0).getElementsByTagName("a")
    i = oA.Length - 2
    If i = 0 Then k = 0
    sResult = oA(i).innertext
    k = CInt(sResult) '--------获取页码的数量
    If k < Pages Then Pages = k
    ReDim arr(Pages * 20 - 1, 6)
    Pages = Pages - 1
    p = 0: k = 0
    Set opage = Nothing
    For i = 0 To Pages
        If p > 0 Then
            isReady = True
            tUrl = "https://book.douban.com/tag/" & ThisWorkbook.Application.EncodeUrl(tName) & "?start=" & CStr(i * 20) & "&type=" & sType
            sResult = HTTP_GetData("GET", tUrl, "https://www.douban.com")
            If isReady = False Then MsgBox "Err": GoTo ErrHandle
            WriteHtml sResult
        Else
            p = 1
        End If
        Set oitem = oHtmlDom.getElementsByclassName("subject-list")
        For Each item In oitem.item(0).Children
            Set oA = item.getElementsByTagName("img")
            arr(k, 5) = oA(0).href
            Set ospan = item.getElementsByclassName("star clearfix")
            If ospan.item(0).Children.Length > 2 Then
                arr(k, 2) = ospan.item(0).Children.item(1).innertext '评分
                arr(k, 3) = ospan.item(0).Children.item(2).innertext '评分人数
            Else
                arr(k, 2) = ospan.item(0).Children.item(1).innertext '评分
                arr(k, 3) = "-"
            End If
            Set oA = item.getElementsByTagName("a")
            arr(k, 4) = oA(1).href '名称
            arr(k, 0) = oA(1).innertext '作品链接
            Set oA = item.ChildNodes.item(3)
            arr(k, 1) = oA.Children.item(1).innertext '详情
            Set ospan = item.getElementsByclassName("info")
            If ospan.item(0).Children.Length > 3 Then
                arr(k, 6) = ospan.item(0).Children.item(3).innertext
            Else
                arr(k, 6) = "-"
            End If
            k = k + 1
        Next
        Set oA = Nothing
        Set ospan = Nothing
        Set oHtmlDom = Nothing
        Set oitem = Nothing
    Next
    Get_doubanBook_Tag_Rank = arr
    Erase arr
    Exit Function
ErrHandle:
    Set oA = Nothing
    Set ospan = Nothing
    Set oHtmlDom = Nothing
    Set oitem = Nothing
    Set opage = Nothing
End Function

Function Get_doubanBook_Top250() As String() '抓取豆瓣读书Top250
    Const tUrl As String = "https://book.douban.com/top250?start=" '从0开始到225结束
    Dim oitem As Object
    Dim oA As Object
    Dim i As Integer, p As Integer, k As Byte
    Dim arr() As String
    Dim item As Object
    Dim ospan As Object
    Dim sResult As String
    Dim sUrl As String
    'i = oiTem.Length - 2
    'i = 0: p = 0
    ReDim arr(249, 6): ReDim Get_doubanBook_Top250(249, 6)
    For k = 0 To 9
        isReady = True
        sUrl = tUrl & CStr(k * 25)
        sResult = HTTP_GetData("GET", sUrl, "https://www.douban.com")
        If isReady = False Then MsgBox "Err": GoTo ErrHandle
        WriteHtml sResult
        Set oitem = oHtmlDom.getElementsByTagName("tr")
        p = 0
        For Each item In oitem
            If p > 0 Then '------------跳过第一个节点
                Set oA = item.getElementsByTagName("img")
                Set ospan = item.getElementsByclassName("star clearfix")
                arr(i, 5) = oA(0).href '封面链接
                arr(i, 2) = ospan.item(0).Children.item(1).innertext '评分
                arr(i, 3) = ospan.item(0).Children.item(2).innertext '评分人数
                Set oA = item.getElementsByTagName("a")
                arr(i, 4) = oA(1).href '名称
                arr(i, 0) = oA(1).innertext '作品链接
                Set oA = item.ChildNodes.item(3)
                arr(i, 1) = oA.Children.item(1).innertext '详情
                If oA.Children.Length > 3 Then
                    arr(i, 6) = oA.Children.item(3).innertext '评语
                Else
                    arr(i, 6) = "-"
                End If
                i = i + 1
            Else
                p = 1
            End If
        Next
        Set oA = Nothing
        Set ospan = Nothing
        Set oHtmlDom = Nothing
        Set oitem = Nothing
    Next
    Get_doubanBook_Top250 = arr
    Erase arr
    Exit Function
ErrHandle:
    Set oA = Nothing
    Set ospan = Nothing
    Set oHtmlDom = Nothing
    Set oitem = Nothing
End Function

Function Get_douban_SearchResult(ByVal strText As String, Optional ByVal iType As Byte = 0) As String() '获取豆瓣的搜索结果'通用搜索,电影搜索,书籍搜索
    Dim sResult As String
    Dim sUrl As String
    Dim oResult As Object
    Dim oA As Object
    Dim oTitle As Object
    Dim ospan As Object
    Dim i As Integer
    Dim strTemp As String
    Dim arr() As String
    Dim item As Object, itemx As Object
    Dim k As Integer, j As Integer
    Dim oInfo As Object
    Const tUrl As String = "https://www.douban.com/search?q=" '通用搜素
    Const mUrl As String = "https://www.douban.com/search?cat=1002&q=" '电影搜索
    Const bUrl As String = "https://www.douban.com/search?cat=1001&q=" '书籍搜索
    
    If iType = 1 Then '选择搜索的类型
        sUrl = bUrl
    ElseIf iType = 2 Then
        sUrl = mUrl
    Else
        sUrl = tUrl
    End If
    sUrl = sUrl & ThisWorkbook.Application.EncodeUrl(strText)
    sResult = HTTP_GetData("GET", sUrl, "https://www.douban.com")
    WriteHtml sResult
    Set oResult = oHtmlDom.getElementsByclassName("result-list") '获取搜索结果
    If oResult Is Nothing Then Set oHtmlDom = Nothing: Exit Function
    If oResult.Length = 0 Then Set oHtmlDom = Nothing: Exit Function
    i = oResult.item(0).Children.Length - 1
    If i < 1 Then Set oHtmlDom = Nothing: Exit Function
    ReDim arr(i, 6)
    ReDim Get_douban_SearchResult(i, 6)
    i = 0
    For Each item In oResult.item(0).Children
        If item.Classname = "result" Then '话题
            Set oTitle = item.getElementsByclassName("title")
            Set ospan = oTitle(0).getElementsByTagName("span")
            strTemp = ospan(0).innertext
            If strTemp = "[书籍]" Or strTemp = "[电影]" Or strTemp = "[电视剧]" Then
            '-------------------------------------类型,名称,链接,评分人数,评分,详情,链接图片
                Set oA = oTitle(0).getElementsByTagName("a") '链接
                arr(i, 0) = strTemp '类型
                arr(i, 1) = oA(0).innertext '名称
                arr(i, 2) = oA(0).href '链接
                Set oInfo = item.getElementsByclassName("rating-info")
                j = 3
                For Each itemx In oInfo(0).Children
                    strTemp = itemx.innertext
                    If Len(strTemp) > 0 Then
                        If itemx.Classname = "subject-cast" Then
                            If j <> 5 Then j = 5: arr(i, 4) = "-"
                        End If
                        arr(i, j) = strTemp
                        j = j + 1
                    End If
                Next
                Set oTitle = item.getElementsByclassName("pic")
                arr(i, 6) = oTitle(0).getElementsByTagName("img")(0).href '图片
                i = i + 1
            End If
            strTemp = ""
        End If
    Next
    Get_douban_SearchResult = arr
    Erase arr
    Set oInfo = Nothing
    Set oTitle = Nothing
    Set ospan = Nothing
    Set oResult = Nothing
    Set oHtmlDom = Nothing
End Function

Function Get_Douban_FilmRank_Type() As String() '获取豆瓣的电影分类的id
    Dim sResult As String
    Dim oType As Object, oA As Object
    Dim arr() As String
    Dim i As Integer, j As Integer
    Dim strTemp As String
    Const ref As String = "https://movie.douban.com/tag/"
    Const tUrl As String = "https://movie.douban.com/chart"
    
    isReady = True
    sResult = HTTP_GetData("GET", tUrl, ref)
    If isReady = False Then Exit Function
    WriteHtml sResult
    Set oType = oHtmlDom.getElementsByclassName("types")
    i = oType.item(0).Children.Length - 1
    j = 0
    ReDim arr(i, 1)
    ReDim Get_Douban_FilmRank_Type(i, 1)
    For Each item In oType.item(0).Children
        Set oA = item.getElementsByTagName("a")
        i = oA.Length - 1
        arr(j, 0) = oA(i).Text
        strTemp = oA(i).href
        strTemp = Split(strTemp, "&interval")(0)
        arr(j, 1) = Right$(strTemp, Len(strTemp) - InStrRev(strTemp, "="))
        j = j + 1
    Next
    Get_Douban_FilmRank_Type = arr
    Set oHtmlDom = Nothing
    Set oType = Nothing
    Set oA = Nothing
End Function

Function Get_Douban_FilmRank_Spe(ByVal tName As String, ByVal sType As String, Optional iStart As Byte = 0, Optional ByVal iLimit As Byte = 20) As String()
    '豆瓣的电影排名从0开始, 每页20条
    'https://movie.douban.com/j/chart/top_list?type=20&interval_id=100%3A90&action=&start=0&limit=20
    Dim sResult As String
    Dim xUrl As String
    Dim ref As String
    Dim arr() As String
    Dim oDic As Object
    Dim item As Object
    Dim ciTems As Variant, aiTem
    Dim strTemp As String
    Dim wb As Workbook
    Dim i As Integer, k As Integer
    Dim m As Integer, n As Integer
    Dim FilePath As String
    Const tUrl As String = "https://movie.douban.com/j/chart/top_list?type="
    
    If iLimit > 100 Then iLimit = 100
    ref = tUrl & sType & "&interval_id=100%3A90&action=&start="
    xUrl = ref & iStart & "&limit=" & iLimit
    sResult = HTTP_GetData("GET", xUrl, ref)
    '-返回12组数据,演员列表,封面url,排名,发行时间,出版国家,标题,评分,类型,作品url,投票人数
    Set oDic = JsonConverter.ParseJson(sResult)
    Set wb = Workbooks.Add
    m = 1
    For Each item In oDic '第一层0开始,二层1开始
        ciTems = item.Items
        k = item.Count - 1
        For i = 1 To k
            If i <> 3 And i <> 14 Then
                If IsObject(ciTems(i)) = True Then
                    For Each aiTem In ciTems(i)
                        strTemp = strTemp & aiTem & ";"
                    Next
                    strTemp = Left$(strTemp, Len(strTemp) - 1)
                Else
                    strTemp = ciTems(i)
                End If
                n = Choose_Column(i)
                If n > 0 Then Creat_douban_Sheet wb, m, n, strTemp
                strTemp = ""
            End If
        Next
        m = m + 1
    Next
    m = m + 5
    ref = "https://movie.douban.com/typerank?type_name=" & ThisWorkbook.Application.EncodeUrl(tName) & "&type=" & sType & "&interval_id=100:90&action="
    With wb.Sheets(1)
        .Name = "douban"
        .Cells(1, 1) = "豆瓣电影"
        .Cells(2, 1) = "类型"
        .Hyperlinks.Add Anchor:=.Cells(2, 3), address:=ref, TextToDisplay:="链接"
        .Cells(2, 2) = sType
        .Cells(3, 1) = "评分排名"
        .Cells(4, 1) = "排名范围:"
        .Cells(4, 2).NumberFormatLocal = "@"
        .Cells(4, 2) = CStr(iStart) + 1 & "-" & CStr(iLimit)
        .Cells(4, 4) = "数据生成时间:"
        .Cells(4, 5) = Now
        .Cells(5, 1).Resize(1, 12) = Array("豆瓣ID", "名称", "排名", "评分", "评分人数", "出品地区", "发行时间", "标签", "演员数量", "演员列表", "作品链接", "作品封面")
        .Range("a5:l5").Font.Bold = True
        .Range("a5:l5").Font.Size = 12
        .Range("a1:a4").Font.Size = 12
        .Range("a1:a4").Font.Bold = True
        .Columns(2).AutoFit
        .Columns(8).AutoFit
        .Columns(7).AutoFit
        .Columns(6).AutoFit
        .Columns(4).AutoFit
        .Columns(5).AutoFit
        .Range("d6:d" & m).NumberFormatLocal = "0.0_ "
        .Range("j6:j" & m).WrapText = True
    End With
    FilePath = ThisWorkbook.Path & "\" & Format(Now, "yyyymmddhhmmss") & ".xlsx"
    wb.SaveAs FilePath
    Set wb = Nothing
    Set oDic = Nothing
End Function

Private Function Choose_Column(ByVal intx As Byte) As Byte
    Dim i As Byte
    Select Case intx
        Case 1: i = 3
        Case 2: i = 12
        Case 4: i = 1
        Case 5: i = 8
        Case 6: i = 6
        Case 7: i = 2
        Case 8: i = 11
        Case 9: i = 7
        Case 10: i = 9
        Case 11: i = 5
        Case 12: i = 4
        Case 13: i = 10
        Case Else: i = 0
    End Select
    Choose_Column = i
End Function

Private Function Creat_douban_Sheet(ByVal wbx As Workbook, ByVal r As Integer, ByVal c As Byte, strText As String)
    r = r + 5
    With wbx.Sheets(1)
        If c = 11 Or c = 12 Then
            .Hyperlinks.Add Anchor:=.Cells(r, c), address:=strText, TextToDisplay:="链接"
        Else
        .Cells(r, c) = strText
        End If
    End With
End Function
'-------------------------------------------------------------------------------------------------------------------豆瓣电影数据

'-------------------------------------------------------------------------------------同花顺
'---------------------------同花顺资金流向-http://data.10jqka.com.cn/funds/ggzjl/
'使用cookie绕过反爬,如果使用单一cookie爬取数量达到250就会出现403 forbidden
Function Get_THS_MoneyFlow() As String '同步
Dim sResult As String
Dim arr() As String
Dim item As Object, itemx As Object
Dim oTable As Object
Dim iCookie As String
Dim i As Integer, k As Integer, p As Integer
Const tUrl As String = "http://data.10jqka.com.cn/funds/ggzjl/field/zdf/order/desc/page/"

ReDim arr(499, 10)
ReDim Get_THS_MoneyFlow(499, 10)
For p = 1 To 10
    isReady = True
    iCookie = Cookie_Lists(p)
    sResult = HTTP_GetData("GET", tUrl & CStr(p) & "/ajax/1/free/1/", "http://data.10jqka.com.cn/funds/ggzjl/", sCharset:="gb2312", sCookie:=iCookie)
    If isReady = False Then GoTo 200
    WriteHtml sResult
    Set oTable = oHtmlDom.getElementsByclassName("m-table J-ajax-table")
    If oTable Is Nothing Then Set oHtmlDom = Nothing: Set oTable = Nothing: GoTo 100
    If oTable.Length = 0 Then Set oHtmlDom = Nothing: Set oTable = Nothing: GoTo 100
    For Each item In oTable.item(0).Children.item(1).Children
        k = 0
        For Each itemx In item.Children
            arr(i, k) = itemx.innertext
            k = k + 1
        Next
        i = i + 1
    Next
200
    Set oTable = Nothing
    Set oHtmlDom = Nothing
    DoEvents
100
    Sleep 200
Next
Get_THS_MoneyFlow = arr
Erase arr
End Function

Function Get_THS_MoneyFlow_Async() '异步
    Dim arr() As New cWinHttpRQ
    Dim arrt
    Dim iCookie As String
    ReDim arr(25) As New cWinHttpRQ
    For i = 0 To 25
        If i Mod 2 = 0 Then iCookie = Cookie_Lists(i \ 2)
        With arr(i)
            .Index = i
            .url = "http://data.10jqka.com.cn/funds/ggzjl/field/zdf/order/desc/page/" & CStr(i + 1) & "/ajax/1/free/1/"
            .StartRe iCookie
            Do Until .isOK = True
            If .IsErr = True Then Exit Do
            DoEvents
            Loop
        End With
    Next
'    arrt = arr(0).Result
End Function

Private Function Cookie_Lists(ByVal i As Byte) As String 'cookie列表
    Dim iCookie As String
    Select Case i
        Case 0: iCookie = "v=AikAzFC_U9UeKm9h0iqO6GGLON6H9h0oh-pBvMsepZBPkkNIE0Yt-Bc6UYpY"
        Case 1: iCookie = "v=ArLCMrQYOEDB8AR8cDZrATO9A_ORQ7bd6EeqAXyL3mVQD1xnZNMG7bjX-h1P"
        Case 2: iCookie = "v=AjMJ4yaYaVerLCWDCJ7kqqfdwjxZaMcqgfwLXuXQj9KJ5FlqbThXepHMm6L2"
        Case 3: iCookie = "v=An3ePsQD_xHyuVv9ceBM0MDkjNJyGrFsu04VQD_CuVQDdpPGB2rBPEueJRvM"
        Case 4: iCookie = "v=AjGS2hgHy01mBWdZXWv4fDQoQLbOHqWQT5JJpBNGLfgXOl_iW261YN_iWXig"
        Case 5: iCookie = "v=ArG7k_Q8S83mhefZ3ZZ4_LSowDZOniUQzxLJJJPGrXiXut9i2-414F9i2f4g"
        Case 6: iCookie = "v=AoSOjGn_hvrD9jI2kP6la3F7VQlznagHasE8S54lEM8SySr9xq14l7rRDNft"
        Case 7: iCookie = "v=AneN8nTw5cPPFWFPZgvmIi5qBmDEPEueJRDPEskkk8ateJnU0Qzb7jXgX27a"
        Case 8: iCookie = "v=AhzRQfIhrqIralp-eZND30T47THKlcC_QjnUg_YdKIfqQbYnHqWQT5JJpB9F"
        Case 9: iCookie = "v=AgNvGHb3WecYfxUT-sbUOhftksypeJe60Qzb7jXgX2LZ9Cn6vUgnCuHcazJG"
        Case 10: iCookie = "v=Aq5aKC-ufBRtoIjw1xjRiULi_w90r3KphHMmjdh3GrFsu0T5QD_CuVQDdp6r"
        Case 11: iCookie = "v=AhgXgVJO8nZn0N7y62uPa2gU6U2uAXyL3mVQD1IJZNMG7bI7-hFMGy51IJSh"
        Case 12: iCookie = "v=AjGi5cMQy01WmGdZabgGAFkjQLbPHqX8T5BJtRNGL68UCltgW261YN_iWXeg"
    End Select
    Cookie_Lists = iCookie
End Function
'----------------------------------------------------------------------------------------------同花顺
Sub Get_eastmoney_FundLists() '获取东方财富基金列表
    Dim i As Byte
    Dim k As Long, m As Long, p As Long, n As Long
    Dim sResult As String
    Dim oRegx As New cRegex
    Dim wb As Workbook
    Dim arr() As String
    Dim artTemp() As String
    Dim dic As New Dictionary
    Dim strTemp As String
    Const tUrl As String = "http://fund.eastmoney.com/js/fundcode_search.js"
    
    isReady = True
    sResult = HTTP_GetData("GET", tUrl)
    If isReady = False Then MsgBox "获取数据失败", vbCritical, "Warning": Exit Sub
    arr = oRegx.xMatch(sResult, Chr(34) & "(.*?)" & Chr(34))
    m = UBound(arr)
    n = (m + 1) / 5
    m = n - 1
    ReDim arrTemp(m, 4)
    dic.CompareMode = BinaryCompare
    For k = 0 To m
        For i = 0 To 4
            strTemp = Replace(arr(p), Chr(34), "")
            arrTemp(k, i) = strTemp
            If i = 3 Then
                If dic.Exists(strTemp) = False Then dic.Add strTemp, 1 Else dic(strTemp) = dic(strTemp) + 1 '统计各个基金类型的数量
            End If
            p = p + 1
        Next
    Next
    Set wb = Workbooks.Add
    With wb.Sheets(1)
        .Name = "Fund_Lists"
        .Cells(1, 2) = "数据来源:"
        .Cells(1, 3) = "东方财富"
        .Cells(2, 2) = "数据抓取时间:"
        .Cells(2, 3) = Now
        .Cells(2, 4) = "基金数量:" & CStr(n)
        .Cells(2, 5) = "基金分类:"
        p = dic.Count - 1
        m = 2
        strTemp = ""
        For i = 0 To p '------字典也是从0开始
            strTemp = strTemp & dic.Keys(i) & ", " '注意这里的dic.Keys的写法, 实际写法为dic.Keys()(i), 如果添加了引用就可以忽略()
            .Cells(m, 7) = dic.Keys(i) & ":"
            .Cells(m, 8) = dic.Items(i)
            m = m + 1
        Next
        .Cells(2, 6) = Left$(strTemp, Len(strTemp) - 2)
        .Cells(3, 2).Resize(1, 5) = Array("基金代码", "基金简拼", "基金名称", "基金类别", "基金全拼")
        .Range("b3:f3").Font.Bold = True
        .Range("b4:b" & n + 3).NumberFormatLocal = "@" '先将这个区域(放置000001类型数据的)的格式调整为文本型, 不然数据会被Excel吞掉
        .Cells(4, 2).Resize(n, 5) = arrTemp
        .Columns.AutoFit
    End With
    wb.SaveAs ThisWorkbook.Path & "\FundLists" & Format(Now, "yyyymmddhhmmss") & ".xlsx"
    Erase arr
    Erase arrTemp
    Set dic = Nothing
    Set wb = Nothing
    Set oRegx = Nothing
End Sub
'http://fund.eastmoney.com/f10/F10DataApi.aspx?type=lsjz&code=001594&sdate=2019-02-01&edate=2020-04-12&per=50
'这个数据接口最大一次只能显示49条数据, 如果指定的显示条数和可显示的条数不相匹配,将返回默认的数据, 处理起来较为麻烦

Function Fund_History_Hexun(ByVal sID As String, ByVal StartDate As String, EndDate As String) As String() '和讯的数据接口
    Dim oTable As Object, oTr As Object
    Dim sResult As String
    Dim item As Object
    Const rUrl As String = "http://jingzhi.funds.hexun.com"
    Const sUrl As String = "http://jingzhi.funds.hexun.com/DataBase/jzzs.aspx?fundcode="
    Dim tUrl As String
    Dim i As Integer, k As Integer, p As InterfaceInfo
    Dim arr() As String
    
    tUrl = sUrl & sID & "&startdate=" & StartDate & "&enddate=" & EndDate '链接 "& startdate="出现空格, 将返回该基金的全部历史数据
    sResult = HTTP_GetData("GET", tUrl, rUrl, sCharset:="gb2312")
    WriteHtml sResult
    Set oTable = oHtmlDom.getElementsByclassName("n_table m_table")  'n_table m_table
    Set oTr = oTable.item(0).getElementsByTagName("tr")
    i = oTr.Length - 2 '去掉第一行数据(表头)
    For Each item In oTr
        If p = 0 Then p = 1: k = item.Children.Length - 1: ReDim arr(i, k): ReDim Fund_History_Hexun(i, k): i = 0: k = 0
        For Each itemx In item.Children
            arr(i, k) = itemx.innertext
            k = k + 1
        Next
        i = i + 1: k = 0
    Next
    Fund_History_Hexun = arr
    Erase arr
    Set oTr = Nothing
    Set oTable = Nothing
    Set oHtmlDom = Nothing
End Function

Function Fund_Essentials(ByVal fID As String, ByVal dMode As String) '基金概要
    '7天,1个月, 3个月, 半年, 1年
    Const tUrl As String = "https://fund.xueqiu.com/dj/open/fund/growth/"
    Dim sResult As String
    Dim sUrl As String
    Dim arr() As String
    Dim arrTemp() As String
    Dim oRegx As New cRegex
    Dim i As Integer, k As Integer, j As Byte, m As Integer, n As Integer, p As Integer
    Dim wb As Workbook
    Dim idown As Integer, iup As Integer, ikeep As Integer '涨跌
    Dim iMax As Double, iMin As Double '最大, 最小值
    Dim strTemp As String
    Dim x As Double
    Dim rUrl As String
    
    rUrl = "https://xueqiu.com/S/F" & fID
    sUrl = tUrl & fID & "?day=" & dMode
    isReady = True
    sResult = HTTP_GetData("GET", sUrl, rUrl)
    If isReady = False Then Exit Function
    arr = oRegx.xMatch(sResult, Chr(34) & "(\-?\d.*?)" & Chr(34))
    k = UBound(arr)
    m = (k + 1) / 5
    k = m - 1
    ReDim arrTemp(k, 3)
    For i = 0 To k
        For j = 0 To 4
            If j < 3 Then
                strTemp = Replace(arr(p), Chr(34), "")
                arrTemp(i, j) = strTemp
                If j = 2 Then
                    x = Val(strTemp)
                    If x > 0 Then
                        iup = iup + 1
                        strTemp = "UP"
                    ElseIf x < 0 Then
                        idown = idown + 1
                        strTemp = "Down"
                    Else
                        ikeep = ikeep + 1
                        strTemp = "Keep"
                    End If
                    arrTemp(i, 3) = strTemp
                End If
            End If
            p = p + 1
        Next
    Next
    Set wb = GetObject("C:\Users\adobe\Desktop\FundLists20200415180408.xlsx")
    With wb.Sheets("Essentials")
        .Cells(5, 2).Resize(m, 4) = arrTemp
        .Range("b5:b" & m + 4).NumberFormatLocal = "@"
        .Range("d5:d" & m + 4).NumberFormatLocal = "0.0000%"
        .Columns.AutoFit
    End With
    wb.Save
    Set wb = Nothing
    Set oRegx = Nothing
    Erase arr
    Erase arrTemp
End Function

Function Fund_Profile(ByVal id As String) '同花顺
    Const tUrl As String = "http://fund.10jqka.com.cn/"
    Const pUrl As String = "http://fund.10jqka.com.cn/data/client/myfund/" '获取基金概要 api
    Dim sUrl As String
    Dim sResult As String
    Dim oList As Object, oTitle As Object
    Dim item As Object, itemx As Object, itema As Object
    Dim arr() As String
    Dim arrTemp() As String, sarrTemp
    Dim oRegx As New cRegex
    Dim i As Integer, k As Integer, iMode As Byte, j As Byte, p As Byte
    Dim date1 As Date, date2 As Date, d1 As Byte, d2 As Byte
    Dim arrx, strTemp As String
    
    arrx = Array(0, 1, 3, 4, 5, 6, 13, 14, 15, 16, 17, 20, 21, 35, 36, 37, 38, 39, 40, 58) '需要的数据
    sUrl = pUrl & id
    sResult = HTTP_GetData("GET", sUrl, tUrl)
    arr = oRegx.xSubmatch(sResult, Chr(34) & "(.*?)" & Chr(34) & ":\s?" & Chr(34) & "(.*?)" & Chr(34))
    Set oRegx = Nothing
    k = UBound(arrx)
    ReDim sarrTemp(k)
    For i = 0 To k
        strTemp = arr(arrx(i), 1)
        If InStr(strTemp, "\u") > 0 Then strTemp = Unicode2Character(strTemp) '数据中的汉字为Unicode字符
        sarrTemp(i) = strTemp
    Next
    '----------------------------------------------------------------------------'基金概要
    sUrl = tUrl & id & "/portfolioindex.html" '获取持仓情况 html
    sResult = HTTP_GetData("GET", sUrl, tUrl)
    WriteHtml sResult
    Set oTitle = oHtmlDom.getElementsByclassName("o-title") '先获得更新的日期数据
    '--------如果日期不相等,那么就获取最新的数据, 获取重仓股,债券
    For Each item In oTitle
        strTemp = item.innertext
        If InStr(strTemp, "重仓股") > 0 Then
            If InStr(strTemp, "数据更新") > 0 Then date1 = CDate(Trim(Split(strTemp, " ")(1))): d1 = 1
        ElseIf InStr(strTemp, "重仓债") > 0 Then
            If InStr(strTemp, "数据更新") > 0 Then date2 = CDate(Trim(Split(strTemp, " ")(1))): d2 = 1
        End If
    Next
    If d1 > 0 And d2 > 0 Then '最起码1组值
        If date1 > date2 Then
            iMode = 1
        ElseIf date1 = date2 Then '获取双组值
            iMode = 3
        Else
            iMode = 2
        End If
    ElseIf d1 = 0 And d2 > 0 Then '获取单组值
        iMode = 2
    ElseIf d1 > 0 And d2 = 0 Then '获取单组值
        iMode = 1
    Else
        Exit Function
    End If
    Set oTitle = Nothing
    '--------------------------获取数据更新时间的情况
    Set oList = oHtmlDom.getElementsByclassName("s-list") '获得3组元素,重仓股,重仓债,调仓
    i = 0: k = 0: j = 0
    Select Case iMode
        Case 1:
        For Each item In oList.item(0).Children
            If p > 0 Then
                For Each itemx In item.Children
                    arrTemp(i, k) = itemx.innertext
                    k = k + 1
                Next
                i = i + 1
                k = 0
            Else
                p = 1
                k = oList.item(0).Children.Length - 2
                ReDim arrTemp(k, 5)
                k = 0
            End If
        Next
        Case 2:
        For Each item In oList.item(1).Children
            If p > 0 Then
                For Each itemx In item.Children
                    arrTemp(i, k) = itemx.innertext
                    k = k + 1
                Next
                i = i + 1
                k = 0
            Else
                p = 1
                k = oList.item(0).Children.Length - 2
                ReDim arrTemp(k, 5)
                k = 0
            End If
        Next
        Case 3:
        k = oList.item(0).Children.Length + oList.item(1).Children.Length - 3
        ReDim arrTemp(k, 5)
        k = 0
        For Each item In oList
            If p = 2 Then Exit For
            For Each itemx In item.Children
                If j > 0 Then
                    For Each itema In itemx.Children
                        arrTemp(i, k) = itema.innertext
                        k = k + 1
                    Next
                    k = 0
                    i = i + 1
                End If
                j = 1
            Next
            j = 0
            p = p + 1
        Next
    End Select
    Set oList = Nothing
    Set oHtmlDom = Nothing
End Function
'----------------------------------------------------------------------虾米音乐
Sub Write_FavoriteLists(ByVal id As String, ByVal Pages As Integer)
'----------------------------虾米的个人收藏列表部分缺失
Dim arr() As String
Dim wb As Workbook
Dim i As Integer
Dim k As Integer
Set wb = Workbooks.Add
With wb.Sheets(1)
    .Name = "xiami"
    For i = 1 To 23
        arr = Get_FavoriteLists_fromXiami(id, CStr(i))
        k = UBound(arr, 1) + 1
        If k > 1 Then
            .Cells(k * (i - 1) + 1, 1).Resize(k, 60) = arr
            Erase arr
        End If
    Next
End With
wb.SaveAs ThisWorkbook.Path & "xiami_FavoriteLists" & Format(Now, "yyyymmddhhmmss") & ".xlsx"
Set wb = Nothing
End Sub

Private Function Get_FavoriteLists_fromXiami(ByVal id As String, Optional ByVal page As String = "1") As String() '获取虾米音乐个人喜欢歌曲列表
    Const tUrl As String = "https://www.xiami.com/api/favorite/getFavorites?_q="
    Dim sKey As String
    Dim sign As String
    Dim sResult As String
    Dim sUrl As String
    Dim arr() As String
    Dim rUrl As String
    
    rUrl = "https://www.xiami.com/list?scene=favorite&type=song&query={%22userId%22:%2212749213%22}"
    xmCookie = "xmgid=2154fdc1-3fc1-49fa-b2cf-efb24dfd51ed; xm_sg_tk=7d6debda649bc154b174178e718bef0c_1586848087715; xm_sg_tk.sig=0myrsohHEcfi0avK_BL9idJ13e2typy3rltlB3Tribg; _uab_collina=158684808593371547326548; cna=Wk8cF5l85k0CARsmFQrvGOtK; _xm_ncoToken_login=web_login_1586848090507_0.42344195358455794; xm_token=7ef8d1895586a53dbbb27d306cd4409btc25c; uidXM=12749213; member_auth=gTqdToZO420DwNLUKdFvMkVonqyCSzjnld8F7M9p4EIhYs96K%2BKdz9PUG3ANjinO4ipiGYSd3zxKFuxcKODenOHP; "
    tkCookie = "7d6debda649bc154b174178e718bef0c_1586848087715"
    sKey = "{" & Chr(34) & "userId" & Chr(34) & ":" & Chr(34) & id & Chr(34) & "," & Chr(34) & "type" & Chr(34) & ":1," & Chr(34) & "pagingVO" & Chr(34) & ":{" & Chr(34) & "page" & Chr(34) & ":" & page & "," & Chr(34) & "pageSize" & Chr(34) & ":30}}"
    sign = Xiami_Sign_Generator(tkCookie, tUrl, sKey)
    '逗号,双引号,:保留
    sKey = UTF8_URLEncoding(sKey)
    sKey = Replace(sKey, "{", "%7B")
    sKey = Replace(sKey, "}", "%7D")
    sKey = Replace(sKey, Chr(34), "%22")
    sUrl = tUrl & sKey & "&_s=" & sign
    sResult = HTTP_GetData("GET", sUrl, rUrl, sCookie:=xmCookie) '等到json格式的数据
    '这部分的json数据结构和搜索部分的数据结构一样, 可以共享
    '-----------截取开头的部分即可
    If InStr(1, Left$(sResult, 30), "success", vbTextCompare) = 0 Then Exit Function '判断有无有效数据返回
    arr = Json_Data_Treat(sResult)
    Get_FavoriteLists_fromXiami = arr
    Erase arr
End Function

Private Function Xiami_Song_Download_Link(ByVal sID As String) As String '获取歌曲的下载链接
    Dim sKey As String
    Dim sign As String
    Dim sResult As String
    Dim sUrl As String
    Dim rUrl As String
    Dim strTemp As String
    Const tUrl As String = "https://www.xiami.com/api/song/getPlayInfo?_q="
    '虾米大量下架无版权的作品, 将会出现404错误
    rUrl = "https://www.xiami.com/"
    sKey = "{" & Chr(34) & "songIds" & Chr(34) & ":[" & sID & "]}"
    If Xiami_Pre_Check(sID) = True Then Exit Function
    sign = Xiami_Sign_Generator(tkCookie, tUrl, sKey)
    '逗号,双引号,:保留
    sKey = UTF8_URLEncoding(sKey)
    sKey = Replace(sKey, "{", "%7B")
    sKey = Replace(sKey, "}", "%7D")
    sKey = Replace(sKey, Chr(34), "%22")
    sUrl = tUrl & sKey & "&_s=" & sign
    sResult = HTTP_GetData("GET", sUrl, rUrl, sCookie:=xmCookie)  '等到json格式的数据
    If InStr(1, sResult, "success", vbTextCompare) > 0 Then
        strTemp = Trim(Split(sResult, Chr(34) & "listenFile" & Chr(34))(1))
        strTemp = Split(strTemp, ",")(0)
        strTemp = Mid$(strTemp, 3, Len(strTemp) - 3) '获取音质最好的链接
    End If
    Xiami_Song_Download_Link = strTemp
End Function
'------------------------通过JsonConverterjson数据
Private Function Get_Download_Links(ByVal strTex As String) As String()
    Dim dic As Object
    Dim sDic As Object
    Dim dDic As Object
    Dim pDic As Object
    Dim item
    Dim itemx
    Dim itema
    Dim i As Byte, k As Byte
    
    Set dic = JsonConverter.ParseJson(strText)
    Set sDic = dic("result")
    Set dDic = sDic("data")
    For Each item In dDic.Items
        For Each itemx In item
            Set pDic = itemx("playInfos")
            i = pDic.Count - 1
            ReDim arr(i)
            ReDim Get_Download_Links(i) '获取所有的下载链接
            For Each itema In pDic
                arr(k) = itema("listenFile")
                k = k + 1
            Next
        Next
    Next
    Get_Download_Links = arr
    Set dic = Nothing
    Set pDic = Nothing
    Set sDic = Nothing
    Set dDic = Nothing
End Function

Function Xiami_Search(ByVal Keyword As String, Optional ByVal sType As Byte = 0) As String() '歌名搜索, 歌手搜索,专辑搜索,歌单搜索
    Const sUrl As String = "https://www.xiami.com/api/search/searchSongs?_q="
    Const arUrl As String = "https://www.xiami.com/api/search/searchArtists?_q="
    Const alUrl As String = "https://www.xiami.com/api/search/searchAlbums?_q="
    Const cUrl As String = "https://www.xiami.com/api/search/searchCollects?_q"
    Dim sKey As String
    Dim sign As String
    Dim sResult As String
    Dim tUrl As String
    Dim arr() As String
    '注意使用的转码方式,https://www.cnblogs.com/qlqwjy/p/9934706.html
    'encodeURIComponent
    'encodeURI
    'application.encodeurl实际上是encodeurlcomponent函数
    Select Case sType
        Case 1: tUrl = arUrl
        Case 2: tUrl = cUrl
        Case 3: tUrl = alUrl
        Case Else: tUrl = sUrl
    End Select
    If Xiami_Pre_Check(Keyword) = True Then Exit Function
    sKey = "{" & Chr(34) & "key" & Chr(34) & ":" & Chr(34) & Keyword & Chr(34) & "," & Chr(34) & "pagingVO" & Chr(34) & ":{" & Chr(34) & "page" & Chr(34) & ":1," & Chr(34) & "pageSize" & Chr(34) & ":30}}"
    sign = Xiami_Sign_Generator(tkCookie, tUrl, sKey)
    '逗号,双引号,:保留
    sKey = UTF8_URLEncoding(sKey)
    sKey = Replace(sKey, "{", "%7B")
    sKey = Replace(sKey, "}", "%7D")
    sKey = Replace(sKey, Chr(34), "%22")
    tUrl = tUrl & sKey & "&_s=" & sign
    sResult = HTTP_GetData("GET", url, tUrl, sCookie:=xmCookie) '等到json格式的数据
    If InStr(1, Left$(sResult, 30), "success", vbTextCompare) = 0 Then Exit Function '判断有无有效数据返回
    arr = Json_Data_Treat(sResult)
    Xiami_Search = arr
'    -如果仅仅是获取单一的值, 如歌曲的id,可使用正则更为方便
'    Dim Regx As New cRegex
'    Dim arrTemp() As String
'    arr = Regx.xMatch(sResult, Chr(34) & "songId" & Chr(34) & ":+[\d]{6,}") '"songId":1770188126
'    Dim i As Integer
'    i = UBound(arr)
'    ReDim arrTemp(i)
'    For k = 0 To 1
'    arrTemp(i) = Trim(Split(arr(i), ":")(1))
'    Next
'    Xiami_Search = arrTemp
'    Set Regx = Nothing
'    -或者将匹配样式修改为
'    sPartten=chr(40) & chr(34) & "songID" & chr(34) &chr(41) & ":"& chr(40) &"+[\d]{6,}" & chr(41)
'    使用Submatch分离数据, 不需要split
End Function
'专辑ID
'专辑语言
'专辑封面
'专辑名称
'专辑字符串id
'歌手Alias (别名)
'作曲家
'歌手封面
'歌手名: 原名
'歌名
'歌曲id'--------------------------主要获取这几个主要项
Private Function Json_Data_Treat(ByVal strText As String) As String() '搜索获取的json数据的处理
    Dim dic As Object
    Dim sDic As Object
    Dim dDic As Object
    Dim pDic As Object
    Dim isDic As Object
    Dim item
    Dim itemx
    Dim i As Byte, k As Byte, p As Byte
    Dim arr() As String
    Dim strTemp As String
    
    Set dic = JsonConverter.ParseJson(strText)
    Set sDic = dic("result")
    Set dDic = sDic("data")
    Set isDic = dDic("songs")
    i = isDic.Count - 1
    For Each item In isDic
        For Each itemx In item.Items
            If p = 0 Then k = item.Count - 1: ReDim arr(i, k): ReDim Json_Data_Treat(i, k): i = 0: k = 0: p = 1
            If IsObject(itemx) = False Then
                If IsNull(itemx) = False Then
                    strTemp = itemx
                    If Len(strTemp) > 0 Then
                        arr(i, k) = strTemp
                    Else
                        arr(i, k) = "-"
                    End If
                Else
                    arr(i, k) = "-"
                End If
                k = k + 1
            End If
        Next
        i = i + 1
        k = 0
    Next
    Json_Data_Treat = arr
    Set dic = Nothing
    Set isDic = Nothing
    Set sDic = Nothing
    Set dDic = Nothing
    Erase arr
End Function

Private Function Xiami_Pre_Check(ByVal strText As String) As Boolean '检查输入的内容有误干扰项, 检查cookie是否获取成功
    Dim i As Byte
    If InStr(1, strText, Chr(34), vbBinaryCompare) > 0 Then
        i = 1
    ElseIf InStr(1, strText, "}", vbBinaryCompare) > 0 Then
        i = 1
    ElseIf InStr(1, strText, "{", vbBinaryCompare) > 0 Then
        i = 1
    End If
    If i = 1 Then MsgBox "存在受限字符", vbInformation, "Tips": Exit Function
    If Xiami_Cookie_Generator = False Then Xiami_Pre_Check = True: MsgBox "获取cookie失败", vbCritical + vbInformation, "Warning": Exit Function
End Function

Private Function Xiami_Cookie_Generator() As Boolean '虾米音乐cookie获取
    '注意cookie有使用寿命,如果cookie失效需要重新获取cookie
    Const rUrl As String = "https://www.xiami.com/"
    isReady = True
    If Len(tkCookie) = 0 Then
        xmCookie = HTTP_GetData("GET", "https://www.xiami.com/", ReturnRP:=2)
        If isReady = True Then tkCookie = Split(Split(xmCookie, "; ")(2), "=")(1)
    End If
    Xiami_Cookie_Generator = isReady
End Function

Private Function Xiami_Sign_Generator(ByVal tCookie As String, ByVal pUrl As String, ByVal qStr As String) As String  '虾米sign生成
    '需要搭配cookie和破解sign的方式才能获取到接口的信息
    'cookie的xm_sg_tk值的第一部分
    '84c38dbe9481c68a787a781f8534545f_1586662991803, 84c38dbe9481c68a787a781f8534545f这部分
    '常量:"_xmMain_"
    '变量: 请求url部分 https://www.xiami.com/api/favorite/getFavorites, 的"/api/favorite/getFavorites"
    '请求数据的_q值
    'sign=getmd5hash_string(xm_sg_tk(0) &"_xmMain_"& "/api/favorite/getFavorites" & _q)
    '-------------------------此签名方式适用于整个虾米页面
    '音乐下载链接不需要cookie
    Dim strText As String
    If InStr(1, tCookie, "_", vbBinaryCompare) Then tCookie = Split(tCookie, "_")(0)
    If InStr(1, pUrl, "https://www.xiami.com/", vbBinaryCompare) > 0 Then
        pUrl = Split(pUrl, "https://www.xiami.com")(1)
        If InStr(1, pUrl, "?", vbBinaryCompare) > 0 Then pUrl = Split(pUrl, "?")(0)
    End If
    strText = tCookie & "_xmMain_" & pUrl & "_" & qStr
    Xiami_Sign_Generator = LCase(GetMD5Hash_String(strText))
End Function

Function Bilibili_Favlist(ByVal sID As String, Optional ByVal IsCreat As Boolean) As String() 'bili站点的个人收藏夹
    Const crUrl As String = "https://api.bilibili.com/x/v3/fav/folder/created/list-all?up_mid="
    Const clUrl As String = "https://api.bilibili.com/x/v3/fav/folder/collected/list?pn=1&ps=20&up_mid="
    Dim sUrl As String
    Dim dic As Object
    Dim sResult As String
    Dim dDic As Object, lDic As Object
    Dim item, i As Integer
    Dim arr() As String
    Dim rUrl As String, tUrl As String
    
    '获取每个收藏夹的id, 名称, 数量 '"id":949337310,"fid":9493373,"mid":441644010,"attr":22,"title":"Finance","fav_state":0,"media_count":1
    tUrl = IIf(IsCreat = False, crUrl, clUrl)
    sUrl = tUrl & sID & "&jsonp=jsonp"
    rUrl = "https://space.bilibili.com/" & sID & "/favlist" '必须有这个refer, 否则将会出现403 forbidden
    sResult = HTTP_GetData("GET", sUrl, rUrl)
    If Len(sResult) = 0 Then Exit Function
    Set dic = JsonConverter.ParseJson(sResult)
    Set dDic = dic("data")
    Set lDic = dDic("list")
    i = lDic.Count - 1
    ReDim arr(i, 2)
    ReDim Bilibili_Favlist_Create(i, 2)
    i = 0
    For Each item In lDic
        arr(i, 0) = item("id")
        arr(i, 1) = item("title")
        arr(i, 2) = item("media_count")
        i = i + 1
    Next
    Bilibili_Favlist_Create = arr
    Set dic = Nothing
    Set lDic = Nothing
    Set dDic = Nothing
    Erase arr
End Function

Function Bilibili_Video_Lists(ByVal sID As String, Optional ByVal iCount As Integer = 0) As String()
    Const tUrl As String = "https://api.bilibili.com/x/v3/fav/resource/list?media_id="
    Dim sUrl As String
    Dim dic As Object
    Dim sResult As String
    Dim dDic As Object, cDic As Object, idic As Object
    Dim item, i As Integer, k As Integer, p As Integer, m As Integer, n
    Dim arr() As String
    Dim rUrl As String
    '每页只显示20条数据, 如果不指定数量, 那么先获取数量
    '封面, 标题,bvid,link
    rUrl = "https://space.bilibili.com/" & sID & "/favlist" '必须有这个refer, 否则将会出现403 forbidden
    If iCount = 0 Then
        sUrl = tUrl & sID & "&pn=1&ps=20&keyword=&order=mtime&type=0&tid=0&jsonp=jsonp"
        sResult = HTTP_GetData("GET", sUrl, rUrl)
        If Len(sResult) = 0 Then Exit Function
        Set dic = JsonConverter.ParseJson(sResult)
        Set dDic = dic("data")
        n = dic("info")("media_count")
        If n = 0 Then Exit Function
        k = n \ 20 '每页的数量
        p = n Mod 20
        If p = 0 Then
            m = k
        Else
            m = k + 1
        End If
        
        If m - 1 = 0 Then Exit Function
        For j = 2 To m
           sUrl = tUrl & sID & "&pn=" & csrt(j) & "&ps=20&keyword=&order=mtime&type=0&tid=0&jsonp=jsonp"
            sResult = HTTP_GetData("GET", sUrl, rUrl)
            Set dic = JsonConverter.ParseJson(sResult)
            Set dDic = dic("data")
        Next
        '----------------------已经显示数据的数量
    Else
        k = iCount \ 20 '每页的数量
        p = iCount Mod 20
        If p = 0 Then
        m = k
        Else
        m = k + 1
        End If
    For j = 1 To m
        sUrl = tUrl & sID & "&pn=" & csrt(j) & "&ps=20&keyword=&order=mtime&type=0&tid=0&jsonp=jsonp"
        sResult = HTTP_GetData("GET", sUrl, rUrl)
        Set dic = JsonConverter.ParseJson(sResult)
        Set dDic = dic("data")
    Next
    Set mDic = dDic("medias")
    i = mDic.Count - 1
    ReDim arr(iCount, 3)
    ReDim Bilibili_Favlist_Create(iCount, 3)
    i = 0
    For Each item In lDic
        arr(i, 0) = item("id")
        arr(i, 1) = item("title")
        arr(i, 2) = item("media_count")
        i = i + 1
    Next
    Bilibili_Favlist_Create = arr
    Set dic = Nothing
    Set lDic = Nothing
    Set dDic = Nothing
    Erase arr
End Function

Sub Xueqiu_Analysis()
    Dim arrtext, arrtitle
    Dim arr() As String
    Dim dic As New Dictionary
    Dim i As Integer, k As Integer, m As Integer, j As Integer, ic As Byte, n As Byte
    Dim wb As Workbook
    Dim arr_title() As Byte
    Dim arr_text() As Integer
    Dim c As String
    Dim strTemp As String
    Dim arrbaidu() As String
    Dim arrtem() As String
    
    arrtitle = Array("推荐", "卖出", "持有", "强烈卖出", "强烈推荐", "增持", "中性", "看淡", "长期看好", "符合预期", "财务")
    arrtext = Array("推荐", "卖出", "持有", "强烈卖出", "强烈推荐", "增持", "中性", "看淡", "长期看好", "债务", "稳健", "亏损", "奎纳纳", "巨亏", "利息", "财务费用", "澳大利亚", "澳洲", "延期", "融资", "超预期", "SQM", "智利", "疫情", "符合预期", "风险", "现金流")
    m = UBound(arrtext)
    n = UBound(arrtitle)
    ReDim arr_title(n)
    ReDim arr_text(m)
    c = "acw_tc=2760820715869420684972450eb93b6e8f8e0128ffb03627f7af01490a7279; device_id=152ec66dc3b609ec1aa5b3d4f899b494; aliyungf_tc=AQAAAPdsDDjBWwEAvxQmG96NhqSdMidu; xq_a_token=48575b79f8efa6d34166cc7bdc5abb09fd83ce63; xqat=48575b79f8efa6d34166cc7bdc5abb09fd83ce63; xq_r_token=7dcc6339975b01fbc2c14240ce55a3a20bdb7873; xq_id_token=eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJ1aWQiOi0xLCJpc3MiOiJ1YyIsImV4cCI6MTU4OTY4MjczMCwiY3RtIjoxNTg4MjE1NjY1NTY0LCJjaWQiOiJkOWQwbjRBWnVwIn0.cEusti02SULvgdNCHO92km374SKOybvClNud3af53nb-97oaaYsKdUK84vsmshhUZPXoQlu87IzrPIZlTwZ1VXeHJ-nQB8OmpXbFU3GVivO22B4dJbZ8EQtR-KhWkToTtZElvpRCAHZNCPkfiUd3cuM5OLcB9BtNlj4FY3xNzleor3qcM-QubYQExKqLcOF0FcLbHAojExGW1gKZk1fBrAdLbDUwJDW6qA0gVoQrHBc0EDiwFovOG9t237LsUOT06CajrCYDC8yswFzUcAoe5eqp8IUPqw6n8F2KoPdIL7ACZPeIQwR7f_Pf7-JKQe4xEaKBFhdOxNCue54rQ9cj8w; u=261588215672097"
    arr = Xueqiu_Company_Research("002466", c)
    If isReady = False Then Exit Sub
    i = UBound(arr, 2)
    For k = 0 To i
        If dic.Exists(arr(3, k)) = False Then dic.Add arr(3, k), 1 Else dic(arr(3, k)) = dic(arr(3, k)) + 1 '基金公司名称统计出现的次数
        For j = 0 To n
            If InStr(1, arr(0, k), arrtitle(j), vbBinaryCompare) > 0 Then
                ic = UBound(Split(arr(0, k), arrtitle(j), , vbBinaryCompare)) '标题指定的关键词出现的次数
                arr_title(j) = arr_title(j) + ic
            End If
        Next
        For j = 0 To m
            If InStr(1, arr(1, k), arrtext(j), vbBinaryCompare) > 0 Then '内容指定关键出现的次数
                ic = UBound(Split(arr(1, k), arrtext(j), , vbBinaryCompare))
                arr_text(j) = arr_text(j) + ic '统计内容中出现
            End If
        Next
    Next
    Set wb = Workbooks.Add
    With wb.Worksheets
        .Add after:=wb.Worksheets(.Count), Count:=2 - .Count '创建2张表
    End With
    With wb
        .Sheets(1).Name = "内容"
        .Sheets(2).Name = "分析"
        With .Sheets(1)
            .Cells(3, 3) = Now
            .Cells(6, 2).Resize(i + 1, 4) = wb.Application.Transpose(arr)
            .Columns.AutoFit
        End With
        With .Sheets(2)
            .Cells(7, 2).Resize(dic.Count, 1) = wb.Application.Transpose(dic.Keys) '基金公司的名称
            .Cells(7, 3).Resize(dic.Count, 1) = wb.Application.Transpose(dic.Items) '基金公司出现的次数
            .Cells(7, 5).Resize(n + 1, 1) = wb.Application.Transpose(arrtitle) '标题关键字
            .Cells(7, 6).Resize(n + 1, 1) = wb.Application.Transpose(arr_title) '标题关键字出现的次数
            .Cells(7, 8).Resize(m + 1, 1) = wb.Application.Transpose(arrtext) '内容关键字
            .Cells(7, 9).Resize(m + 1, 1) = wb.Application.Transpose(arr_text) '内容关键字出现的次数
            dic.RemoveAll
            For j = 0 To i
                DoEvents
                isReady = True
                arrbaidu = Baidu_TextAnalysis_API(arr(1, j), True)
                If isReady = True Then
                    m = UBound(arrbaidu)
                    For k = 0 To m
                        If dic.Exists(arrbaidu(k, 0)) = False Then
                            dic.Add arrbaidu(k, 0), Int(arrbaidu(j, 1))
                        Else
                            dic(arrbaidu(k, 0)) = dic(arrbaidu(k, 0)) + Int(arrbaidu(k, 1))
                        End If
                    Next
                End If
            Next
            m = dic.Count
            .Cells(7, 11).Resize(m, 1) = wb.Application.Transpose(dic.Keys)
            .Cells(7, 12).Resize(m, 1) = wb.Application.Transpose(dic.Items)
            .Columns.AutoFit
        End With
        Sleep 500
        .SaveAs "C:\Users\adobe\Desktop\002460_Research_Report.xlsx"
    End With
    Set wb = Nothing
    Set dic = Nothing
End Sub

Function Xueqiu_Company_Research(ByVal sID As String, ByVal cCookie As String, Optional ByVal iPage As Byte = 6) As String() '获取雪球上公司的研报
    Const url As String = "https://xueqiu.com/statuses/stock_timeline.json?symbol_id="
    Dim sUrl As String
    Dim i As Integer
    Dim sResult As String
    Dim rUrl As String
    Dim dic As Object
    Dim lDic As Object
    Dim item
    Dim arr() As String, p As Integer
    Dim k As Integer, strTemp As String
    Dim mFlag As String
    
    '总共获取前6页的数据, count的参数也可以该, 大概最多一次可以获取40条的数据
    '获取标题, 时间,文本,发布的证券公司
    strTemp = Left$(sID, 1)
    If strTemp = "6" Then
        mFlag = "SH"
    Else
        mFlag = "SZ"
    End If
    rUrl = "https://xueqiu.com/S/" & mFlag & sID
    k = -1: p = 0
    For i = 1 To iPage
        isReady = True
        sUrl = url & mFlag & sID & "&count=10&source=%E7%A0%94%E6%8A%A5&page=" & CStr(i)
        sResult = HTTP_GetData("GET", sUrl, rUrl, sCookie:=cCookie)
        If isReady = True Then
            If InStr(1, sResult, "error_code", vbBinaryCompare) = 0 Then
                Set dic = JsonConverter.ParseJson(sResult)
                Set lDic = dic("list")
                k = k + lDic.Count
                ReDim Preserve arr(3, k) '二维数组的第一维无法进行redim
                For Each item In lDic
                    strTemp = item("title")
                    arr(0, p) = strTemp '标题部分
                    If InStr(1, strTemp, ChrW(65306), vbBinaryCompare) > 0 Then strTemp = Trim(Split(strTemp, ChrW(65306))(0))
                    strTemp = Right$(strTemp, Len(strTemp) - 1)
                    arr(3, p) = strTemp '发布公司
                    strTemp = item("text")
                    If InStr(1, strTemp, "<br/><br/>", vbBinaryCompare) > 0 Then strTemp = Split(strTemp, "<br/><br/>")(0)
                    strTemp = Replace$(strTemp, "<br/>", ChrW(65307), 1, , vbBinaryCompare)
                    arr(1, p) = strTemp '文本
                    strTemp = item("timeBefore")
                    If Len(strTemp) < 15 Then strTemp = "2020-" & strTemp
                    arr(2, p) = strTemp '时间
                    strTemp = ""
                    p = p + 1
                Next
                Set dic = Nothing
                Set lDic = Nothing
            End If
        End If
    Next
    If p > 0 Then
        ReDim Xueqiu_Company_Research(3, k)
        Xueqiu_Company_Research = arr
        Erase arr
        isReady = True
    Else
        isReady = False
    End If
End Function

'https://aip.baidubce.com/rpc/2.0/nlp/v1/lexer?charset=UTF-8&access_token=24.f9ba9c5241b67688bb4adbed8bc91dec.2592000.1485570332.282335-8574074
'24.8593a6238b8358db3421ad744d0ab7b6.2592000.1590819835.282335-19673825
'默认返回gbk类型的数据
'最长的字符串不允许超过20000
Function Baidu_TextAnalysis_API(ByVal strText As String, Optional ByVal isGBK As Boolean) As String() '使用百度云开放接口
    Const aUrl As String = "https://aip.baidubce.com/rpc/2.0/nlp/v1/lexer"  '分词接口
    Const aToken As String = "24.8593a6238b8358db3421ad744d0ab7b6.2592000.1590819835.282335-19673825" '百度的api的访问token/1个月
    Dim sUrl As String
    Dim sResult As String
    Dim dic As Object, idic As Object
    Dim item
    Dim arr() As String
    Dim strTemp As String
    Dim k As Integer, i As Integer
    Dim cDic As New Dictionary
    
    If Len(strText) = 0 Then Exit Function
    If Len(aToken) = 0 Then Exit Function
    If isGBK = True Then
        sUrl = aUrl & "?access_token=" & aToken '暂时只使用GBK编码
    Else
        sUrl = aUrl & "?charset=UTF-8" & "&access_token=" & aToken
    End If
    isReady = True
    '构造postdata, 这一步是关键, 由于缺少Python灵活的工具,这种混合参数的postdata的发送较为麻烦
    strText = "{" & Chr(34) & "text" & Chr(34) & ": " & Chr(34) & Charter2UnisCode(strText) & Chr(34) & "}"  'strText = "{""text"": ""\u8bcd\u6cd5\u5206\u6790\u95ee\u9898""}"
    sResult = HTTP_GetData("POST", sUrl, cType:="application/json", sCharset:="gb2312", sPostdata:=strText, isBaidu:=True)
    If isReady = False Then Exit Function
    If InStr(1, sResult, "error_code", vbBinaryCompare) > 0 Then isReady = False: Exit Function
    If Len(sResult) = 0 Then isReady = False: Exit Function
    Set dic = JsonConverter.ParseJson(sResult)
    Set idic = dic("items")
    k = idic.Count - 1
    ReDim arr(k): k = 0
    For Each item In idic
        strTemp = item("item")
        If Len(strTemp) > 1 Then
            If strTemp Like "*[一-龥]*" Then '防止多个符号:"
                If cDic.Exists(strTemp) = False Then cDic.Add strTemp, 1 Else cDic(strTemp) = cDic(strTemp) + 1
            End If
        End If
    Next
    k = cDic.Count
    If k > 0 Then
        isReady = True
        k = k - 1
        ReDim Baidu_TextAnalysis_API(k, 1)
        ReDim arr(k, 1)
        For i = 0 To k
            arr(i, 0) = cDic.Keys()(i)
            arr(i, 1) = cDic.Items()(i)
        Next
        Baidu_TextAnalysis_API = arr
    Else
        isReady = False
    End If
    Debug.Print sResult
    Erase arr
    Set cDic = Nothing
    Set dic = Nothing
    Set idic = Nothing
End Function

'datatable9306298这个参数并不重要, 随便使用即可
Sub East_Money_Write() '将数据写入
    Dim arr() As String
    Dim dicn As New Dictionary
    Dim dics As New Dictionary
    Dim dici As New Dictionary
    Dim dicr As New Dictionary
    Dim wb As Workbook
    Dim k As Long, i As Long, j As Long
    Dim t As Single
    
    t = Timer
    DisEvents
    isReady = True
    arr = East_Money
    If isReady = False Then Exit Sub
    k = CLng(50) * 1268 - 1
    dicn.CompareMode = BinaryCompare
    dics.CompareMode = BinaryCompare
    dici.CompareMode = BinaryCompare
    dicr.CompareMode = BinaryCompare
    For i = 0 To k
        If Len(arr(i, 1)) = 0 Then Exit For
        If dicn.Exists(arr(i, 1)) = False Then '公司
            dicn.Add arr(i, 1), 1
        Else
            dicn(arr(i, 1)) = dicn(arr(i, 1)) + 1
        End If
        
        If dics.Exists(arr(i, 3)) = False Then '股票
            dics.Add arr(i, 3), 1
        Else
            dics(arr(i, 3)) = dics(arr(i, 3)) + 1
        End If
        
        If dicr.Exists(arr(i, 5)) = False Then '评级
            dicr.Add arr(i, 5), 1
        Else
            dicr(arr(i, 5)) = dicr(arr(i, 5)) + 1
        End If
        
        If dici.Exists(arr(i, 6)) = False Then '行业
            dici.Add arr(i, 6), 1
        Else
            dici(arr(i, 6)) = dici(arr(i, 6)) + 1
        End If
    Next
    Set wb = Workbooks.Add
    With wb.Worksheets
        .Add after:=wb.Worksheets(.Count), Count:=2 - .Count '创建2张表
    End With
    With wb
        .Sheets(1).Name = "内容"
        .Sheets(1).Cells(2, 2).Resize(k + 1, 8) = arr
        With .Sheets(2)
            .Name = "分析"
            j = dicn.Count
            .Cells(2, 2).Resize(j, 1) = wb.Application.Transpose(dicn.Keys)
            .Cells(2, 3).Resize(j, 1) = wb.Application.Transpose(dicn.Items)
            
            j = dicr.Count
            .Cells(2, 5).Resize(j, 1) = wb.Application.Transpose(dicr.Keys)
            .Cells(2, 6).Resize(j, 1) = wb.Application.Transpose(dicr.Items)
            
            j = dici.Count
            .Cells(2, 8).Resize(j, 1) = wb.Application.Transpose(dici.Keys)
            .Cells(2, 9).Resize(j, 1) = wb.Application.Transpose(dici.Items)
            
            j = dics.Count
            .Cells(2, 11).Resize(j, 1) = wb.Application.Transpose(dics.Keys)
            .Cells(2, 12).Resize(j, 1) = wb.Application.Transpose(dics.Items)
            .Columns.AutoFit
        End With
        .Sheets(1).Columns.AutoFit
    End With
    Erase arr
    Set dicn = Nothing
    Set dici = Nothing
    Set dics = Nothing
    Set dicr = Nothing
    EnEvents
    MsgBox "分析完成: " & Timer - t
    wb.SaveAs ThisWorkbook.Path & "\Reportx.xlsx"
    Set wb = Nothing
End Sub

Function East_Money() As String() '获取东方财富的研报数据
    Const rUrl As String = "http://data.eastmoney.com/report/stock.jshtml" 'http://reportapi.eastmoney.com/report/list?pageSize=100&beginTime=&endTime=&pageNo=1&qType=1 '行业api
    Const tUrl As String = "http://reportapi.eastmoney.com/report/list?cb=datatable9306298&industryCode=*&pageSize=50&industry=*&rating=&ratingChange=&beginTime=2018-05-05&endTime=2020-05-05&pageNo="
    Dim sUrl As String
    Dim sResult As String
    Dim i As Integer
    Dim dic As Object
    Dim arr() As String
    Dim dicdata As Object, item
    Dim p As Long, k As Long
    Dim m As Long
    
    k = CLng(50) * 1268 - 1
    ReDim arr(k, 7)
    For i = 1 To 1268
        isReady = True
        sUrl = tUrl & CStr(i) & "&fields=&qType=0&orgCode=&code=*"
        DoEvents
        sResult = HTTP_GetData("GET", sUrl, rUrl)
        If isReady = True Then
            m = InStr(1, sResult, "(", vbBinaryCompare) + 1
            sResult = Mid$(sResult, m, Len(sResult) - m)
            Set dic = JsonConverter.ParseJson(sResult)
            Set dicdata = dic("data")
            For Each item In dicdata
                arr(p, 0) = item("title") '标题
                arr(p, 1) = item("orgSName") '公司名称
                arr(p, 2) = item("researcher") '研究人员
                arr(p, 3) = item("stockCode") '股票代码
                arr(p, 4) = item("stockName") '涉及股票
                arr(p, 5) = item("sRatingName") '评级名称
                arr(p, 6) = item("indvInduName") '涉及行业
                arr(p, 7) = item("publishDate") '时间
                p = p + 1
            Next
            Set dic = Nothing
            Set dicdata = Nothing
        End If
    Next
    isReady = False
    If p > 0 Then
        ReDim East_Money(k, 6)
        isReady = True
        East_Money = arr
    End If
    Erase arr
End Function

'-----------------------------------------------------------------多条件输入条件时, 空格或者逗号隔开
Function Get_iWencai_Search(ByVal strKey As String) As String() ' 获取同花顺iwencai智能投资的搜索数据
    Const tUrl As String = "http://www.iwencai.com/stockpick/search?source=phone&ts=2&f=2&qs=result_original&querytype=stock&q=kao&&queryarea=all&tid=stockpick&perpage=31536000&Max-Age=91536000&expires=Wed&w="
    Const cCookie As String = "PHPSESSID=2e1fe5f3a02bd4c5f9fdc243c4ead8a6; cid=2e1fe5f3a02bd4c5f9fdc243c4ead8a61589004018; ComputerID=2e1fe5f3a02bd4c5f9fdc243c4ead8a61589004018; WafStatus=0; vvvv=1; v=ApdNKdVGhVhZ2QHUtw_GQg6KJgDk3Gs-RbDvsunEs2bNGLk08az7jlWAfwP6"
    Dim sUrl As String
    Dim oWin As Object
    Dim sResult As String
    Dim arr() As String
    Dim cjHtml As New cJson_html
    Dim strTemp As String
    Dim i As Byte, k As Integer, m As Byte, n As Integer
    
    strKey = ThisWorkbook.Application.EncodeUrl(strKey)
    sUrl = tUrl & strKey
    isReady = True
    sResult = HTTP_GetData("GET", sUrl, sCookie:=cCookie) ' isMobile:=True
    If isReady = False Then Exit Function
    If InStr(1, sResult, "xuangu", vbBinaryCompare) = 0 Then Exit Function
    k = InStr(1, sResult, "title", vbBinaryCompare) + Len("title") + 3
    strTemp = Mid$(sResult, k, 1)
    If strTemp = "]" Then Exit Function '没有找到数据, 如果是空[]
    Set oWin = cjHtml.Json_oParse(sResult)
    oWin.eval "var r=o['xuangu']['blocks'][0]['data']['result']"    '---https://www.w3school.com.cn/js/js_arrays.asp 数组
    oWin.eval "var t=o['xuangu']['blocks'][0]['data']['title']"
    k = oWin.eval("r.length"): i = oWin.eval("r[0].length")
    i = i - 1
    ReDim arr(k, i)
    ReDim Get_iWencai_Search(k, i)
    k = k - 1
    For m = 0 To i
       arr(0, m) = oWin.eval("t[" & m & "]") '获取标题部分
    Next
    For n = 0 To k
        For m = 0 To i
            arr(n + 1, m) = oWin.eval("r[" & n & "][" & m & "]") '获取搜索结果
        Next
    Next
     Get_iWencai_Search = arr
    Set oWin = Nothing
    Set cjHtml = Nothing
End Function

Sub Get_EastMoney_Fund_Rank() '获取东方财富的基金排行
    Const tUrl As String = "http://fund.eastmoney.com/data/rankhandler.aspx?op=ph&dt=kf&ft=all&rs=&gs=0&sc=zzf&st=desc&sd=2017-06-02&ed=2018-06-02&qdii=&tabSubtype=,,,,,&pi=1&pn=50&dx=1&v=0."
    Const rUrl As String = "http://fund.eastmoney.com/data/fundranking.html"
    Dim sResult As String
    Dim sUrl As String
    '------------------------------这里的关键是使用时间戳替换掉参数v后面的内容
    sUrl = tUrl & Get_Timestamp
    sResult = HTTP_GetData("GET", sUrl, rUrl)
End Sub

Sub Weather_2345() '2345天气api接口_备用
Debug.Print HTTP_GetData("GET", "http://tianqi.2345.com/t/wea_history/js/202001/54511_202001.js", "http://tianqi.2345.com/", acLang:="zh-CN,zh;q=0.8", sCharset:="gb2312")
End Sub

Sub Tianapi_api() 'https://www.tianapi.com/ ' 天行数据api接口
Dim s As String
s = HTTP_GetData("GET", "http://api.tianapi.com/txapi/mobilelocal/index?key=e58ec7d052ab3aac787ebd6cd3447ba3&phone=18526790668")
Debug.Print s
End Sub

Sub Test_ecp() '国电_招标公告动态部分的数据
    Dim c As String
    Dim s As String
    Dim sj As Object
    Dim html As Object
    Dim i As Integer, k As Integer
    '-------------------------------https://www.w3school.com.cn/js/jsref_tostring_number.asp
    c = "BIGipServerpool_ecp2_0=QJ6lkdY0LCWYVFQqlHEOFwPG8f3ZUTiQ/EX+eTs9Lam59+XgvweGjr2CQHUqENPklYCCfTMzb0j3BQ==; JSESSIONID=352B56D4E9D36F7C09CD32A4CFE96925"
    s = "var o=" & HTTP_GetData("POST", "https://ecp.sgcc.com.cn/ecp2.0/ecpwcmcore//index/firstPage/getList/1", "https://ecp.sgcc.com.cn/ecp2.0/portal/", oRig:="https://ecp.sgcc.com.cn", _
    sPostdata:="", sCookie:=c, cType:="application/json", cHost:="ecp.sgcc.com.cn")
    Set html = CreateObject("htmlfile")
    Set sj = html.parentwindow
    sj.eval s, "JScript"
    sj.eval "var r=o['resultValue']['list']"
    k = sj.eval("r.length") - 1
    For i = 0 To k
        If sj.eval("r[" & i & "]['children']['1']['menuName']") = "招标公告" Then '这里得到的数据将数字内容定义为double类型, 需要将数据转为字符串
            sj.eval "var i = " & "r[" & i & "]['children']['1']['cldFirstPageMenuId']; var sTime = i.toString()" 'toString() 方法可把一个 Number 对象转换为一个字符串，并返回结果。
            Debug.Print sj.sTime
            Exit For
        End If
    Next
    Set html = Nothing
    Set sj = Nothing
End Sub

Sub Job51_Jobsearch_list(ByVal strText As String, Optional ByVal cCity As String = "040000") '前程无忧, 获取搜索职位结果
    Const tUrl As String = "https://search.51job.com/list/ '040000为城市id, 表示深圳"
    Dim sResult As String
    Dim iPage As Byte
    Dim sUrl As String
    Dim sOR As Object
    Dim oList As Object
    Dim oEl As Object
    Dim arrpn() As String, arrcn() As String, arrps() As String, arrsl() As String, arrUrl() As String, arrdate() As Date
    Dim arrdetail() As String
    Dim i As Integer, k As Integer, p As Integer
    Dim oA As Object
    Dim reDate As Date, item As Object, tDate As Now
    
    tDate = Now
    strText = ThisWorkbook.Application.EncodeUrl(strText)
    i = -1
    sUrl = tUrl & cCity & ",000000,0000,00,9,99," & strText & ",2,"
    For iPage = 1 To 10
        sUrl = sUrl & iPage & ".html"
        sResult = HTTP_GetData("GET", sUrl, sCharset:="gb2312")
        isReady = True
        If isReady = True Then
            WriteHtml sResult
            Set sOR = oHtmlDom.getElementsByclassName("dw_table")
            Set oEl = sOR(0).getElementsByclassName("el")
            i = i + oEl.Length
            ReDim Preserve arrpn(i)
            ReDim Preserve arrcn(i)
            ReDim Preserve arrps(i)
            ReDim Preserve arrsl(i)
            ReDim Preserve arrUrl(i)
            ReDim Preserve arrdate(i)
            For Each item In oEl
                If k = 1 Then
                    reDate = CDate(item.Children(4).innertext)      '职位发布的日期, 如果超过七天, 就跳出
                    If DateDiff("d", reDate, tDate) > 7 Then GoTo OverHandle
                    arrdate(p) = reDate
                    arrpn(p) = item.Children(0).innertext           '职位名称
                    Set oA = item.Children(0).getElementsByTagName("a")
                    arrUrl(p) = oA(0).href                          '招聘具体信息链接
                    arrcn(p) = item.Children(1).innertext           '公司名称
                    arrps(p) = item.Children(2).innertext           '公司所在区域
                    arrsl(p) = item.Children(3).innertext           '薪酬
                    p = p + 1
                Else
                    k = 1
                End If
            Next
            k = 0
        End If
        oHtmlDom.Clear
    Next
OverHandle:
    Set oA = Nothing
    Set oHtmlDom = Nothing
    Set oList = Nothing
    Set oEl = Nothing
    Set sOR = Nothing
    arrdetail = Job51_Detail(arrUrl, p)
End Sub

Private Function Job51_Detail(ByRef arrUrl() As String, ByVal iCount As Integer) As String() '获取招聘页面的详细信息
    Dim arrdetail() As String
    Dim arr() As String
    Dim i As Integer, k As Integer
    Dim sResult As String
    Dim oDetail As Object
    Dim oP As Object, strTemp As String, item As Object
    
    arr = arrUrl
    i = iCount - 1
    ReDim arrdetail(i)
    ReDim Job51_Detail(i)
    For k = 0 To i
        isReady = True
        sResult = HTTP_GetData("GET", arr(k), sCharset:="gb2312")
        If isReady = True Then
            WriteHtml sResult
            Set oDetail = oHtmlDom.getElementsByclassName("bmsg job_msg inbox")
            Set oP = oDetail(0).getElementsByTagName("p")
            For Each item In oP
                strTemp = strTemp & item.innertext
            Next
            arrdetail(k) = strTemp
            strTemp = ""
            oHtmlDom.Clear
        End If
    Next
    Job51_Detail = arrdetail
    Set oHtmlDom = Nothing
    Set oP = Nothing
    Set oDetail = Nothing
End Function

Function Get_String_MD5_from_Web(ByVal strText As String, Optional ByVal rType As Boolean = False) As String '在线md5_js计算字符串hash
    Const tUrl As String = "http://www.cmd5.com/md5.js"
    Dim oJS As Object
    Dim sResult As String
    sResult = HTTP_GetData("GET", tUrl)
    Set oJS = CreateObject("msscriptcontrol.scriptcontrol")
    oJS.Language = "JavaScript"
    oJS.addcode sResult
    If rType = True Then
        sResult = oJS.eval("var str='" & js.CodeObject.hex_md5(strText) & "';str.toUpperCase();") '直接使用js转换大写
    Else
        sResult = oJS.CodeObject.hex_md5(strText)   '小写
    End If
    Get_String_MD5_from_Web = sResult
    Set oJS = Nothing
End Function

Sub TaobaoStock() '获取淘宝商品的库存, 需要登陆后的cookie, 需要refer(商品的主id)
'detailskip.taobao.com, 库存的数据接口
Dim url As String
Dim c As String
c = Cells(1, 1).Value
url = "https://detailskip.taobao.com/service/getData/1/p1/item/detail/sib.htm?itemId=559637432662&sellerId=2860658045&modules=dynStock,qrcode,viewer,price,duty,xmpPromotion,delivery,upp,activity,fqg,zjys,couponActivity,soldQuantity,page,originalPrice,tradeContract&callback=onSibRequestSuccess"
Debug.Print HTTP_GetData("GET", url, "https://item.taobao.com/item.htm?id=559637432662", sCookie:=c)
End Sub

Sub Douban_API() '豆瓣api, 关键在于key
'https://api.douban.com/v2/movie/imdb/tt9683478?apikey=0df993c66c0c636e29ecbb5344252a4a
End Sub

'-----------------------------------------------------------------------------------------------------通用
'typedef enum RefreshConstants {
'    REFRESH_NORMAL = 0,
'    REFRESH_IFEXPIRED = 1,
'    REFRESH_COMPLETELY = 3
'} RefreshConstants;
'----------------------
'Property ReadyState As tagREADYSTATE
'Const READYSTATE_COMPLETE = 4
'Const READYSTATE_INTERACTIVE = 3
'Const READYSTATE_LOADED = 2
'Const READYSTATE_LOADING = 1
'Const READYSTATE_UNINITIALIZED = 0
Private Function Cookie_Generator(ByVal url As String, ByVal iCount As Byte) As String() '通过ie来获取到cookie
    'https://docs.microsoft.com/en-us/previous-versions//aa768363(v=vs.85)?redirectedfrom=MSDN
    'https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa752066%28v%3dvs.85%29
    Dim IE As Object
    Dim iCookie As String
    Dim i As Byte
    Dim t As Long
    Dim rt As Byte
    Dim arr() As String
    'Const tUrl As String = "http://data.10jqka.com.cn/funds/ggzjl/"
    iCount = iCount - 1
    ReDim arr(iCount)
    ReDim Cookie_Generator(iCount)
    Set IE = CreateObject("InternetExplorer.Application")
    With IE
        .Visible = False
        .Silent = True
        For i = 0 To iCount
Redo:
            .Navigate url
            t = timeGetTime
            Do While .readyState <> 4 And timeGetTime - t < 4000 '等待页面加载完成
                DoEvents
            Loop
            .Refresh2 3 '强制清空缓存刷新,以产生新的cookie
            t = timeGetTime
            Do While .readyState <> 4 And timeGetTime - t < 4000
                DoEvents
            Loop
            iCookie = .Document.Cookie
            If Len(iCookie) > 0 Then
                arr(i) = iCookie
                iCookie = ""
                rt = 0
            Else
                If rt < 3 Then rt = rt + 1: GoTo Redo
            End If
        Next
        .Quit
    End With
    Cookie_Generator = arr
    Set IE = Nothing
End Function

Function HTTP_GetData(ByVal sVerb As String, ByVal sUrl As String, Optional ByVal refUrl As String = "https://www.baidu.com", _
Optional ByVal sProxy As String, Optional ByVal sCharset As String = "utf-8", Optional ByVal sPostdata As Variant = "", _
Optional ByVal cType As String = "application/x-www-form-urlencoded", Optional sCookie As String = "", _
Optional ByVal acType As String, Optional ByVal cHost As String, Optional ByVal isRedirect As Boolean = False, Optional ByVal oRig As String, _
Optional ByVal acLang As String, Optional ByVal xReqw As String, Optional ByVal acEncode As String, _
Optional ByVal rsTimeOut As Long = 3000, Optional ByVal cTimeOut As Long = 3000, Optional ByVal sTimeOut As Long = 5000, Optional ByVal rcTimeOut As Long = 3000, _
Optional ByVal IsSave As Boolean, Optional ByVal ReturnRP As Byte, Optional ByVal isBaidu As Boolean, Optional ByVal isMobile As Boolean = False) As String
    '---------ReturnRP返回响应头 0不返回,1返回全部, 2, 获取全部的cookie,3返回单个cookie,3返回编码类型(Optional ByVal cCharset As Boolean = False)
    '--------------------------sVerb为发送的Html请求的方法,sUrl为具体的网址,sCharset为网址对应的字符集编码,sPostData为Post方法对应的发送body
    '- <form method="post" action="http://www.yuedu88.com/zb_system/cmd.php?act=search"><input type="text" name="q" id="edtSearch" size="12" /><input type="submit" value="搜索" name="btnPost" id="btnPost" /></form>
    Dim oWinHttpRQ As Object
    Dim bResult() As Byte
    Dim strTemp  As String
    '----------------------https://blog.csdn.net/tylm22733367/article/details/52596990
    '------------------------------https://msdn.microsoft.com/en-us/library/windows/desktop/aa384106(v=vs.85).aspx
    '------------------------https://docs.microsoft.com/en-us/windows/win32/winhttp/iwinhttprequest-interface
    On Error GoTo ErrHandle
    If LCase$(Left$(sUrl, 4)) <> "http" Then isReady = False: MsgBox "链接不合法", vbCritical, "Warning": Exit Function
    Set oWinHttpRQ = CreateObject("WinHttp.WinHttpRequest.5.1")
    With oWinHttpRQ
        .Option(6) = isRedirect '为 True 时，当请求页面重定向跳转时自动跳转，False 不自动跳转，截取服务端返回的302状态
        '--------------如果不设置禁用重定向,如有道词典无法有效处理post的数据,将会跳转有道翻译的首页,返回不必要的数据
        .setTimeouts rsTimeOut, cTimeOut, sTimeOut, rcTimeOut 'ResolveTimeout, ConnectTimeout, SendTimeout, ReceiveTimeout
        Select Case sVerb
        '----------Specifies the HTTP verb used for the Open method, such as "GET" or "PUT". Always use uppercase as some servers ignore lowercase HTTP verbs.
        Case "GET"
            .Open "GET", sUrl, False '---url, This must be an absolute URL.
        Case "POST"
            .Open "POST", sUrl, False
            .setRequestHeader "Content-Type", cType
        End Select
        If Len(sProxy) > 0 Then '检测格式是否满足要求
            If LCase(sProxy) <> "localhost:8888" Then
            '-------------------注意fiddler无法直接抓取whq的请求, 需要将代理设置为localhost:8888端口
                If InStr(sProxy, ":") > 0 And InStr(sProxy, ".") > 0 Then
                    If UBound(Split(sProxy, ".")) = 3 Then .SetProxy 2, sProxy 'localhost:8888----代理服务器/需要增加错误判断(并不是每一个代理都可用)
                End If
            Else
                .SetProxy 2, sProxy
            End If
        End If
        '-----------------主要应用于伪装成正常的浏览器以规避网站的反爬虫
        If Len(xReqw) > 0 Then .setRequestHeader "X-Requested-With", xReqw
        If Len(acEncode) > 0 Then .setRequestHeader "Accept-Encoding", acEncode
        If Len(acLang) > 0 Then .setRequestHeader "Accept-Language", acLang
        If Len(acType) > 0 Then .setRequestHeader "Accept", acType
        If Len(cHost) > 0 Then .setRequestHeader "Host", cHost
        If Len(oRig) > 0 Then .setRequestHeader "Origin", oRig
        If Len(sCookie) > 0 Then .setRequestHeader "Cookie", sCookie
        If isBaidu = False Then
            .setRequestHeader "Referer", refUrl '伪装从特定的url而来
            .setRequestHeader "User-Agent", Random_UserAgent(isMobile) 'Random_UserAgent '伪造浏览器的ua
        End If
        If sVerb = "POST" Then
            .Send (sPostdata)
        Else
            .Send
        End If
        '---------------这里可以根据返回的错误值来加以判断网页的访问状态,来决定是否需要重新进行访问(如:返回404,那么就不应该再继续访问,403就需要检查是否触发了网站的反爬机制,需要启用代理)
        If .Status <> 200 Then isReady = False: Set oWinHttpRQ = Nothing: Exit Function
        '------------------------------------判断网页内容的字符集的类型
        '---------一般页面多采用UTF-8编码, 如果编码不正确将会有部分的字符出现乱码
        '------------'并不是所有的响应头都会有setcookie
        If ReturnRP > 0 Then '----------获取响应头,判断编码的类型(注意部分的站点如果不加以伪装浏览器,获取的响应头的编码并不是网站的编码,可能是反爬的响应部分的编码)
            strTemp = .getAllResponseHeaders
            Select Case ReturnRP
                Case 1:
                    HTTP_GetData = .getAllResponseHeaders '获取全部的响应头
                Case 2: '------------------------全部cookie
                    Dim xCookie As Variant
                    Dim i As Byte, k As Byte
                    If InStr(1, strTemp, "set-cookie", vbTextCompare) > 0 Then
                        xCookie = Split(strTemp, "Set-Cookie:")
                        i = UBound(xCookie)
                        strTemp = ""
                        For k = 1 To i
                            If InStr(1, xCookie(k), ";", vbBinaryCompare) > 0 Then strTemp = strTemp & Trim(Split(xCookie(k), ";")(0)) & "; " '拼接在一起
                        Next
                        strTemp = Trim$(strTemp)
                        HTTP_GetData = Left$(strTemp, Len(strTemp) - 1)
                    End If
                Case 3: '----------------------------单个cookie,
                    If InStr(1, strTemp, "set-cookie", vbTextCompare) > 0 Then HTTP_GetData = .getResponseHeader("Set-Cookie") '这里如果没有set-cookie将会出现错误
                Case 4: '----------------------------------------------------------编码类型
                    ' -----------.getResponseHeader("Content-Type")
                    If InStr(1, strTemp, "charset=", vbTextCompare) > 0 Then
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
                    End If
            End Select
            If ReturnRP <> 4 Then Set oWinHttpRQ = Nothing: Exit Function
        End If
        bResult = .responseBody '按照指定的字符编码显示
        '-获取返回的字节数组 (用于应付可能潜在的网站的编码问题造成的返回结果乱码)
        HTTP_GetData = ByteHandle(bResult, sCharset, IsSave)
    End With
    Set oWinHttpRQ = Nothing
    Exit Function
ErrHandle:
    If Err.Number = -2147012867 Then MsgBox "无法链接服务器", vbCritical, "Warning!"
    isReady = False
    Set oWinHttpRQ = Nothing
End Function
'---------------------------------------https://www.w3school.com.cn/ado/index.asp
Private Function ByteHandle(ByRef bContent() As Byte, ByVal sCharset As String, Optional ByVal IsSave As Boolean) As String
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
            '-------------------------& Format(Now, "yyyymmddhhmmss") & CStr(RandNumx(10000)) & ".mp3"
            .SaveToFile ThisWorkbook.Path & "\voice.mp3", 2
            .Close
            Set oStream = Nothing
            Exit Function
        End If
        .Position = 0
        .type = adTypeText
        .CharSet = sCharset
         ByteHandle = .ReadText
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
    '-----------------免费的代理,如果仅仅是ping通,并不意味着可以正常放回数据
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

Private Function Random_UserAgent(ByVal isMobile As Boolean, Optional ByVal ForceIE As Boolean = False) As String '随机浏览器伪装/手机-PC
    Dim i As Byte
    Dim UA As String

    If ForceIE = True Then '使用ie
        UA = "Mozilla/5.0 (Windows NT 10.0; WOW64; Trident/7.0; rv:11.0) like Gecko" 'Mozilla/5.0(compatible;MSIE9.0;WindowsNT6.1;Trident/5.0)
    Else
        i = RandNumx(10)
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

Function Unicode2Character(ByVal strText As String) '将Unicode转为文字
    'https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa752599(v=vs.85)
    With CreateObject("htmlfile")
        .Write "<script></script>"
        '--------------------------https://www.w3school.com.cn/jsref/jsref_unescape.asp
        '该函数的工作原理是这样的：通过找到形式为 %xx 和 %uxxxx 的字符序列（x 表示十六进制的数字），用 Unicode 字符 \u00xx 和 \uxxxx 替换这样的字符序列进行解码
        'ECMAScript v3 已从标准中删除了 unescape() 函数，并反对使用它，因此应该用 decodeURI() 和 decodeURIComponent() 取而代之。
        Unicode2Character = .parentwindow.unescape(Replace(strText, "\u", "%u"))
    End With
End Function

Function sUnicode2Character(strText As String) As String '\u30a2\u30e1\u30ea\u30ab\u5927\u7d71\u9818\u9078\u6319\u304c\u307e\u3082\u306a\u304f\u59cb\u307e\u308b
    With CreateObject("MSScriptControl.ScriptControl")
        .Language = "javascript"
        sUnicode2Character = .eval("('" & strText & "').replace(/&#\d+;/g,function(b){return String.fromCharCode(b.slice(2,b.length-1))});")
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
'https://www.w3school.com.cn/jsref/jsref_getTime.asp
'https://www.cnblogs.com/bender/p/3362449.html
'https://www.cnblogs.com/brucemengbm/p/7245040.html
'https://blog.csdn.net/wkj001/article/details/100944856
'https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa741364(v=vs.85)
Function Random_Ten() As String '返回0-9的随机数
    Dim oDom As Object, oWin As Object
    
    Set oDom = CreateObject("htmlfile")
    Set oWin = oDom.parentwindow
    Random_Ten = oWin.eval("parseInt(10*Math.random(),10)") 'oWin.execScript(旧版本)
    Set oDom = Nothing
    Set oWin = Nothing
End Function

Function Get_Timestamp() As String
    Dim oDom As Object, oWin As Object
    '--------------------------------------https://www.runoob.com/jsref/jsref-obj-math.html
    Set oDom = CreateObject("htmlfile")
    Set oWin = oDom.parentwindow
    Get_Timestamp = oWin.eval("new Date().getTime()") '毫秒级别的时间戳
    Set oDom = Nothing
    Set oWin = Nothing
End Function

Function aUnicode2Character(ByVal strText As String) As String 'Unicode转为常规字符串
    Dim oHtml As Object
    Dim oWindow As Object
    
    Set oHtml = CreateObject("htmlfile")
    Set oWindow = oHtml.parentwindow
    With oWindow
        .eval "var sResult='" & strText & "'"
        aUnicode2Character = .sResult
    End With
    Set oWindow = Nothing
    Set oHtml = Nothing
End Function

Function U_Charter2UnisCode(ByVal strText As String) As String '支持全部中英文数字都转为标准的\uxxxx(4位)
    Dim oHtml As Object
    Dim oWindow As Object
    Dim sCode As String
    Dim strTemp As String
    '------------------------------------------https://www.w3school.com.cn/jsref/jsref_slice_string.asp
    If Len(strText) = 0 Then Exit Function
    Set oHtml = CreateObject("htmlfile")
    Set oWindow = oHtml.parentwindow
    sCode = sCode & "function encodeUnicode(str)"
    sCode = sCode & "{"
    sCode = sCode & "var res = [];"
    sCode = sCode & "for (var i = 0; i < str.length; i++) {"                                                    'charcodeat, 可返回指定位置的字符的 Unicode 编码。这个返回值是 0 - 65535 之间的整数
    sCode = sCode & "res[i] = ( " & Chr(34) & "00" & Chr(34) & " + str.charCodeAt(i).toString(16) ).slice(-4);" 'slice表示提取字符串部分, -4表示从后面倒数进行取值, 倒数第4位
    sCode = sCode & "}"
    sCode = sCode & "return " & Chr(34) & "\\u" & Chr(34) & " + res.join(" & Chr(34) & "\\u" & Chr(34) & ");"
    sCode = sCode & "}"
    sCode = sCode & "encodeUnicode (" & Chr(39) & strText & Chr(39) & ")" '注意这里的单双引号的使用， 当原文本同时带有单双引号时， 需要加反斜杠\作为转义
    strTemp = oWindow.eval(sCode)
    U_Charter2UnisCode = strTemp
    Set oHtml = Nothing
    Set oWindow = Nothing
End Function

Function Charter2UnisCode(ByVal strText As String) As String '将常规字符串转为UniCode /注意数字 , 不支持全部都转为\u+4位
    Dim oHtml As Object
    Dim oWindow As Object
    Dim sCode As String
    Dim cReg As New cRegex
    Dim strTemp As String
    Dim arr() As String
    Dim i As Integer, k As Integer
    If Len(strText) = 0 Then Exit Function
'    strText = cReg.ReplaceText(strText, "[\(|\（]?-?[\d]{1,}\.?[\d]{1,}%?[\)|\）]?", "")
'    strText = cReg.ReplaceText(strText, "\\", "")               '注意正则本身的保留字符
'    strText = cReg.ReplaceText(strText, Chr(34), "\" & Chr(34)) '这些字符相当于保留字符,当表达自己的时候需要转义
'    strText = cReg.ReplaceText(strText, Chr(39), "\" & Chr(39))
    arr = cReg.xMatch(strText, "[^\x00-\xff]{1,}") '只提取双字节字符
    k = UBound(arr)
    strText = ""
    For i = 0 To k
        strText = strText & arr(i)
    Next
'    strText = Replace(strText, Chr(10), "", 1, , vbBinaryCompare)
'    strText = Replace(strText, Chr(13), "", 1, , vbBinaryCompare) '剔除掉换行符
    '-------------------------------------------------------------------------------前期数据处理, 否则可能出错
    Set oHtml = CreateObject("htmlfile")
    Set oWindow = oHtml.parentwindow
    sCode = sCode & "function ToUnisCode(str)"
    sCode = sCode & "{"
    sCode = sCode & "return escape(str).replace(/%/g," & Chr(34) & "\\" & Chr(34) & ").toLowerCase();" 'g表示全局匹配, i表示不区分大小写, javascript的字符串直接可以调用正则来完成查找, 替换等任务
    sCode = sCode & "}"
    sCode = sCode & "ToUnisCode (" & Chr(39) & strText & Chr(39) & ")" '注意这里的单双引号的使用， 当原文本同时带有单双引号时， 需要加反斜杠\作为转义
    strTemp = oWindow.eval(sCode)
    Charter2UnisCode = strTemp
    Set cReg = Nothing
    Set oHtml = Nothing
    Set oWindow = Nothing
End Function

'------------------------------------------------------------- https://www.bejson.com/convert/ox2str/
Function sUnisCode2Charter(ByVal strText As String) As String '将'\x5f\x63\x68\x61\x6e\x67\x65\x49\x74\x65\x6d\x43\x72\x6f\x73\x73\x4c\x61\x79\x65\x72转为常规字符
    Dim oHtml As Object
    Dim oWindow As Object
    Dim sCode As String
    Dim strTemp As String
    '-------------------------- 不支持中文
    If Len(strText) = 0 Then Exit Function
    Set oHtml = CreateObject("htmlfile")
    Set oWindow = oHtml.parentwindow
    sCode = sCode & "function decode(str)"         '----------------------fromCharCode() 可接受一个指定的 Unicode 值，然后返回一个字符串。
    sCode = sCode & "{"                            '----------------------https://www.runoob.com/jsref/jsref-fromcharcode.html
    sCode = sCode & "return str.replace(/\\x(\w{2})/g,function(_,$1){ return String.fromCharCode(parseInt($1,16)) });"
    sCode = sCode & "}"
    sCode = sCode & "decode (" & Chr(39) & strText & Chr(39) & ")" '注意这里的单双引号的使用， 当原文本同时带有单双引号时， 需要加反斜杠\作为转义
    strTemp = oWindow.eval(sCode)
    sUnisCode2Charter = strTemp
    Set oHtml = Nothing
    Set oWindow = Nothing
End Function

'------------------------------------------------------------- 支持中/英文
Function xUnisCode2Charter(ByVal strText As String) As String '将'\xE5\x85\x84\xE5\xBC\x9F\xE9\x9A\xBE\xE5\xBD\x93 \xE6\x9D\x9C\xE6\xAD\x8C转为常规字符
    Dim oHtml As Object
    Dim oWindow As Object
    Dim sCode As String
    Dim strTemp As String
    'charCodeAt() 方法可返回指定位置的字符的 Unicode 编码。这个返回值是 0 - 65535 之间的整数。
    '方法 charCodeAt() 与 charAt() 方法执行的操作相似，只不过前者返回的是位于指定位置的字符的编码，而后者返回的是字符子串。
    '--------------------------https://www.w3school.com.cn/jsref/jsref_charCodeAt.asp
    If Len(strText) = 0 Then Exit Function
    Set oHtml = CreateObject("htmlfile")
    Set oWindow = oHtml.parentwindow
    sCode = sCode & "function Decode(str)"
    sCode = sCode & "{"
    sCode = sCode & "var temp = '';"
    sCode = sCode & "for(var j=0;j<str.length;j++){"
    sCode = sCode & "j < str.length? (temp += '%' + str.charCodeAt(j).toString(16)) : (temp += str)"
    sCode = sCode & "}"
    sCode = sCode & "var realName = decodeURIComponent(temp);"
    sCode = sCode & "return realName"
    sCode = sCode & "}"
    sCode = sCode & "Decode (" & Chr(39) & strText & Chr(39) & ")" '注意这里的单双引号的使用， 当原文本同时带有单双引号时， 需要加反斜杠\作为转义
    strTemp = oWindow.eval(sCode)
    xUnisCode2Charter = strTemp
    Set oHtml = Nothing
    Set oWindow = Nothing
End Function

'-----------------------------无法处理带有换行符, 保留字符的字符
'-----------------------------执行效率远低于于直接执行vbs的正则
Function Html_Reg_Test(ByVal strText As String, ByVal sPatern As String) As Boolean '利用javascript的正则
    Dim oHtml As Object
    Dim oWin As Object
    Dim sCode As String
    
    Set oHtml = CreateObject("htmlfile")
    Set oWin = oHtml.parentwindow
    sCode = "var str=" & Chr(39) & strText & Chr(39) & ";"
    sCode = sCode & "var cPatern= new RegExp(" & sPatern & ");"
    sCode = sCode & "var sResult = cPatern.test(str);"
    oWin.eval sCode
    Html_Reg_Test = oWin.sResult
    Set oHtml = Nothing
    Set oWin = Nothing
End Function

Function Get_Time_Zone() As String '获取时区
    Dim oHtml As Object
    Dim oWin As Object
    Dim sCode As String
    
    Set oHtml = CreateObject("htmlfile")
    Set oWin = oHtml.parentwindow
    sCode = "var oTime=new Date().getTimezoneOffset()/60;"
    oWin.eval sCode
    Get_Time_Zone = oWin.oTime
    Set oHtml = Nothing
    Set oWin = Nothing
End Function

Function unix_Timestamp2commonTime(ByVal sTime As String) As String 'js转换时间戳为正常时间
    Dim sCode As String
    Dim oHtml As Object
    Dim oWin As Object
    
    Set oHtml = CreateObject("htmlfile")
    Set oWin = oHtml.parentwindow
    sCode = "var time =" & sTime & ";"
    sCode = sCode & "var unixTimestamp = new Date(time*1000);"
    sCode = sCode & "var cmTime = unixTimestamp.toLocaleString()"
    oWin.eval sCode
    sCode = oWin.cmTime
    unix_Timestamp2commonTime = sCode
    Set oHtml = Nothing
    Set oWin = Nothing
End Function




