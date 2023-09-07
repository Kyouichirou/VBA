Attribute VB_Name = "��Ϣ��ȡ"
'These constants and corresponding values indicate HTTP status codes returned by servers on the Internet.
'HTTP_STATUS_CONTINUE        '��ҳ��Ӧ״̬����
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
'1.����ʱ�����
'2.���ص���������
'3.�����ص�����,����,htmlfile,json
'4.����ip����
'5.�첽
'6.���߳�
'post json���͵�����, Ҳ��Խ���վ�����json������Ϊpostdata
'WEBSERVICE ����- https://support.microsoft.com/zh-cn/office/webservice-%E5%87%BD%E6%95%B0-0546a35a-ecc6-4739-aed7-c0b7ce1562c4
'ENCODEURL ����- https://support.microsoft.com/zh-cn/office/encodeurl-%E5%87%BD%E6%95%B0-07c7fb90-7c60-4bff-8687-fac50fe33d0e
'FILTERXML ����- https://support.microsoft.com/zh-cn/office/filterxml-%E5%87%BD%E6%95%B0-4df72efc-11ec-4951-86f5-c1374812f5b7?ui=zh-cn&rs=zh-cn&ad=cn
'http://excel880.com/blog/archives/3527
'https://www.bilibili.com/video/av75931500
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function timeGetTime Lib "winmm.dll" () As Long 'ʱ�� -����ѵ��
Private Declare Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Dim oHtmlDom As Object  '�ĵ�
Dim isReady As Boolean '�ж�ִ�е����
Public Pagexs As Integer '����ҳ��
Dim regEx As Object '������ʽ
Dim oTli As Object '��ȡ��������
Dim xmCookie As String
Dim tkCookie As String 'Ϻ��cookie

 '-----------------------��ȡС˵
Function ObtainPage_Info(ByVal strText As String, ByVal cmCode As Byte) As String() '��ȡҳ����Ϣ(�������-�½�Ŀ¼-ҳ����)
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
        strx = ThisWorkbook.Application.EncodeUrl(strText) 'ת��
        Urlx = searchUrl & strx
    Else
        Urlx = strText
    End If
    sResult = HTTP_GetData(sVerb, Urlx)
    If Len(sResult) = 0 Then Exit Function
    WriteHtml sResult
    If cmCode = 1 Then
        arr = searchData '�������
    Else
        Pagexs = 0
        arr = ObtainLists
        Pagexs = ObtainPages
    End If
    ObtainPage_Info = arr
    ThisWorkbook.Application.ScreenUpdating = True
    Set oHtmlDom = Nothing
End Function

Private Function searchData() As String() '�����������ص�����
    Dim oHtml As Object
    Dim oA As Object
    Dim i As Integer, k As Integer
    Dim arr() As String
    Dim oCrumbs As Object
    
    On Error GoTo ErrHandle
    With oHtmlDom
        Set oHtml = .getElementById("searchText").Children(1) '��վ������������ŵ�λ��
        If oHtml.Children.Length = 0 Then GoTo ErrHandle 'û������������
        Set oA = oHtml.getElementsByTagName("a") '������
        k = oA.Length - 1
        ReDim arr(k, 1)
        ReDim searchData(k, 1)
        For i = 0 To k
            arr(i, 0) = oA(i).Text '�����������ʾ����
            arr(i, 1) = oA(i).href '�����Ӧ�ĳ����� 'https://www.w3school.com.cn/tags/att_a_href.asp
        Next
    End With
    searchData = arr
ErrHandle:
    Set oHtmlDom = Nothing
    Set oHtml = Nothing
    Set oA = Nothing
End Function

Private Function ObtainLists() As String() '����ҳ���½�Ŀ¼
    Dim oHtml As Object
    Dim item As Object
    Dim i As Integer, j As Integer
    Dim oA As Object
    Dim arr() As String
    
    Set oHtml = oHtmlDom.getElementsByTagName("li") '��ȡ�б���Ϣ
    j = oHtml.Length
    If j > 0 Then '��ȡ����Ч����Ϣ
        ReDim arr(j - 1, 1)
        ReDim ObtainLists(j - 1, 1)
        j = 0
        For Each item In oHtml
            Set oA = item.getElementsByTagName("a")
            i = oA.Length
            If i > 0 Then
                i = i - 1
                arr(j, 0) = oA(i).Text '�����������ʾ����
                arr(j, 1) = oA(i).href '�����Ӧ�ĳ�����
                j = j + 1
            End If
        Next
    End If
    ObtainLists = arr
    Set oA = Nothing
    Set oHtml = Nothing
End Function

Private Function ObtainPages() As Integer '��ȡ�ܹ��ж���ҳ
    Dim oHtml As Object
    Dim i As Integer, j As Integer
    Dim oA As Object
    Dim strTemp As String
    
    Set oHtml = oHtmlDom.getElementsByclassName("pagebar") '��ȡҳ��
    i = oHtml.Length
    If i > 0 Then
        i = i - 1
        j = oHtml.item(i).Children.Length - 1
        strTemp = oHtml.item(i).Children.item(j).href '���ҳ�������
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

Private Function MoreDetail() As String() '��ȡ�����鼮����Ϣ
    Dim oHtml As Object
    Dim item As Object
    Dim oA As Object
    Dim arr(4) As String
    
    With oHtmlDom
        For Each item In oHtmlDom.all
            arr(0) = item.Children.item(0).Children.item("description").Content: Exit For '���ݼ��/��Ϣλ�ڵ�һ��item
        Next
        Set oHtml = .getElementsByclassName("bookinfo")
        Set oA = oHtml.getElementsByTagName("em")
        arr(1) = oA.innertext '--------------------------����
        Set oHtml = .getElementsByclassName("stats")
        Set oA = oHtml.getElementsByTagName("a")
        arr(2) = oA(0).href '---------------�����½�����
        arr(3) = oA(0).innertext '----------�½�����
        Set oHtml = .getElementsByclassName("intro") '����ͼƬ
        Set oA = oHtml.getElementsByTagName("img")
        arr(4) = oA(0).href
    End With
    ReDim MoreDetail(4)
    MoreDetail = arr
    Set oA = Nothing
    Set oHtml = Nothing
End Function
'------------------------------------��ȡС˵

'-----------------------------------------------------��ɽ/�е��ʵ�-����
Function sGet_Translation_Youdao(ByVal strText As String, Optional ByVal iType As Byte = 0) As String
    Dim strType As String
    Dim strPost As String
    Dim sResult As String
    Dim strTemp As String
    Dim xError As Integer
    Dim js As Object
    Dim i As Integer, k As Integer             '---------'"http://fanyi.youdao.com/translate_o?smartresult=dict&smartresult=rule" '�����Ҫʹ�ô˽ӿ�, ��Ҫ����е��ʵ�ķ���������
    '---------------------------------m.youdao.comû�з�������,����û�����ݽӿ�,��Ҫ����html����ȡ��Ϣ
    Const Urlx As String = "http://fanyi.youdao.com/translate?smartresult=dict&smartresult=rule&smartresult=ugc&sessionFrom=null"
    Const cUrl As String = "http://fanyi.youdao.com/"
    '�е�sign�ļ��㷽��
    'tsStr=timestamp 'ʱ���13λ
    'saltStr=tsStr & randnumx(9) '����һλ(0-9)�������-14λ
    '����aStr="fanyideskweb"
    'Ҫ��������� strText
    '��һ���仯�����ĳ��� Youdao_Const
    'sign=getmd5hash_string(astr & strText & salt & Youdao_Const) '�ȵ�32λ��hashֵ,�����е���sign
    '--------------------------------------------------------------------------------------------------����ȴ�����ò���,���Ƿ���50 Error
    strType = "&type=AUTO" '& Translate_Type(iType) '�е��ķ������;�����������(ֻ��ѡauto��)
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
        If xError = 50 Then Set js = Nothing: Exit Function '��ȡִ�еĴ������(50Ϊ�޷���ֵ)
        k = .eval("HLA.translateResult[0].length")
        For i = 1 To k
            strTemp = strTemp & .eval("HLA.translateResult[0][" & i - 1 & "].tgt")
            '------------------------------{"type":"FR2ZH_CN","errorCode":0,"elapsedTime":2,"translateResult":[[{"src":"Garde contre les uv","tgt":"���⾯��"}]]}
        Next
    End With
    sGet_Translation_Youdao = strTemp
    Set js = Nothing
 End Function
'�е��ĳ����ı仯��С�� bv��ֵ������ָ��,��������ָ��һ��md5ֵ
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
    eStr = ThisWorkbook.Application.EncodeUrl(strText) '���������ݽ��б���
    sTime = Get_Timestamp 'ʱ��� '---������ʱ�����ʱ���Ӻ�����
    Salt = sTime & Random_Ten 'ʱ���+0-9�������
    strTemp = aSign & strText & Salt & cSign
    sign = LCase(GetMD5Hash_String(strTemp)) '����ǩ��
    '------------------------------------------------------�ƽ��е��ķ���
    PostData = "i=" & eStr & "&from=" & sFrom & "&to=" & sTo & sAssist & Salt & "&sign=" & sign & "&ts=" & sTime & cAssist
    isReady = True
    sResult = HTTP_GetData("POST", sUrl, rUrl, sCookie:=cCookie, sPostdata:=PostData)
    If isReady = False Then Exit Function
    strTemp = cReg.sMatch(sResult, Chr(34) & "errorCode" & Chr(34) & ":\d{1,}")
    If Split(strTemp, ":")(1) = "50" Then Set cReg = Nothing: Exit Function '50��ʾ���ص���������
    strTemp = cReg.sMatch(sResult, Chr(34) & "tgt" & Chr(34) & ":.*?,")
    strTemp = Split(strTemp, ":", , vbBinaryCompare)(1)
    strTemp = Mid$(strTemp, 2, Len(strTemp) - 3)
    Get_Youdao_Translate = strTemp
    Set cReg = Nothing
End Function

Private Function Translate_Type(ByVal IorY As Boolean, Optional ByVal iType As Byte = 0, Optional ByVal nType As Byte = 0) As String 'ѡ���������
    Dim sType As String
    Dim fType As String
    Dim tType As String
    If IorY = True Then
        Select Case iType
            Case 1: sType = "en2zh-CHS"
            Case 2: sType = "zh-CHS2en"
            Case 3: sType = "ja2zh-CHS"
            Case 4: sType = "zh-CHS2ja" '"&from=zh-CHS&to=ja" '
            Case Else: sType = "AUTO2AUTO" '����ҳ�Զ��б����Ե�����
        End Select
    Else
        Select Case iType
            Case 1: fType = "zh"
            Case 2: fType = "en"
            Case 3: fType = "ja"
            Case 4: fType = "ko" '"&from=zh-CHS&to=ja" '
            Case 5: fType = "fr"
            Case Else: fType = "auto": iType = 0 '����ҳ�Զ��б����Ե�����
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

Private Function Youdao_Const() As String '��ȡ���ڼ����е�sign�ĳ���,������λ��min.js����
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
        '----------------------------------ƥ������͵�����n.md5("fanyideskweb"+e+i+"Nw(nmmbP%A-r6U3EUn]Aj")}}
        .Pattern = "n\.md5\(\" & Chr(34) & ".*\)\}\}\;" '����\ת��
        .Global = True '�����ִ�Сд
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
'----------------����ʹ�ý�ɽ�ʰ���Ϊ�������ѡ,��ɽ�ʰ�û���е��ʵ���ô�������,֧�ֶ�����,������Ϊ���ŵ�������������޷����������-GoogleҲ����(�и�û�ж��ϵ�©��)
'----------------------------------------------------------------------google translation
'source_code_name:[{code:'auto',name:'�������'},{code:'sq',name:'������������'},{code:'ar',name:'��������'},{code:'am',name:'��ķ������'},{code:'az',name:'�����ݽ���'},{code:'ga',name:'��������'},{code:'et',name:'��ɳ������'},
'{code:'or',name:'��������(��������)'},{code:'eu',name:'��˹����'},{code:'be',name:'�׶���˹��'},{code:'bg',name:'����������'},{code:'is',name:'������'},{code:'pl',name:'������'},{code:'bs',name:'��˹������'},
'{code:'fa',name:'��˹��'},{code:'af',name:'������(�ϷǺ�����)'},{code:'tt',name:'������'},{code:'da',name:'������'},{code:'de',name:'����'},{code:'ru',name:'����'},{code:'fr',name:'����'},{code:'tl',name:'���ɱ���'},
'{code:'fi',name:'������'},{code:'fy',name:'��������'},{code:'km',name:'������'},{code:'ka',name:'��³������'},{code:'gu',name:'�ż�������'},{code:'kk',name:'��������'},{code:'ht',name:'���ؿ���¶���'},{code:'ko',name:'����'},
'{code:'ha',name:'������'},{code:'nl',name:'������'},{code:'ky',name:'������˹��'},{code:'gl',name:'����������'},{code:'ca',name:'��̩��������'},{code:'cs',name:'�ݿ���'},{code:'kn',name:'���ɴ���'},{code:'co',name:'��������'},
'{code:'hr',name:'���޵�����'},{code:'ku',name:'�������'},{code:'la',name:'������'},{code:'lv',name:'����ά����'},{code:'lo',name:'������'},{code:'lt',name:'��������'},{code:'lb',name:'¬ɭ����'},{code:'rw',name:'¬������'},
'{code:'ro',name:'����������'},{code:'mg',name:'�����ʲ��'},{code:'mt',name:'�������'},{code:'mr',name:'��������'},{code:'ml',name:'��������ķ��'},{code:'ms',name:'������'},{code:'mk',name:'�������'},{code:'mi',name:'ë����'},
'{code:'mn',name:'�ɹ���'},{code:'bn',name:'�ϼ�����'},{code:'my',name:'�����'},{code:'hmn',name:'����'},{code:'xh',name:'�Ϸǿ�����'},{code:'zu',name:'�Ϸ���³��'},{code:'ne',name:'�Ჴ����'},{code:'no',name:'Ų����'},
'{code:'pa',name:'��������'},{code:'pt',name:'��������'},{code:'ps',name:'��ʲͼ��'},{code:'ny',name:'��������'},{code:'ja',name:'����'},{code:'sv',name:'�����'},{code:'sm',name:'��Ħ����'},{code:'sr',name:'����ά����'},
'{code:'st',name:'��������'},{code:'si',name:'ɮ٤����'},{code:'eo',name:'������'},{code:'sk',name:'˹�工����'},{code:'sl',name:'˹����������'},{code:'sw',name:'˹��ϣ����'},{code:'gd',name:'�ո����Ƕ���'},
'{code:'ceb',name:'������'},{code:'so',name:'��������'},{code:'tg',name:'��������'},{code:'te',name:'̩¬����'},{code:'ta',name:'̩�׶���'},{code:'th',name:'̩��'},{code:'tr',name:'��������'},{code:'tk',name:'��������'},
'{code:'cy',name:'����ʿ��'},{code:'ug',name:'ά�����'},{code:'ur',name:'�ڶ�����'},{code:'uk',name:'�ڿ�����'},{code:'uz',name:'���ȱ����'},{code:'es',name:'��������'},{code:'iw',name:'ϣ������'},{code:'el',name:'ϣ����'},
'{code:'haw',name:'��������'},{code:'sd',name:'�ŵ���'},{code:'hu',name:'��������'},{code:'sn',name:'������'},{code:'hy',name:'����������'},{code:'ig',name:'������'},{code:'it',name:'�������'},{code:'yi',name:'�������'},
'{code:'hi',name:'ӡ����'},{code:'su',name:'ӡ��������'},{code:'id',name:'ӡ����'},{code:'jw',name:'ӡ��צ����'},{code:'en',name:'Ӣ��'},{code:'yo',name:'Լ³����'},{code:'vi',name:'Խ����'},{code:'zh-CN',name:'����'}]

Function Google_Translation(ByVal strText As String) As String '������Ĵ����滻��(sl=auto)(Ҫ���������) auto or (tl=en)en(�����Ŀ������)
    Const tUrl As String = "http://translate.google.cn/translate_a/single?client=gtx&dt=t&ie=UTF-8&oe=UTF-8&sl=auto&tl=zh-CN&q="
    Dim sResult As String
    Dim strTemp As String
    Dim xResult As Variant
    Dim i As Integer, k As Integer
    
    If Len(strText) > 1024 Then MsgShow "���ݳ��ȳ�����Χ", "Tips", 1200: Exit Function
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
'-------------------------ʹ��TK�ķ���̫��,��Ҫ�����������,������Ҫ����ֵ,ִ���ٶ�ƫ��
Function Get_Google_Translation_TK(ByVal strText As String) '��ȡTKֵ
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
        .Position = 0 '����ʹ��position��ʵ�ֶ��ı��ĸ���д��
        .WriteText sResult
        .Flush
        .SaveToFile FilePath, 2 '����������д���ȥ
        .Close
    End With
    Set oStream = Nothing
    Set regEx = Nothing
    '---------------------https://www.cnblogs.com/qinshou/p/5932274.html
    '��Ҫע�����,��Ҫ��htmlfile��ע����:!-- saved from url=(0013)about:internet --,�������־�����ʾ
    Set oHtml = CreateObject(FilePath)
    '���������д��htmlfile,�޷�����js����
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
        sResult = item.innertext: Exit For '����λ�ڵ�һ��item��
    Next
    Set oHtmlDom = Nothing
    Get_Google_Translation_TKK = Google_Tkk(sResult)
End Function

Private Function Google_Tkk(ByVal strText As String) As String '��ȡgoogle_translation��tkkֵ
    Dim oReg As Object
    Dim Matches As Object
    Dim match As Object
    Set oReg = CreateObject("VBScript.RegExp")
    With oReg
        .Pattern = "([0-9]{6,})+(\.[0-9]{9,})" 'ƥ�䳤��Ϊ6λ��+С����+9λ�������,\ת��, \. ��ʾ������"."
        .Global = True '�����ִ�Сд
        .IgnoreCase = True
        Set Matches = .Execute(strText)
        For Each match In Matches
            Google_Tkk = match.Value: Exit For
        Next
    End With
    Set oReg = Nothing
    Set Matches = Nothing
End Function
'����Ҫ�ֿ������������������Ϊ�鷳,���������������
'total //��׼Ϊ100���ַ�Ϊһ�� Ȼ��total���ܹ�������
'idx//�±�,�ڼ���
'textlen //�������ֵļ���,��������ַ����ĳ���
'q //�ϳɵ���������,���ĵĻ���Ҫת����utf-8��Ϊ����
'tl //���ԣ���ʶ�������һ��
Sub Online_Voice_Google(ByVal strText As String) '�г��ȵ�����
    Dim i As Byte, j As Byte, k As Integer
    Dim tUrl As String, strText As String
    Const gUrl As String = "https://translate.google.cn/translate_tts?ie=UTF-8&client=tw-ob&ttsspeed=1"
    i = 0
    j = 1
    k = Len(strText)
    If k > 144 Then MsgShow "���ȳ�����Χ", "Tips", 1200: Exit Sub
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

Sub Online_Voice(ByVal strText As String, ByVal iType As Byte, Optional ByVal vYoudao As Boolean = False, Optional IsUK As Boolean = False) '��Ӣ��/ ��ʽӢʽ
    Const usUrl As String = "https://fanyi.baidu.com/gettts?lan=en&text=" '����-�ٶ�
    Const ukUrl As String = "https://fanyi.baidu.com/gettts?lan=uk&text="
    Const zhUrl As String = "https://fanyi.baidu.com/gettts?lan=zh&text="
    Const jaUrl As String = "https://fanyi.baidu.com/gettts?lan=jp&text="
    '----------------baidu
    Const yUrl As String = "http://dict.youdao.com/dictvoice?audio=" '�е�
    Dim tUrl As String
    Dim strx As String
    
    strText = Trim(strText)
    If Len(strText) = 0 Then Exit Sub
    strText = ThisWorkbook.Application.EncodeUrl(strText)
    If vYoudao = False Then '����ѡ��ٶ���Ϊ����Դ
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

'urlΪֱ��ת�� , �����ض��ַ���������
'https://fanyi.qq.com/api/tts?platform=PC_Website&lang=zh&text=%E5%A5%B9%E6%98%AF%E4%B8%AA%E5%B0%8F%E7%BE%8E%E5%A5%B3&guid=700b30a9-8bef-4f3f-9b13-238ca7f51f9a
'����ֻʹ��guid��������, guid��Դ��cookie
Sub Get_Voice_fromTencent(ByVal strText As String, ByVal cCookie As String, Optional ByVal iType As Boolean) '����Ѷ�������ȡ����
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
    
    If InStr(strText, "\u") > 0 Then MsgBox "��������������Ƿ��������", vbInformation, "Tips": Exit Function
    isReady = True
    strType = Translate_Type(False, iType, nType) & "&w="
    sPost = strType & ThisWorkbook.Application.EncodeUrl(strText)
    sResult = HTTP_GetData("POST", xUrl, ref, acType:=rType, sPostdata:=sPost) '"&f=zh&t=ja&w="
    If isReady = False Then Exit Function
    Set oJS = CreateObject("scriptcontrol")
    With oJS
        .Language = "jscript"
        .addcode "HLA=" & sResult
        xError = .eval("HLA.content.err_no") '�������
        If xError = 50 Then Set oJS = Nothing: Exit Function
        strTemp = Trim(.eval("HLA.content.out"))
        If Left$(strTemp, 2) = "\u" Then strTemp = Unicode2Character(strTemp) '���ص�������unicode�ַ����磺���ģ�/���ķ���
    End With
    Set oJS = Nothing
    Get_Translation_iCiba = strTemp
End Function
'-----------------------------------------------------------------------------����

'------------------------------------------------��˼¼ETF
Sub GetETF_Lists() '��ȡETF�б���Ϣ
    Dim Urlx As String
    Dim sResult As String
    Dim sVerb As String
    Dim strT As String
    Const tUrl As String = "https://www.jisilu.cn/data/etf/etf_list/?___jsl=LST___t="
    
    DisEvents
    sVerb = "GET"
    strT = TimeStamp 'ʱ���
    Urlx = tUrl & strT & "&rp=25&page=1"
    sResult = HTTP_GetData(sVerb, Urlx)
    ETF_Lists sResult, strT
    Set oHtmlDom = Nothing
    EnEvents
End Sub

Private Function cETF_Lists(ByVal strText As String) As String 'ͨ����ģ����ʵ��ETF��Ϣ��ȡ
    Dim cjs As New cJSON
    Dim objdic As Object
    Dim objrow As Object
    Dim objcell As Object
    Dim objtemp As Object
    Dim k As Integer, i As Integer, p As Integer
    Dim item
    Dim arr() As String
    '���ʹ��jsonconvertor�Ļ�, ��cjsonģ���ʹ������������,�ڻ�ȡ�����ֵ,jsonconvertor�޷�ֱ�ӷ��ض�Ӧ��ֵ,���ǿ���ֱ�ӷ���ֵ������(variant)w = b.Items
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

Private Sub ETF_Lists(ByVal strText As String, ByVal filen As String) '��ȡ��˼¼�ϵ�ETF��Ϣ
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
    
    Set regEx = CreateObject("VBScript.RegExp")             ' ����������ʽ��
    Set objSc = CreateObject("ScriptControl")
    objSc.Language = "JScript"
    strTemp = ReplaceText(strText, "cell", "icell") '�滻���г�ͻ�Ĺؼ���
    strTemp = ReplaceText(strTemp, "rows", "irows")
    strTemp = ReplaceText(strTemp, "page", "ipage")
    strTemp = ReplaceText(strTemp, "null", "-")
    Set objJson = objSc.eval("s=" & strTemp)
    Set objtemp = objJson.irows
    '----------https://docs.microsoft.com/en-us/previous-versions/cc175542(v=vs.90)
    '----------https://docs.microsoft.com/en-us/previous-versions/cc177247%28v%3dvs.90%29
    Set oTli = CreateObject("TLI.TLIApplication")
    '-----------------------�˹��ܲ�ֱ��������x64,��Ҫ�ֶ�ע��dll
    Set objinfo = oTli.InterfaceInfoFromObject(objtemp)
    im = objinfo.Members.Count
    '------------������д�뵽�µĹ�������
    arrTemp = Array("����", "����", "�ּ�", "�Ƿ�(%)", "�ɽ���(��Ԫ)", "ָ��", "ָ��PE", "ָ��PB", "ָ���Ƿ�(%)", "��ֵ", "��ֵ", "��ֵ����", "�����(%)", "��С����(���)", "���з�(%)", "�ݶ�(���)", "��ģ�仯(��Ԫ)", "��ģ(��Ԫ)", "����˾", "Index ID")
    Set wb = Workbooks.Add
    With wb
        .Sheets(1).Name = "ETFLists"
        With .Sheets(1)
            .Cells(3, 1).Resize(1, 20) = arrTemp '��ͷ
            .Range("a3:t3").HorizontalAlignment = xlCenter '����
            .Range("a3:t3").Font.Bold = True
            '------------------��ӳ�����
            .Hyperlinks.Add Anchor:=.Cells(1, 2), address:= _
            "https://www.jisilu.cn/data/etf/", TextToDisplay:="��˼¼"
            .Cells(1, 1) = "������Դ:"
            .Cells(1, 1).Font.Bold = True
            .Cells(2, 1) = "���ݸ���ʱ��:"
            .Cells(2, 1).Font.Bold = True
            .Cells(2, 2) = filen
            .Cells(2, 2).NumberFormatLocal = "000000"
            .Cells(2, 3) = "(ʱ���)"
            For Each objItem In objJson.irows
                Set objcell = objItem.icell
                If j = 0 Then arrt = ObtainObjInfo(objcell): j = 1: k = UBound(arrt): ReDim arrFund(1 To im, 1 To k): i = 1 '���鲻��0��ʼ,�������������ֱ�ӷŽ����
                For p = 1 To k
                    If p <> 13 And p <> 14 And p <> 17 And p <> 25 Then
                        '-------'��ֹ����null�����
                        'If IsNull(CallByName(objcell, arrT(p), VbGet)) = False Then arrFund(i, p) = CallByName(objcell, arrT(p), VbGet)
                        arrFund(i, p) = CallByName(objcell, arrt(p), VbGet)
                        '-----------------https://docs.microsoft.com/zh-cn/office/vba/Language/Reference/user-interface-help/callbyname-function
                    End If
                Next
                i = i + 1
            Next
            '----------------���������½�������,������Ӧλ��
            For j = 1 To i
                k = ColumnChoice(j)
                If k > 0 Then
                    '-----------------https://docs.microsoft.com/zh-cn/office/vba/api/Excel.WorksheetFunction.Index
                    arrTemp = ThisWorkbook.Application.Index(arrFund, , j) '������Ĳ���(ĳһ��)������ȡд����
                    WriteList wb, k, im, arrTemp
                End If
            Next
            For j = 1 To im '����������Ӹ�ʽ����Ϊ������/��ʾ����˾������
                .Hyperlinks.Add Anchor:=.Cells(j + 3, "s"), address:= _
                arrFund(j, 7), TextToDisplay:=arrFund(j, 6)
            Next
            '----------------------������ʾ���ֵĸ�ʽ
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
            .Columns.AutoFit '�����п�
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

Private Function ColumnChoice(ByVal intx As Byte) As Byte '�����ݷ����Ӧ����
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

Private Sub WriteList(ByVal wbx As Workbook, ByVal Indexi As Byte, ByVal p As Integer, ByRef arrx()) '������д����
    With wbx.Sheets(1)
        .Cells(4, Indexi).Resize(p, 1) = arrx
    End With
End Sub

Private Function ObtainObjInfo(ByVal objx As Object) As String() '��ȡ������������
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

Private Function ReplaceText(ByVal strText As String, ByVal strFind As String, ByVal rpText As String) As String '�����滻
    With regEx
        .Global = True '��Ҫ����ȫ��,���ܽ�ȫ���ҵ���ֵ�滻��
        .Pattern = strFind
        .IgnoreCase = True
        ReplaceText = .Replace(strText, rpText)
    End With
End Function
'---------------------------------------------------------------------------------------------------------------------��˼¼ETF

'-------------------------------------------------------------------------------------����
Function Code_City_Weather(ByVal cName As String) As String() '��ȡ���������б�
    Dim Urlx As String
    Dim sResult As String
    Dim xV As Variant, xva As Variant
    Dim strTemp As String
    Dim arr() As String
    Dim i As Integer, k As Integer
    '--------------------------------�õ��������б�ĳ��еĵ����Ͷ�Ӧ�ı��,���ص����ݽ���,�����ü򵥵��ı�����ʽ��ʵ�����ݵ���ȡ
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
        strTemp = xva(0) '���ڵ�������
        arr(i, 0) = Trim$(Right$(strTemp, Len(strTemp) - 8))
        arr(i, 1) = Trim$(xva(2)) '���ڵ�����
        strTemp = xva(UBound(xva))
        arr(i, 2) = Left$(strTemp, Len(strTemp) - 1) '���ڵ���������
    Next
    Code_City_Weather = arr
End Function
'-----------------------�����й�������û����ط����ȡ���ݵĶ˿�,���������վ��ȡ����
'Function Get_Weather(ByVal cityCode As String, Optional ByVal dType As Byte) As String
'Dim urlx As String
'Dim xtype As String
'If dType = 1 Then
'xtype = "weather" '7��
'Else
'xtype = "weather1d" '1��
'End If
'urlx = "http://www.weather.com.cn/" & xtype & "/" & cityCode & ".shtml"
'End Function
'----------------------------------------------------------------------
Function Get_Weather_fromAPI(ByVal cityCode As String) As String() '��ȡ��������Ϣ /15��
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
    For i = 1 To k    '�Ƚϵ�һ������һ������ 15,12/11
        Set objtemp = objinfo(i)
        j = objtemp.Count
        If i = 1 Then ReDim arr(1 To k, 1 To m): ReDim Get_Weather_fromAPI(1 To k, 1 To m): p = 1 '���ݲ�����ȫһ��
        For Each item In objtemp
            arr(i, p) = objtemp(item)
            p = p + 1
            If j < m Then
                If p = 8 Then p = 9: arr(i, 8) = "-" '������������ȱʧ
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

Function Get_IP_CityCode() As String() '��ȡ����ip��Ӧ�ĵ���������id
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
        arr(k) = Trim$(Replace$(Split(temp(k), "=")(1), Chr(34), "")) 'ip��ַ,���д���, ���ڵ���
    Next
    Get_IP_CityCode = arr
End Function
'--------------------------------------------------------------------------------------------------------����

'---------------------------------����
'��Ҫע����Ƕ������Ϣ��ȡ,ֱ����html����Ϣ����ַ�������Ϣ
'���ǿ���ֱ�ӻ�ȡ��ӦidԪ�ص���Ϣ,ȴ�����ƹ�����ķ�����ʩ
Sub sGet_doubanBook_Tag()
    Dim sTag As Variant
    Dim itemx
    Dim i As Integer, k As Integer
    
    Set sTag = Get_doubanBook_Tag
    For Each itemx In sTag
        Debug.Print itemx   '��ķ���
        i = sTag(itemx).Count
        For k = 1 To i
            Debug.Print sTag(itemx)(k) '��Tag
        Next
    Next
    Set sTag = Nothing
End Sub

Function Get_doubanBook_Tag() As Variant '��ȡ�����鼮����Tag
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

Function Get_doubanBook_Tag_Rank(ByVal tName As String, Optional ByVal rType As Byte = 1, Optional ByVal Pages As Byte = 5) '��ȡ�����ǩ�鼮����
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
    '-------------------ÿҳ20��,ץȡǰ5ҳ(�п�������20)
    If rType = 1 Then
        sType = "S" '��������
    ElseIf rType = 2 Then
        sType = "R" '���ճ�������
    Else
        sType = "" '�ۺ�����
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
    k = CInt(sResult) '--------��ȡҳ�������
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
                arr(k, 2) = ospan.item(0).Children.item(1).innertext '����
                arr(k, 3) = ospan.item(0).Children.item(2).innertext '��������
            Else
                arr(k, 2) = ospan.item(0).Children.item(1).innertext '����
                arr(k, 3) = "-"
            End If
            Set oA = item.getElementsByTagName("a")
            arr(k, 4) = oA(1).href '����
            arr(k, 0) = oA(1).innertext '��Ʒ����
            Set oA = item.ChildNodes.item(3)
            arr(k, 1) = oA.Children.item(1).innertext '����
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

Function Get_doubanBook_Top250() As String() 'ץȡ�������Top250
    Const tUrl As String = "https://book.douban.com/top250?start=" '��0��ʼ��225����
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
            If p > 0 Then '------------������һ���ڵ�
                Set oA = item.getElementsByTagName("img")
                Set ospan = item.getElementsByclassName("star clearfix")
                arr(i, 5) = oA(0).href '��������
                arr(i, 2) = ospan.item(0).Children.item(1).innertext '����
                arr(i, 3) = ospan.item(0).Children.item(2).innertext '��������
                Set oA = item.getElementsByTagName("a")
                arr(i, 4) = oA(1).href '����
                arr(i, 0) = oA(1).innertext '��Ʒ����
                Set oA = item.ChildNodes.item(3)
                arr(i, 1) = oA.Children.item(1).innertext '����
                If oA.Children.Length > 3 Then
                    arr(i, 6) = oA.Children.item(3).innertext '����
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

Function Get_douban_SearchResult(ByVal strText As String, Optional ByVal iType As Byte = 0) As String() '��ȡ������������'ͨ������,��Ӱ����,�鼮����
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
    Const tUrl As String = "https://www.douban.com/search?q=" 'ͨ������
    Const mUrl As String = "https://www.douban.com/search?cat=1002&q=" '��Ӱ����
    Const bUrl As String = "https://www.douban.com/search?cat=1001&q=" '�鼮����
    
    If iType = 1 Then 'ѡ������������
        sUrl = bUrl
    ElseIf iType = 2 Then
        sUrl = mUrl
    Else
        sUrl = tUrl
    End If
    sUrl = sUrl & ThisWorkbook.Application.EncodeUrl(strText)
    sResult = HTTP_GetData("GET", sUrl, "https://www.douban.com")
    WriteHtml sResult
    Set oResult = oHtmlDom.getElementsByclassName("result-list") '��ȡ�������
    If oResult Is Nothing Then Set oHtmlDom = Nothing: Exit Function
    If oResult.Length = 0 Then Set oHtmlDom = Nothing: Exit Function
    i = oResult.item(0).Children.Length - 1
    If i < 1 Then Set oHtmlDom = Nothing: Exit Function
    ReDim arr(i, 6)
    ReDim Get_douban_SearchResult(i, 6)
    i = 0
    For Each item In oResult.item(0).Children
        If item.Classname = "result" Then '����
            Set oTitle = item.getElementsByclassName("title")
            Set ospan = oTitle(0).getElementsByTagName("span")
            strTemp = ospan(0).innertext
            If strTemp = "[�鼮]" Or strTemp = "[��Ӱ]" Or strTemp = "[���Ӿ�]" Then
            '-------------------------------------����,����,����,��������,����,����,����ͼƬ
                Set oA = oTitle(0).getElementsByTagName("a") '����
                arr(i, 0) = strTemp '����
                arr(i, 1) = oA(0).innertext '����
                arr(i, 2) = oA(0).href '����
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
                arr(i, 6) = oTitle(0).getElementsByTagName("img")(0).href 'ͼƬ
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

Function Get_Douban_FilmRank_Type() As String() '��ȡ����ĵ�Ӱ�����id
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
    '����ĵ�Ӱ������0��ʼ, ÿҳ20��
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
    '-����12������,��Ա�б�,����url,����,����ʱ��,�������,����,����,����,��Ʒurl,ͶƱ����
    Set oDic = JsonConverter.ParseJson(sResult)
    Set wb = Workbooks.Add
    m = 1
    For Each item In oDic '��һ��0��ʼ,����1��ʼ
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
        .Cells(1, 1) = "�����Ӱ"
        .Cells(2, 1) = "����"
        .Hyperlinks.Add Anchor:=.Cells(2, 3), address:=ref, TextToDisplay:="����"
        .Cells(2, 2) = sType
        .Cells(3, 1) = "��������"
        .Cells(4, 1) = "������Χ:"
        .Cells(4, 2).NumberFormatLocal = "@"
        .Cells(4, 2) = CStr(iStart) + 1 & "-" & CStr(iLimit)
        .Cells(4, 4) = "��������ʱ��:"
        .Cells(4, 5) = Now
        .Cells(5, 1).Resize(1, 12) = Array("����ID", "����", "����", "����", "��������", "��Ʒ����", "����ʱ��", "��ǩ", "��Ա����", "��Ա�б�", "��Ʒ����", "��Ʒ����")
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
            .Hyperlinks.Add Anchor:=.Cells(r, c), address:=strText, TextToDisplay:="����"
        Else
        .Cells(r, c) = strText
        End If
    End With
End Function
'-------------------------------------------------------------------------------------------------------------------�����Ӱ����

'-------------------------------------------------------------------------------------ͬ��˳
'---------------------------ͬ��˳�ʽ�����-http://data.10jqka.com.cn/funds/ggzjl/
'ʹ��cookie�ƹ�����,���ʹ�õ�һcookie��ȡ�����ﵽ250�ͻ����403 forbidden
Function Get_THS_MoneyFlow() As String 'ͬ��
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

Function Get_THS_MoneyFlow_Async() '�첽
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

Private Function Cookie_Lists(ByVal i As Byte) As String 'cookie�б�
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
'----------------------------------------------------------------------------------------------ͬ��˳
Sub Get_eastmoney_FundLists() '��ȡ�����Ƹ������б�
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
    If isReady = False Then MsgBox "��ȡ����ʧ��", vbCritical, "Warning": Exit Sub
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
                If dic.Exists(strTemp) = False Then dic.Add strTemp, 1 Else dic(strTemp) = dic(strTemp) + 1 'ͳ�Ƹ����������͵�����
            End If
            p = p + 1
        Next
    Next
    Set wb = Workbooks.Add
    With wb.Sheets(1)
        .Name = "Fund_Lists"
        .Cells(1, 2) = "������Դ:"
        .Cells(1, 3) = "�����Ƹ�"
        .Cells(2, 2) = "����ץȡʱ��:"
        .Cells(2, 3) = Now
        .Cells(2, 4) = "��������:" & CStr(n)
        .Cells(2, 5) = "�������:"
        p = dic.Count - 1
        m = 2
        strTemp = ""
        For i = 0 To p '------�ֵ�Ҳ�Ǵ�0��ʼ
            strTemp = strTemp & dic.Keys(i) & ", " 'ע�������dic.Keys��д��, ʵ��д��Ϊdic.Keys()(i), �����������þͿ��Ժ���()
            .Cells(m, 7) = dic.Keys(i) & ":"
            .Cells(m, 8) = dic.Items(i)
            m = m + 1
        Next
        .Cells(2, 6) = Left$(strTemp, Len(strTemp) - 2)
        .Cells(3, 2).Resize(1, 5) = Array("�������", "�����ƴ", "��������", "�������", "����ȫƴ")
        .Range("b3:f3").Font.Bold = True
        .Range("b4:b" & n + 3).NumberFormatLocal = "@" '�Ƚ��������(����000001�������ݵ�)�ĸ�ʽ����Ϊ�ı���, ��Ȼ���ݻᱻExcel�̵�
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
'������ݽӿ����һ��ֻ����ʾ49������, ���ָ������ʾ�����Ϳ���ʾ����������ƥ��,������Ĭ�ϵ�����, ����������Ϊ�鷳

Function Fund_History_Hexun(ByVal sID As String, ByVal StartDate As String, EndDate As String) As String() '��Ѷ�����ݽӿ�
    Dim oTable As Object, oTr As Object
    Dim sResult As String
    Dim item As Object
    Const rUrl As String = "http://jingzhi.funds.hexun.com"
    Const sUrl As String = "http://jingzhi.funds.hexun.com/DataBase/jzzs.aspx?fundcode="
    Dim tUrl As String
    Dim i As Integer, k As Integer, p As InterfaceInfo
    Dim arr() As String
    
    tUrl = sUrl & sID & "&startdate=" & StartDate & "&enddate=" & EndDate '���� "& startdate="���ֿո�, �����ظû����ȫ����ʷ����
    sResult = HTTP_GetData("GET", tUrl, rUrl, sCharset:="gb2312")
    WriteHtml sResult
    Set oTable = oHtmlDom.getElementsByclassName("n_table m_table")  'n_table m_table
    Set oTr = oTable.item(0).getElementsByTagName("tr")
    i = oTr.Length - 2 'ȥ����һ������(��ͷ)
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

Function Fund_Essentials(ByVal fID As String, ByVal dMode As String) '�����Ҫ
    '7��,1����, 3����, ����, 1��
    Const tUrl As String = "https://fund.xueqiu.com/dj/open/fund/growth/"
    Dim sResult As String
    Dim sUrl As String
    Dim arr() As String
    Dim arrTemp() As String
    Dim oRegx As New cRegex
    Dim i As Integer, k As Integer, j As Byte, m As Integer, n As Integer, p As Integer
    Dim wb As Workbook
    Dim idown As Integer, iup As Integer, ikeep As Integer '�ǵ�
    Dim iMax As Double, iMin As Double '���, ��Сֵ
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

Function Fund_Profile(ByVal id As String) 'ͬ��˳
    Const tUrl As String = "http://fund.10jqka.com.cn/"
    Const pUrl As String = "http://fund.10jqka.com.cn/data/client/myfund/" '��ȡ�����Ҫ api
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
    
    arrx = Array(0, 1, 3, 4, 5, 6, 13, 14, 15, 16, 17, 20, 21, 35, 36, 37, 38, 39, 40, 58) '��Ҫ������
    sUrl = pUrl & id
    sResult = HTTP_GetData("GET", sUrl, tUrl)
    arr = oRegx.xSubmatch(sResult, Chr(34) & "(.*?)" & Chr(34) & ":\s?" & Chr(34) & "(.*?)" & Chr(34))
    Set oRegx = Nothing
    k = UBound(arrx)
    ReDim sarrTemp(k)
    For i = 0 To k
        strTemp = arr(arrx(i), 1)
        If InStr(strTemp, "\u") > 0 Then strTemp = Unicode2Character(strTemp) '�����еĺ���ΪUnicode�ַ�
        sarrTemp(i) = strTemp
    Next
    '----------------------------------------------------------------------------'�����Ҫ
    sUrl = tUrl & id & "/portfolioindex.html" '��ȡ�ֲ���� html
    sResult = HTTP_GetData("GET", sUrl, tUrl)
    WriteHtml sResult
    Set oTitle = oHtmlDom.getElementsByclassName("o-title") '�Ȼ�ø��µ���������
    '--------������ڲ����,��ô�ͻ�ȡ���µ�����, ��ȡ�زֹ�,ծȯ
    For Each item In oTitle
        strTemp = item.innertext
        If InStr(strTemp, "�زֹ�") > 0 Then
            If InStr(strTemp, "���ݸ���") > 0 Then date1 = CDate(Trim(Split(strTemp, " ")(1))): d1 = 1
        ElseIf InStr(strTemp, "�ز�ծ") > 0 Then
            If InStr(strTemp, "���ݸ���") > 0 Then date2 = CDate(Trim(Split(strTemp, " ")(1))): d2 = 1
        End If
    Next
    If d1 > 0 And d2 > 0 Then '������1��ֵ
        If date1 > date2 Then
            iMode = 1
        ElseIf date1 = date2 Then '��ȡ˫��ֵ
            iMode = 3
        Else
            iMode = 2
        End If
    ElseIf d1 = 0 And d2 > 0 Then '��ȡ����ֵ
        iMode = 2
    ElseIf d1 > 0 And d2 = 0 Then '��ȡ����ֵ
        iMode = 1
    Else
        Exit Function
    End If
    Set oTitle = Nothing
    '--------------------------��ȡ���ݸ���ʱ������
    Set oList = oHtmlDom.getElementsByclassName("s-list") '���3��Ԫ��,�زֹ�,�ز�ծ,����
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
'----------------------------------------------------------------------Ϻ������
Sub Write_FavoriteLists(ByVal id As String, ByVal Pages As Integer)
'----------------------------Ϻ�׵ĸ����ղ��б���ȱʧ
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

Private Function Get_FavoriteLists_fromXiami(ByVal id As String, Optional ByVal page As String = "1") As String() '��ȡϺ�����ָ���ϲ�������б�
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
    '����,˫����,:����
    sKey = UTF8_URLEncoding(sKey)
    sKey = Replace(sKey, "{", "%7B")
    sKey = Replace(sKey, "}", "%7D")
    sKey = Replace(sKey, Chr(34), "%22")
    sUrl = tUrl & sKey & "&_s=" & sign
    sResult = HTTP_GetData("GET", sUrl, rUrl, sCookie:=xmCookie) '�ȵ�json��ʽ������
    '�ⲿ�ֵ�json���ݽṹ���������ֵ����ݽṹһ��, ���Թ���
    '-----------��ȡ��ͷ�Ĳ��ּ���
    If InStr(1, Left$(sResult, 30), "success", vbTextCompare) = 0 Then Exit Function '�ж�������Ч���ݷ���
    arr = Json_Data_Treat(sResult)
    Get_FavoriteLists_fromXiami = arr
    Erase arr
End Function

Private Function Xiami_Song_Download_Link(ByVal sID As String) As String '��ȡ��������������
    Dim sKey As String
    Dim sign As String
    Dim sResult As String
    Dim sUrl As String
    Dim rUrl As String
    Dim strTemp As String
    Const tUrl As String = "https://www.xiami.com/api/song/getPlayInfo?_q="
    'Ϻ�״����¼��ް�Ȩ����Ʒ, �������404����
    rUrl = "https://www.xiami.com/"
    sKey = "{" & Chr(34) & "songIds" & Chr(34) & ":[" & sID & "]}"
    If Xiami_Pre_Check(sID) = True Then Exit Function
    sign = Xiami_Sign_Generator(tkCookie, tUrl, sKey)
    '����,˫����,:����
    sKey = UTF8_URLEncoding(sKey)
    sKey = Replace(sKey, "{", "%7B")
    sKey = Replace(sKey, "}", "%7D")
    sKey = Replace(sKey, Chr(34), "%22")
    sUrl = tUrl & sKey & "&_s=" & sign
    sResult = HTTP_GetData("GET", sUrl, rUrl, sCookie:=xmCookie)  '�ȵ�json��ʽ������
    If InStr(1, sResult, "success", vbTextCompare) > 0 Then
        strTemp = Trim(Split(sResult, Chr(34) & "listenFile" & Chr(34))(1))
        strTemp = Split(strTemp, ",")(0)
        strTemp = Mid$(strTemp, 3, Len(strTemp) - 3) '��ȡ������õ�����
    End If
    Xiami_Song_Download_Link = strTemp
End Function
'------------------------ͨ��JsonConverterjson����
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
            ReDim Get_Download_Links(i) '��ȡ���е���������
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

Function Xiami_Search(ByVal Keyword As String, Optional ByVal sType As Byte = 0) As String() '��������, ��������,ר������,�赥����
    Const sUrl As String = "https://www.xiami.com/api/search/searchSongs?_q="
    Const arUrl As String = "https://www.xiami.com/api/search/searchArtists?_q="
    Const alUrl As String = "https://www.xiami.com/api/search/searchAlbums?_q="
    Const cUrl As String = "https://www.xiami.com/api/search/searchCollects?_q"
    Dim sKey As String
    Dim sign As String
    Dim sResult As String
    Dim tUrl As String
    Dim arr() As String
    'ע��ʹ�õ�ת�뷽ʽ,https://www.cnblogs.com/qlqwjy/p/9934706.html
    'encodeURIComponent
    'encodeURI
    'application.encodeurlʵ������encodeurlcomponent����
    Select Case sType
        Case 1: tUrl = arUrl
        Case 2: tUrl = cUrl
        Case 3: tUrl = alUrl
        Case Else: tUrl = sUrl
    End Select
    If Xiami_Pre_Check(Keyword) = True Then Exit Function
    sKey = "{" & Chr(34) & "key" & Chr(34) & ":" & Chr(34) & Keyword & Chr(34) & "," & Chr(34) & "pagingVO" & Chr(34) & ":{" & Chr(34) & "page" & Chr(34) & ":1," & Chr(34) & "pageSize" & Chr(34) & ":30}}"
    sign = Xiami_Sign_Generator(tkCookie, tUrl, sKey)
    '����,˫����,:����
    sKey = UTF8_URLEncoding(sKey)
    sKey = Replace(sKey, "{", "%7B")
    sKey = Replace(sKey, "}", "%7D")
    sKey = Replace(sKey, Chr(34), "%22")
    tUrl = tUrl & sKey & "&_s=" & sign
    sResult = HTTP_GetData("GET", url, tUrl, sCookie:=xmCookie) '�ȵ�json��ʽ������
    If InStr(1, Left$(sResult, 30), "success", vbTextCompare) = 0 Then Exit Function '�ж�������Ч���ݷ���
    arr = Json_Data_Treat(sResult)
    Xiami_Search = arr
'    -��������ǻ�ȡ��һ��ֵ, �������id,��ʹ�������Ϊ����
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
'    -���߽�ƥ����ʽ�޸�Ϊ
'    sPartten=chr(40) & chr(34) & "songID" & chr(34) &chr(41) & ":"& chr(40) &"+[\d]{6,}" & chr(41)
'    ʹ��Submatch��������, ����Ҫsplit
End Function
'ר��ID
'ר������
'ר������
'ר������
'ר���ַ���id
'����Alias (����)
'������
'���ַ���
'������: ԭ��
'����
'����id'--------------------------��Ҫ��ȡ�⼸����Ҫ��
Private Function Json_Data_Treat(ByVal strText As String) As String() '������ȡ��json���ݵĴ���
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

Private Function Xiami_Pre_Check(ByVal strText As String) As Boolean '���������������������, ���cookie�Ƿ��ȡ�ɹ�
    Dim i As Byte
    If InStr(1, strText, Chr(34), vbBinaryCompare) > 0 Then
        i = 1
    ElseIf InStr(1, strText, "}", vbBinaryCompare) > 0 Then
        i = 1
    ElseIf InStr(1, strText, "{", vbBinaryCompare) > 0 Then
        i = 1
    End If
    If i = 1 Then MsgBox "���������ַ�", vbInformation, "Tips": Exit Function
    If Xiami_Cookie_Generator = False Then Xiami_Pre_Check = True: MsgBox "��ȡcookieʧ��", vbCritical + vbInformation, "Warning": Exit Function
End Function

Private Function Xiami_Cookie_Generator() As Boolean 'Ϻ������cookie��ȡ
    'ע��cookie��ʹ������,���cookieʧЧ��Ҫ���»�ȡcookie
    Const rUrl As String = "https://www.xiami.com/"
    isReady = True
    If Len(tkCookie) = 0 Then
        xmCookie = HTTP_GetData("GET", "https://www.xiami.com/", ReturnRP:=2)
        If isReady = True Then tkCookie = Split(Split(xmCookie, "; ")(2), "=")(1)
    End If
    Xiami_Cookie_Generator = isReady
End Function

Private Function Xiami_Sign_Generator(ByVal tCookie As String, ByVal pUrl As String, ByVal qStr As String) As String  'Ϻ��sign����
    '��Ҫ����cookie���ƽ�sign�ķ�ʽ���ܻ�ȡ���ӿڵ���Ϣ
    'cookie��xm_sg_tkֵ�ĵ�һ����
    '84c38dbe9481c68a787a781f8534545f_1586662991803, 84c38dbe9481c68a787a781f8534545f�ⲿ��
    '����:"_xmMain_"
    '����: ����url���� https://www.xiami.com/api/favorite/getFavorites, ��"/api/favorite/getFavorites"
    '�������ݵ�_qֵ
    'sign=getmd5hash_string(xm_sg_tk(0) &"_xmMain_"& "/api/favorite/getFavorites" & _q)
    '-------------------------��ǩ����ʽ����������Ϻ��ҳ��
    '�����������Ӳ���Ҫcookie
    Dim strText As String
    If InStr(1, tCookie, "_", vbBinaryCompare) Then tCookie = Split(tCookie, "_")(0)
    If InStr(1, pUrl, "https://www.xiami.com/", vbBinaryCompare) > 0 Then
        pUrl = Split(pUrl, "https://www.xiami.com")(1)
        If InStr(1, pUrl, "?", vbBinaryCompare) > 0 Then pUrl = Split(pUrl, "?")(0)
    End If
    strText = tCookie & "_xmMain_" & pUrl & "_" & qStr
    Xiami_Sign_Generator = LCase(GetMD5Hash_String(strText))
End Function

Function Bilibili_Favlist(ByVal sID As String, Optional ByVal IsCreat As Boolean) As String() 'biliվ��ĸ����ղؼ�
    Const crUrl As String = "https://api.bilibili.com/x/v3/fav/folder/created/list-all?up_mid="
    Const clUrl As String = "https://api.bilibili.com/x/v3/fav/folder/collected/list?pn=1&ps=20&up_mid="
    Dim sUrl As String
    Dim dic As Object
    Dim sResult As String
    Dim dDic As Object, lDic As Object
    Dim item, i As Integer
    Dim arr() As String
    Dim rUrl As String, tUrl As String
    
    '��ȡÿ���ղؼе�id, ����, ���� '"id":949337310,"fid":9493373,"mid":441644010,"attr":22,"title":"Finance","fav_state":0,"media_count":1
    tUrl = IIf(IsCreat = False, crUrl, clUrl)
    sUrl = tUrl & sID & "&jsonp=jsonp"
    rUrl = "https://space.bilibili.com/" & sID & "/favlist" '���������refer, ���򽫻����403 forbidden
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
    'ÿҳֻ��ʾ20������, �����ָ������, ��ô�Ȼ�ȡ����
    '����, ����,bvid,link
    rUrl = "https://space.bilibili.com/" & sID & "/favlist" '���������refer, ���򽫻����403 forbidden
    If iCount = 0 Then
        sUrl = tUrl & sID & "&pn=1&ps=20&keyword=&order=mtime&type=0&tid=0&jsonp=jsonp"
        sResult = HTTP_GetData("GET", sUrl, rUrl)
        If Len(sResult) = 0 Then Exit Function
        Set dic = JsonConverter.ParseJson(sResult)
        Set dDic = dic("data")
        n = dic("info")("media_count")
        If n = 0 Then Exit Function
        k = n \ 20 'ÿҳ������
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
        '----------------------�Ѿ���ʾ���ݵ�����
    Else
        k = iCount \ 20 'ÿҳ������
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
    
    arrtitle = Array("�Ƽ�", "����", "����", "ǿ������", "ǿ���Ƽ�", "����", "����", "����", "���ڿ���", "����Ԥ��", "����")
    arrtext = Array("�Ƽ�", "����", "����", "ǿ������", "ǿ���Ƽ�", "����", "����", "����", "���ڿ���", "ծ��", "�Ƚ�", "����", "������", "�޿�", "��Ϣ", "�������", "�Ĵ�����", "����", "����", "����", "��Ԥ��", "SQM", "����", "����", "����Ԥ��", "����", "�ֽ���")
    m = UBound(arrtext)
    n = UBound(arrtitle)
    ReDim arr_title(n)
    ReDim arr_text(m)
    c = "acw_tc=2760820715869420684972450eb93b6e8f8e0128ffb03627f7af01490a7279; device_id=152ec66dc3b609ec1aa5b3d4f899b494; aliyungf_tc=AQAAAPdsDDjBWwEAvxQmG96NhqSdMidu; xq_a_token=48575b79f8efa6d34166cc7bdc5abb09fd83ce63; xqat=48575b79f8efa6d34166cc7bdc5abb09fd83ce63; xq_r_token=7dcc6339975b01fbc2c14240ce55a3a20bdb7873; xq_id_token=eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJ1aWQiOi0xLCJpc3MiOiJ1YyIsImV4cCI6MTU4OTY4MjczMCwiY3RtIjoxNTg4MjE1NjY1NTY0LCJjaWQiOiJkOWQwbjRBWnVwIn0.cEusti02SULvgdNCHO92km374SKOybvClNud3af53nb-97oaaYsKdUK84vsmshhUZPXoQlu87IzrPIZlTwZ1VXeHJ-nQB8OmpXbFU3GVivO22B4dJbZ8EQtR-KhWkToTtZElvpRCAHZNCPkfiUd3cuM5OLcB9BtNlj4FY3xNzleor3qcM-QubYQExKqLcOF0FcLbHAojExGW1gKZk1fBrAdLbDUwJDW6qA0gVoQrHBc0EDiwFovOG9t237LsUOT06CajrCYDC8yswFzUcAoe5eqp8IUPqw6n8F2KoPdIL7ACZPeIQwR7f_Pf7-JKQe4xEaKBFhdOxNCue54rQ9cj8w; u=261588215672097"
    arr = Xueqiu_Company_Research("002466", c)
    If isReady = False Then Exit Sub
    i = UBound(arr, 2)
    For k = 0 To i
        If dic.Exists(arr(3, k)) = False Then dic.Add arr(3, k), 1 Else dic(arr(3, k)) = dic(arr(3, k)) + 1 '����˾����ͳ�Ƴ��ֵĴ���
        For j = 0 To n
            If InStr(1, arr(0, k), arrtitle(j), vbBinaryCompare) > 0 Then
                ic = UBound(Split(arr(0, k), arrtitle(j), , vbBinaryCompare)) '����ָ���Ĺؼ��ʳ��ֵĴ���
                arr_title(j) = arr_title(j) + ic
            End If
        Next
        For j = 0 To m
            If InStr(1, arr(1, k), arrtext(j), vbBinaryCompare) > 0 Then '����ָ���ؼ����ֵĴ���
                ic = UBound(Split(arr(1, k), arrtext(j), , vbBinaryCompare))
                arr_text(j) = arr_text(j) + ic 'ͳ�������г���
            End If
        Next
    Next
    Set wb = Workbooks.Add
    With wb.Worksheets
        .Add after:=wb.Worksheets(.Count), Count:=2 - .Count '����2�ű�
    End With
    With wb
        .Sheets(1).Name = "����"
        .Sheets(2).Name = "����"
        With .Sheets(1)
            .Cells(3, 3) = Now
            .Cells(6, 2).Resize(i + 1, 4) = wb.Application.Transpose(arr)
            .Columns.AutoFit
        End With
        With .Sheets(2)
            .Cells(7, 2).Resize(dic.Count, 1) = wb.Application.Transpose(dic.Keys) '����˾������
            .Cells(7, 3).Resize(dic.Count, 1) = wb.Application.Transpose(dic.Items) '����˾���ֵĴ���
            .Cells(7, 5).Resize(n + 1, 1) = wb.Application.Transpose(arrtitle) '����ؼ���
            .Cells(7, 6).Resize(n + 1, 1) = wb.Application.Transpose(arr_title) '����ؼ��ֳ��ֵĴ���
            .Cells(7, 8).Resize(m + 1, 1) = wb.Application.Transpose(arrtext) '���ݹؼ���
            .Cells(7, 9).Resize(m + 1, 1) = wb.Application.Transpose(arr_text) '���ݹؼ��ֳ��ֵĴ���
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

Function Xueqiu_Company_Research(ByVal sID As String, ByVal cCookie As String, Optional ByVal iPage As Byte = 6) As String() '��ȡѩ���Ϲ�˾���б�
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
    
    '�ܹ���ȡǰ6ҳ������, count�Ĳ���Ҳ���Ը�, ������һ�ο��Ի�ȡ40��������
    '��ȡ����, ʱ��,�ı�,������֤ȯ��˾
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
                ReDim Preserve arr(3, k) '��ά����ĵ�һά�޷�����redim
                For Each item In lDic
                    strTemp = item("title")
                    arr(0, p) = strTemp '���ⲿ��
                    If InStr(1, strTemp, ChrW(65306), vbBinaryCompare) > 0 Then strTemp = Trim(Split(strTemp, ChrW(65306))(0))
                    strTemp = Right$(strTemp, Len(strTemp) - 1)
                    arr(3, p) = strTemp '������˾
                    strTemp = item("text")
                    If InStr(1, strTemp, "<br/><br/>", vbBinaryCompare) > 0 Then strTemp = Split(strTemp, "<br/><br/>")(0)
                    strTemp = Replace$(strTemp, "<br/>", ChrW(65307), 1, , vbBinaryCompare)
                    arr(1, p) = strTemp '�ı�
                    strTemp = item("timeBefore")
                    If Len(strTemp) < 15 Then strTemp = "2020-" & strTemp
                    arr(2, p) = strTemp 'ʱ��
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
'Ĭ�Ϸ���gbk���͵�����
'����ַ�����������20000
Function Baidu_TextAnalysis_API(ByVal strText As String, Optional ByVal isGBK As Boolean) As String() 'ʹ�ðٶ��ƿ��Žӿ�
    Const aUrl As String = "https://aip.baidubce.com/rpc/2.0/nlp/v1/lexer"  '�ִʽӿ�
    Const aToken As String = "24.8593a6238b8358db3421ad744d0ab7b6.2592000.1590819835.282335-19673825" '�ٶȵ�api�ķ���token/1����
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
        sUrl = aUrl & "?access_token=" & aToken '��ʱֻʹ��GBK����
    Else
        sUrl = aUrl & "?charset=UTF-8" & "&access_token=" & aToken
    End If
    isReady = True
    '����postdata, ��һ���ǹؼ�, ����ȱ��Python���Ĺ���,���ֻ�ϲ�����postdata�ķ��ͽ�Ϊ�鷳
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
            If strTemp Like "*[һ-��]*" Then '��ֹ�������:"
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

'datatable9306298�������������Ҫ, ���ʹ�ü���
Sub East_Money_Write() '������д��
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
        If dicn.Exists(arr(i, 1)) = False Then '��˾
            dicn.Add arr(i, 1), 1
        Else
            dicn(arr(i, 1)) = dicn(arr(i, 1)) + 1
        End If
        
        If dics.Exists(arr(i, 3)) = False Then '��Ʊ
            dics.Add arr(i, 3), 1
        Else
            dics(arr(i, 3)) = dics(arr(i, 3)) + 1
        End If
        
        If dicr.Exists(arr(i, 5)) = False Then '����
            dicr.Add arr(i, 5), 1
        Else
            dicr(arr(i, 5)) = dicr(arr(i, 5)) + 1
        End If
        
        If dici.Exists(arr(i, 6)) = False Then '��ҵ
            dici.Add arr(i, 6), 1
        Else
            dici(arr(i, 6)) = dici(arr(i, 6)) + 1
        End If
    Next
    Set wb = Workbooks.Add
    With wb.Worksheets
        .Add after:=wb.Worksheets(.Count), Count:=2 - .Count '����2�ű�
    End With
    With wb
        .Sheets(1).Name = "����"
        .Sheets(1).Cells(2, 2).Resize(k + 1, 8) = arr
        With .Sheets(2)
            .Name = "����"
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
    MsgBox "�������: " & Timer - t
    wb.SaveAs ThisWorkbook.Path & "\Reportx.xlsx"
    Set wb = Nothing
End Sub

Function East_Money() As String() '��ȡ�����Ƹ����б�����
    Const rUrl As String = "http://data.eastmoney.com/report/stock.jshtml" 'http://reportapi.eastmoney.com/report/list?pageSize=100&beginTime=&endTime=&pageNo=1&qType=1 '��ҵapi
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
                arr(p, 0) = item("title") '����
                arr(p, 1) = item("orgSName") '��˾����
                arr(p, 2) = item("researcher") '�о���Ա
                arr(p, 3) = item("stockCode") '��Ʊ����
                arr(p, 4) = item("stockName") '�漰��Ʊ
                arr(p, 5) = item("sRatingName") '��������
                arr(p, 6) = item("indvInduName") '�漰��ҵ
                arr(p, 7) = item("publishDate") 'ʱ��
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

'-----------------------------------------------------------------��������������ʱ, �ո���߶��Ÿ���
Function Get_iWencai_Search(ByVal strKey As String) As String() ' ��ȡͬ��˳iwencai����Ͷ�ʵ���������
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
    If strTemp = "]" Then Exit Function 'û���ҵ�����, ����ǿ�[]
    Set oWin = cjHtml.Json_oParse(sResult)
    oWin.eval "var r=o['xuangu']['blocks'][0]['data']['result']"    '---https://www.w3school.com.cn/js/js_arrays.asp ����
    oWin.eval "var t=o['xuangu']['blocks'][0]['data']['title']"
    k = oWin.eval("r.length"): i = oWin.eval("r[0].length")
    i = i - 1
    ReDim arr(k, i)
    ReDim Get_iWencai_Search(k, i)
    k = k - 1
    For m = 0 To i
       arr(0, m) = oWin.eval("t[" & m & "]") '��ȡ���ⲿ��
    Next
    For n = 0 To k
        For m = 0 To i
            arr(n + 1, m) = oWin.eval("r[" & n & "][" & m & "]") '��ȡ�������
        Next
    Next
     Get_iWencai_Search = arr
    Set oWin = Nothing
    Set cjHtml = Nothing
End Function

Sub Get_EastMoney_Fund_Rank() '��ȡ�����Ƹ��Ļ�������
    Const tUrl As String = "http://fund.eastmoney.com/data/rankhandler.aspx?op=ph&dt=kf&ft=all&rs=&gs=0&sc=zzf&st=desc&sd=2017-06-02&ed=2018-06-02&qdii=&tabSubtype=,,,,,&pi=1&pn=50&dx=1&v=0."
    Const rUrl As String = "http://fund.eastmoney.com/data/fundranking.html"
    Dim sResult As String
    Dim sUrl As String
    '------------------------------����Ĺؼ���ʹ��ʱ����滻������v���������
    sUrl = tUrl & Get_Timestamp
    sResult = HTTP_GetData("GET", sUrl, rUrl)
End Sub

Sub Weather_2345() '2345����api�ӿ�_����
Debug.Print HTTP_GetData("GET", "http://tianqi.2345.com/t/wea_history/js/202001/54511_202001.js", "http://tianqi.2345.com/", acLang:="zh-CN,zh;q=0.8", sCharset:="gb2312")
End Sub

Sub Tianapi_api() 'https://www.tianapi.com/ ' ��������api�ӿ�
Dim s As String
s = HTTP_GetData("GET", "http://api.tianapi.com/txapi/mobilelocal/index?key=e58ec7d052ab3aac787ebd6cd3447ba3&phone=18526790668")
Debug.Print s
End Sub

Sub Test_ecp() '����_�б깫�涯̬���ֵ�����
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
        If sj.eval("r[" & i & "]['children']['1']['menuName']") = "�б깫��" Then '����õ������ݽ��������ݶ���Ϊdouble����, ��Ҫ������תΪ�ַ���
            sj.eval "var i = " & "r[" & i & "]['children']['1']['cldFirstPageMenuId']; var sTime = i.toString()" 'toString() �����ɰ�һ�� Number ����ת��Ϊһ���ַ����������ؽ����
            Debug.Print sj.sTime
            Exit For
        End If
    Next
    Set html = Nothing
    Set sj = Nothing
End Sub

Sub Job51_Jobsearch_list(ByVal strText As String, Optional ByVal cCity As String = "040000") 'ǰ������, ��ȡ����ְλ���
    Const tUrl As String = "https://search.51job.com/list/ '040000Ϊ����id, ��ʾ����"
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
                    reDate = CDate(item.Children(4).innertext)      'ְλ����������, �����������, ������
                    If DateDiff("d", reDate, tDate) > 7 Then GoTo OverHandle
                    arrdate(p) = reDate
                    arrpn(p) = item.Children(0).innertext           'ְλ����
                    Set oA = item.Children(0).getElementsByTagName("a")
                    arrUrl(p) = oA(0).href                          '��Ƹ������Ϣ����
                    arrcn(p) = item.Children(1).innertext           '��˾����
                    arrps(p) = item.Children(2).innertext           '��˾��������
                    arrsl(p) = item.Children(3).innertext           'н��
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

Private Function Job51_Detail(ByRef arrUrl() As String, ByVal iCount As Integer) As String() '��ȡ��Ƹҳ�����ϸ��Ϣ
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

Function Get_String_MD5_from_Web(ByVal strText As String, Optional ByVal rType As Boolean = False) As String '����md5_js�����ַ���hash
    Const tUrl As String = "http://www.cmd5.com/md5.js"
    Dim oJS As Object
    Dim sResult As String
    sResult = HTTP_GetData("GET", tUrl)
    Set oJS = CreateObject("msscriptcontrol.scriptcontrol")
    oJS.Language = "JavaScript"
    oJS.addcode sResult
    If rType = True Then
        sResult = oJS.eval("var str='" & js.CodeObject.hex_md5(strText) & "';str.toUpperCase();") 'ֱ��ʹ��jsת����д
    Else
        sResult = oJS.CodeObject.hex_md5(strText)   'Сд
    End If
    Get_String_MD5_from_Web = sResult
    Set oJS = Nothing
End Function

Sub TaobaoStock() '��ȡ�Ա���Ʒ�Ŀ��, ��Ҫ��½���cookie, ��Ҫrefer(��Ʒ����id)
'detailskip.taobao.com, �������ݽӿ�
Dim url As String
Dim c As String
c = Cells(1, 1).Value
url = "https://detailskip.taobao.com/service/getData/1/p1/item/detail/sib.htm?itemId=559637432662&sellerId=2860658045&modules=dynStock,qrcode,viewer,price,duty,xmpPromotion,delivery,upp,activity,fqg,zjys,couponActivity,soldQuantity,page,originalPrice,tradeContract&callback=onSibRequestSuccess"
Debug.Print HTTP_GetData("GET", url, "https://item.taobao.com/item.htm?id=559637432662", sCookie:=c)
End Sub

Sub Douban_API() '����api, �ؼ�����key
'https://api.douban.com/v2/movie/imdb/tt9683478?apikey=0df993c66c0c636e29ecbb5344252a4a
End Sub

'-----------------------------------------------------------------------------------------------------ͨ��
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
Private Function Cookie_Generator(ByVal url As String, ByVal iCount As Byte) As String() 'ͨ��ie����ȡ��cookie
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
            Do While .readyState <> 4 And timeGetTime - t < 4000 '�ȴ�ҳ��������
                DoEvents
            Loop
            .Refresh2 3 'ǿ����ջ���ˢ��,�Բ����µ�cookie
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
    '---------ReturnRP������Ӧͷ 0������,1����ȫ��, 2, ��ȡȫ����cookie,3���ص���cookie,3���ر�������(Optional ByVal cCharset As Boolean = False)
    '--------------------------sVerbΪ���͵�Html����ķ���,sUrlΪ�������ַ,sCharsetΪ��ַ��Ӧ���ַ�������,sPostDataΪPost������Ӧ�ķ���body
    '- <form method="post" action="http://www.yuedu88.com/zb_system/cmd.php?act=search"><input type="text" name="q" id="edtSearch" size="12" /><input type="submit" value="����" name="btnPost" id="btnPost" /></form>
    Dim oWinHttpRQ As Object
    Dim bResult() As Byte
    Dim strTemp  As String
    '----------------------https://blog.csdn.net/tylm22733367/article/details/52596990
    '------------------------------https://msdn.microsoft.com/en-us/library/windows/desktop/aa384106(v=vs.85).aspx
    '------------------------https://docs.microsoft.com/en-us/windows/win32/winhttp/iwinhttprequest-interface
    On Error GoTo ErrHandle
    If LCase$(Left$(sUrl, 4)) <> "http" Then isReady = False: MsgBox "���Ӳ��Ϸ�", vbCritical, "Warning": Exit Function
    Set oWinHttpRQ = CreateObject("WinHttp.WinHttpRequest.5.1")
    With oWinHttpRQ
        .Option(6) = isRedirect 'Ϊ True ʱ��������ҳ���ض�����תʱ�Զ���ת��False ���Զ���ת����ȡ����˷��ص�302״̬
        '--------------��������ý����ض���,���е��ʵ��޷���Ч����post������,������ת�е��������ҳ,���ز���Ҫ������
        .setTimeouts rsTimeOut, cTimeOut, sTimeOut, rcTimeOut 'ResolveTimeout, ConnectTimeout, SendTimeout, ReceiveTimeout
        Select Case sVerb
        '----------Specifies the HTTP verb used for the Open method, such as "GET" or "PUT". Always use uppercase as some servers ignore lowercase HTTP verbs.
        Case "GET"
            .Open "GET", sUrl, False '---url, This must be an absolute URL.
        Case "POST"
            .Open "POST", sUrl, False
            .setRequestHeader "Content-Type", cType
        End Select
        If Len(sProxy) > 0 Then '����ʽ�Ƿ�����Ҫ��
            If LCase(sProxy) <> "localhost:8888" Then
            '-------------------ע��fiddler�޷�ֱ��ץȡwhq������, ��Ҫ����������Ϊlocalhost:8888�˿�
                If InStr(sProxy, ":") > 0 And InStr(sProxy, ".") > 0 Then
                    If UBound(Split(sProxy, ".")) = 3 Then .SetProxy 2, sProxy 'localhost:8888----���������/��Ҫ���Ӵ����ж�(������ÿһ����������)
                End If
            Else
                .SetProxy 2, sProxy
            End If
        End If
        '-----------------��ҪӦ����αװ��������������Թ����վ�ķ�����
        If Len(xReqw) > 0 Then .setRequestHeader "X-Requested-With", xReqw
        If Len(acEncode) > 0 Then .setRequestHeader "Accept-Encoding", acEncode
        If Len(acLang) > 0 Then .setRequestHeader "Accept-Language", acLang
        If Len(acType) > 0 Then .setRequestHeader "Accept", acType
        If Len(cHost) > 0 Then .setRequestHeader "Host", cHost
        If Len(oRig) > 0 Then .setRequestHeader "Origin", oRig
        If Len(sCookie) > 0 Then .setRequestHeader "Cookie", sCookie
        If isBaidu = False Then
            .setRequestHeader "Referer", refUrl 'αװ���ض���url����
            .setRequestHeader "User-Agent", Random_UserAgent(isMobile) 'Random_UserAgent 'α���������ua
        End If
        If sVerb = "POST" Then
            .Send (sPostdata)
        Else
            .Send
        End If
        '---------------������Ը��ݷ��صĴ���ֵ�������ж���ҳ�ķ���״̬,�������Ƿ���Ҫ���½��з���(��:����404,��ô�Ͳ�Ӧ���ټ�������,403����Ҫ����Ƿ񴥷�����վ�ķ�������,��Ҫ���ô���)
        If .Status <> 200 Then isReady = False: Set oWinHttpRQ = Nothing: Exit Function
        '------------------------------------�ж���ҳ���ݵ��ַ���������
        '---------һ��ҳ������UTF-8����, ������벻��ȷ�����в��ֵ��ַ���������
        '------------'���������е���Ӧͷ������setcookie
        If ReturnRP > 0 Then '----------��ȡ��Ӧͷ,�жϱ��������(ע�ⲿ�ֵ�վ�����������αװ�����,��ȡ����Ӧͷ�ı��벢������վ�ı���,�����Ƿ�������Ӧ���ֵı���)
            strTemp = .getAllResponseHeaders
            Select Case ReturnRP
                Case 1:
                    HTTP_GetData = .getAllResponseHeaders '��ȡȫ������Ӧͷ
                Case 2: '------------------------ȫ��cookie
                    Dim xCookie As Variant
                    Dim i As Byte, k As Byte
                    If InStr(1, strTemp, "set-cookie", vbTextCompare) > 0 Then
                        xCookie = Split(strTemp, "Set-Cookie:")
                        i = UBound(xCookie)
                        strTemp = ""
                        For k = 1 To i
                            If InStr(1, xCookie(k), ";", vbBinaryCompare) > 0 Then strTemp = strTemp & Trim(Split(xCookie(k), ";")(0)) & "; " 'ƴ����һ��
                        Next
                        strTemp = Trim$(strTemp)
                        HTTP_GetData = Left$(strTemp, Len(strTemp) - 1)
                    End If
                Case 3: '----------------------------����cookie,
                    If InStr(1, strTemp, "set-cookie", vbTextCompare) > 0 Then HTTP_GetData = .getResponseHeader("Set-Cookie") '�������û��set-cookie������ִ���
                Case 4: '----------------------------------------------------------��������
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
        bResult = .responseBody '����ָ�����ַ�������ʾ
        '-��ȡ���ص��ֽ����� (����Ӧ������Ǳ�ڵ���վ�ı���������ɵķ��ؽ������)
        HTTP_GetData = ByteHandle(bResult, sCharset, IsSave)
    End With
    Set oWinHttpRQ = Nothing
    Exit Function
ErrHandle:
    If Err.Number = -2147012867 Then MsgBox "�޷����ӷ�����", vbCritical, "Warning!"
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
    '----------------------����adodb���ֽ�תΪ�ַ���
    Set oStream = CreateObject("ADODB.Stream")
    With oStream
        .Open
        .type = adTypeBinary
        .Write bContent
        If IsSave = True Then '��ȡ����
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

Private Sub WriteHtml(ByVal sHtml As String) '��ҳ����Ϣд��html file
    'https://www.w3.org/TR/DOM-Level-2-HTML/html
    'https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa752574%28v%3dvs.85%29
    '----------------------------------------https://ken3memo.hatenablog.com/entry/20090904/1252025888
    'https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa752573(v=vs.85)
    Set oHtmlDom = CreateObject("htmlfile")
    With oHtmlDom
        .DesignMode = "on" ' �����༭ģʽ(��Ҫֱ��ʹ��.body.innerhtml=shtml,�����ᵼ��IE�������)
        .Write sHtml ' д������
    End With
End Sub

Private Function Random_IP() As String '�ڴ���ip�б��������ѡip/����Ҫ�����ж�ip�Ƿ����
    Dim i As Integer
    Dim arr() As String
    '-----------------��ѵĴ���,���������pingͨ,������ζ�ſ��������Ż�����
    isReady = True
    arr = Proxy_IP
    If isReady = False Then Random_IP = "127.0.0.1:8888": Exit Function '����/ʹ��fiddler
    i = UBound(arr)
    i = RandNumx(i)
    If i = 0 Then i = 1
    Random_IP = arr(i, 1) & ":" & arr(i, 2)
    Set oHtmlDom = Nothing
    Erase arr
End Function
'--------------����(proxy)����
'----------https://docs.microsoft.com/zh-cn/windows/win32/winhttp/iwinhttprequest-setproxy
'HTTPREQUEST_PROXYSETTING_DEFAULT ��0����Default proxy setting. Equivalent to HTTPREQUEST_PROXYSETTING_PRECONFIG.
'HTTPREQUEST_PROXYSETTING_PRECONFIG��0����Indicates that the proxy settings should be obtained from the registry.
'This assumes that Proxycfg.exe has been run. If Proxycfg.exe has not been run and HTTPREQUEST_PROXYSETTING_PRECONFIG is specified, then the behavior is equivalent to HTTPREQUEST_PROXYSETTING_DIRECT.
'HTTPREQUEST_PROXYSETTING_DIRECT��1����Indicates that all HTTP and HTTPS servers should be accessed directly.
'Use this command if there is no proxy server.
'HTTPREQUEST_PROXYSETTING_PROXY��2����When HTTPREQUEST_PROXYSETTING_PROXY is specified, varProxyServer should be set to a proxy server string
'and varBypassList should be set to a domain bypass list string. This proxy configuration applies only to the current instance of the WinHttpRequest object.
Private Function Proxy_IP() As String() '��ȡhttp����ip��ַ�б�
    Dim sResult As String
    Dim oHtml As Object
    Dim objList As Object
    Dim arr() As String
    Dim list_item As Object, item As Object, itemx As Object
    Dim i As Integer, k As Integer
    
    On Error Resume Next
    sResult = HTTP_GetData("GET", "https://www.xicidaili.com/wn/") '��վ����н�Ϊ���еķ��������(ֱ��xmlhttp���ʻ����503���󷵻�)
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
            arr(i, k) = itemx.innertext '1,ip, 2,port, 3,��ַ, 4,����, 5, http/https
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

Private Function Random_UserAgent(ByVal isMobile As Boolean, Optional ByVal ForceIE As Boolean = False) As String '��������αװ/�ֻ�-PC
    Dim i As Byte
    Dim UA As String

    If ForceIE = True Then 'ʹ��ie
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

Function Unicode2Character(ByVal strText As String) '��UnicodeתΪ����
    'https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa752599(v=vs.85)
    With CreateObject("htmlfile")
        .Write "<script></script>"
        '--------------------------https://www.w3school.com.cn/jsref/jsref_unescape.asp
        '�ú����Ĺ���ԭ���������ģ�ͨ���ҵ���ʽΪ %xx �� %uxxxx ���ַ����У�x ��ʾʮ�����Ƶ����֣����� Unicode �ַ� \u00xx �� \uxxxx �滻�������ַ����н��н���
        'ECMAScript v3 �Ѵӱ�׼��ɾ���� unescape() ������������ʹ���������Ӧ���� decodeURI() �� decodeURIComponent() ȡ����֮��
        Unicode2Character = .parentwindow.unescape(Replace(strText, "\u", "%u"))
    End With
End Function

Function sUnicode2Character(strText As String) As String '\u30a2\u30e1\u30ea\u30ab\u5927\u7d71\u9818\u9078\u6319\u304c\u307e\u3082\u306a\u304f\u59cb\u307e\u308b
    With CreateObject("MSScriptControl.ScriptControl")
        .Language = "javascript"
        sUnicode2Character = .eval("('" & strText & "').replace(/&#\d+;/g,function(b){return String.fromCharCode(b.slice(2,b.length-1))});")
    End With
End Function

Private Function oEncodeUrl(ByVal strText As String) As String '���ַ������б���,��Ҫע����Ǹ�����ŵĴ���
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
Function Random_Ten() As String '����0-9�������
    Dim oDom As Object, oWin As Object
    
    Set oDom = CreateObject("htmlfile")
    Set oWin = oDom.parentwindow
    Random_Ten = oWin.eval("parseInt(10*Math.random(),10)") 'oWin.execScript(�ɰ汾)
    Set oDom = Nothing
    Set oWin = Nothing
End Function

Function Get_Timestamp() As String
    Dim oDom As Object, oWin As Object
    '--------------------------------------https://www.runoob.com/jsref/jsref-obj-math.html
    Set oDom = CreateObject("htmlfile")
    Set oWin = oDom.parentwindow
    Get_Timestamp = oWin.eval("new Date().getTime()") '���뼶���ʱ���
    Set oDom = Nothing
    Set oWin = Nothing
End Function

Function aUnicode2Character(ByVal strText As String) As String 'UnicodeתΪ�����ַ���
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

Function U_Charter2UnisCode(ByVal strText As String) As String '֧��ȫ����Ӣ�����ֶ�תΪ��׼��\uxxxx(4λ)
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
    sCode = sCode & "for (var i = 0; i < str.length; i++) {"                                                    'charcodeat, �ɷ���ָ��λ�õ��ַ��� Unicode ���롣�������ֵ�� 0 - 65535 ֮�������
    sCode = sCode & "res[i] = ( " & Chr(34) & "00" & Chr(34) & " + str.charCodeAt(i).toString(16) ).slice(-4);" 'slice��ʾ��ȡ�ַ�������, -4��ʾ�Ӻ��浹������ȡֵ, ������4λ
    sCode = sCode & "}"
    sCode = sCode & "return " & Chr(34) & "\\u" & Chr(34) & " + res.join(" & Chr(34) & "\\u" & Chr(34) & ");"
    sCode = sCode & "}"
    sCode = sCode & "encodeUnicode (" & Chr(39) & strText & Chr(39) & ")" 'ע������ĵ�˫���ŵ�ʹ�ã� ��ԭ�ı�ͬʱ���е�˫����ʱ�� ��Ҫ�ӷ�б��\��Ϊת��
    strTemp = oWindow.eval(sCode)
    U_Charter2UnisCode = strTemp
    Set oHtml = Nothing
    Set oWindow = Nothing
End Function

Function Charter2UnisCode(ByVal strText As String) As String '�������ַ���תΪUniCode /ע������ , ��֧��ȫ����תΪ\u+4λ
    Dim oHtml As Object
    Dim oWindow As Object
    Dim sCode As String
    Dim cReg As New cRegex
    Dim strTemp As String
    Dim arr() As String
    Dim i As Integer, k As Integer
    If Len(strText) = 0 Then Exit Function
'    strText = cReg.ReplaceText(strText, "[\(|\��]?-?[\d]{1,}\.?[\d]{1,}%?[\)|\��]?", "")
'    strText = cReg.ReplaceText(strText, "\\", "")               'ע��������ı����ַ�
'    strText = cReg.ReplaceText(strText, Chr(34), "\" & Chr(34)) '��Щ�ַ��൱�ڱ����ַ�,������Լ���ʱ����Ҫת��
'    strText = cReg.ReplaceText(strText, Chr(39), "\" & Chr(39))
    arr = cReg.xMatch(strText, "[^\x00-\xff]{1,}") 'ֻ��ȡ˫�ֽ��ַ�
    k = UBound(arr)
    strText = ""
    For i = 0 To k
        strText = strText & arr(i)
    Next
'    strText = Replace(strText, Chr(10), "", 1, , vbBinaryCompare)
'    strText = Replace(strText, Chr(13), "", 1, , vbBinaryCompare) '�޳������з�
    '-------------------------------------------------------------------------------ǰ�����ݴ���, ������ܳ���
    Set oHtml = CreateObject("htmlfile")
    Set oWindow = oHtml.parentwindow
    sCode = sCode & "function ToUnisCode(str)"
    sCode = sCode & "{"
    sCode = sCode & "return escape(str).replace(/%/g," & Chr(34) & "\\" & Chr(34) & ").toLowerCase();" 'g��ʾȫ��ƥ��, i��ʾ�����ִ�Сд, javascript���ַ���ֱ�ӿ��Ե�����������ɲ���, �滻������
    sCode = sCode & "}"
    sCode = sCode & "ToUnisCode (" & Chr(39) & strText & Chr(39) & ")" 'ע������ĵ�˫���ŵ�ʹ�ã� ��ԭ�ı�ͬʱ���е�˫����ʱ�� ��Ҫ�ӷ�б��\��Ϊת��
    strTemp = oWindow.eval(sCode)
    Charter2UnisCode = strTemp
    Set cReg = Nothing
    Set oHtml = Nothing
    Set oWindow = Nothing
End Function

'------------------------------------------------------------- https://www.bejson.com/convert/ox2str/
Function sUnisCode2Charter(ByVal strText As String) As String '��'\x5f\x63\x68\x61\x6e\x67\x65\x49\x74\x65\x6d\x43\x72\x6f\x73\x73\x4c\x61\x79\x65\x72תΪ�����ַ�
    Dim oHtml As Object
    Dim oWindow As Object
    Dim sCode As String
    Dim strTemp As String
    '-------------------------- ��֧������
    If Len(strText) = 0 Then Exit Function
    Set oHtml = CreateObject("htmlfile")
    Set oWindow = oHtml.parentwindow
    sCode = sCode & "function decode(str)"         '----------------------fromCharCode() �ɽ���һ��ָ���� Unicode ֵ��Ȼ�󷵻�һ���ַ�����
    sCode = sCode & "{"                            '----------------------https://www.runoob.com/jsref/jsref-fromcharcode.html
    sCode = sCode & "return str.replace(/\\x(\w{2})/g,function(_,$1){ return String.fromCharCode(parseInt($1,16)) });"
    sCode = sCode & "}"
    sCode = sCode & "decode (" & Chr(39) & strText & Chr(39) & ")" 'ע������ĵ�˫���ŵ�ʹ�ã� ��ԭ�ı�ͬʱ���е�˫����ʱ�� ��Ҫ�ӷ�б��\��Ϊת��
    strTemp = oWindow.eval(sCode)
    sUnisCode2Charter = strTemp
    Set oHtml = Nothing
    Set oWindow = Nothing
End Function

'------------------------------------------------------------- ֧����/Ӣ��
Function xUnisCode2Charter(ByVal strText As String) As String '��'\xE5\x85\x84\xE5\xBC\x9F\xE9\x9A\xBE\xE5\xBD\x93 \xE6\x9D\x9C\xE6\xAD\x8CתΪ�����ַ�
    Dim oHtml As Object
    Dim oWindow As Object
    Dim sCode As String
    Dim strTemp As String
    'charCodeAt() �����ɷ���ָ��λ�õ��ַ��� Unicode ���롣�������ֵ�� 0 - 65535 ֮���������
    '���� charCodeAt() �� charAt() ����ִ�еĲ������ƣ�ֻ����ǰ�߷��ص���λ��ָ��λ�õ��ַ��ı��룬�����߷��ص����ַ��Ӵ���
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
    sCode = sCode & "Decode (" & Chr(39) & strText & Chr(39) & ")" 'ע������ĵ�˫���ŵ�ʹ�ã� ��ԭ�ı�ͬʱ���е�˫����ʱ�� ��Ҫ�ӷ�б��\��Ϊת��
    strTemp = oWindow.eval(sCode)
    xUnisCode2Charter = strTemp
    Set oHtml = Nothing
    Set oWindow = Nothing
End Function

'-----------------------------�޷�������л��з�, �����ַ����ַ�
'-----------------------------ִ��Ч��Զ������ֱ��ִ��vbs������
Function Html_Reg_Test(ByVal strText As String, ByVal sPatern As String) As Boolean '����javascript������
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

Function Get_Time_Zone() As String '��ȡʱ��
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

Function unix_Timestamp2commonTime(ByVal sTime As String) As String 'jsת��ʱ���Ϊ����ʱ��
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




