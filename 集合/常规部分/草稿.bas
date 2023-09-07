Attribute VB_Name = "草稿"
'***********资源合集*********"
'https://developer.microsoft.com/zh-cn/windows/downloads/virtual-machines               'win10 集成版本虚拟机
'https://docs.microsoft.com/zh-cn/cpp/mfc/mfc-desktop-applications?view=vs-2019         'MFC
'https://docs.microsoft.com/zh-cn/windows/win32/index                                   'Win32 API
'https://www.qqxiuzi.cn/                                                                '汉字

'Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long
'Private Declare Function PlayWaveSound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszsoundname As String, _
'ByVal uflags As Long) As Long


Sub Get_Data()
'    Dim c As String
'    Dim o As Object
'    Dim a As Object
'    Dim item As Object
'    Dim i As Byte
'    Dim arr(249, 1) As String
'    Dim k As Integer
'    k = 0
'    For i = 0 To 4
'        c = HTTP_GetData("POST", "http://ec.chng.com.cn/ecmall/more.do", "http://ec.chng.com.cn/ecmall/more.do", sPostData:="type=103&searchWay=onTitle&search=&ifend=in&start=" & CStr(i * 50) & "&limit=50")
'        If Len(c) > 0 Then
'            WriteHtml c
'            Set o = oHtmlDom.getElementsByTagName("li")
'            For Each item In o
'            Set a = item.getElementsByTagName("a")
'            arr(k, 0) = a(0).href
'            arr(k, 1) = a(0).innertext
'            k = k + 1
'            Next
'            c = ""
'            Set oHtmlDom = Nothing
'            Set a = Nothing
'        End If
'    Next
'Dim c As String
'Dim o As Object
'Dim item, a
'c = HTTP_GetData("GET", "http://ecp.sgcc.com.cn/html/topic/all/topic00/list_1.html")
'WriteHtml c
'Set o = oHtmlDom.getelementsbyclassname("titleList_01")
'For Each item In o
'Set a = item.getElementsByTagName("a")
'Debug.Print a(0).getAttribute("onclick")
'Debug.Print a(0).href
'Debug.Print a(0).innertext
'Next
'Sub kdkld()
'Dim c As String
'Dim o As Object
'Dim item, a
'Dim i As Byte
'Dim arr(50, 2) As String
'c = HTTP_GetData("GET", "http://ecp.sgcc.com.cn/html/topic/all/topic00/list_1.html")
'WriteHtml c
'Set o = oHtmlDom.getElementsByTagName("li")
'For Each item In o
'    If InStr(item.innerhtml, "titleList_01") > 0 Then 'titleList_bj
'        Stop
'        Set a = item.getElementsByTagName("a")
'        arr(i, 0) = a(0).getAttribute("onclick")
'        arr(i, 1) = a(0).innertext
'        arr(i, 2) = item.Children.item(1).innertext '日期
'        i = i + 1
'    End If
'Next
'Set oHtmlDom = Nothing
'End Sub
End Sub


Option Explicit
Sub SaveHTML()
Dim ch As New cWinHttpRequest

With ch
.Request "https://www.xiami.com/"
.CharSet
.Redirect
.TimeOut
.Referer = "https://www.xiami.com/"
.UserAgent
.Send
Debug.Print .Get_Cookie
End With

Set ch = Nothing
End Sub

Sub dkfk()



End Sub
 Sub fklflgl()
Dim rd As Date
rd = CDate("06-01")
Debug.Print DateDiff("d", rd, Now)
Dim xr As New cRegex
 End Sub
'    For Each item In sgDic
'        For j = 0 To 5
'            If IsNull(item(arrx(j))) = False Then arr(j) = item(arrx(j))
'        Next
'        arr(6)=
'
'
'        For Each itemx In item.Items
'            If p = 0 Then k = item.Count - 1: ReDim arr(i, k): ReDim Get_Collection_Detail(i, k): p = 1
'            If IsObject(itemx) = False Then
'                If IsNull(itemx) = False Then
'                    strTemp = itemx
'                    If Len(strTemp) > 0 Then
'                        arr(n, m) = strTemp
'                    Else
'                        arr(n, m) = "-"
'                    End If
'                Else
'                    arr(n, m) = "-"
'                End If
'                m = m + 1
'            Else
'                strT = item.Keys()(ic)
'                If strT = "listenFiles" Then  '歌词和下载链接
'                j = itemx.Count
'                If j > q Then q = j: k = k + q: ReDim Preserve arr(i, k): ReDim Preserve Get_Collection_Detail(i, k)
'                    For a = 1 To j
'                        If IsNull(itemx.item(a)("listenFile")) = False Then arr(n, m) = itemx.item(a)("listenFile")
'                        m = m + 1
'                    Next
'                ElseIf strT = "lyricInfo" Then
'                    If IsNull(itemx("lyricFile")) = False Then arr(n, m) = itemx("lyricFile"): m = m + 1
'                End If
'            End If
'            ic = ic + 1
'        Next
'        m = 0
'        ic = 0
'        n = n + 1
'    Next
'Sub Test()
'Dim s As String
'Dim i As Long
'Dim matches As Object
'Dim match As Object
'Dim sPattern
'Dim arr() As String
'Dim i As Integer, k As Integer, j As Integer
'Dim cr As Object
'Set cr = CreateObject("VBScript.RegExp")
'sPatern = Array("[\d|\u4e00-\u9fa5]{4,}[\d]{2,}[^\x00-\xff]{6,}", "20[0-2]{1}[\d]{1}-([0-1]{1})?[\d]{1}-([0-3]{1})?[\d]{1}", "showProject\([\d]{1,}\)")
's = HTTP_GetData("GET", "http://sdhrzb.cn/messagesContentAll?pageNo=1&pageSize=10&status=0&pname=&pno=&pstate=")
's = Replace$(s, " ", "", 1, , vbBinaryCompare)
'ReDim arr(30)
'For j = 0 To 2
'    With cr
'        .Pattern = sPattern(j)
'        .Global = True
'        .IgnoreCase = True
'        Set matches = .execute(s)
'    End With
'    i = matches.Count - 1
'    If i >= 0 Then
'        i = 0
'        For Each match In matches
'            arr(i) = match.Value: i = i + 1
'        Next
'        Set matches = Nothing
'    End If
'Next
'Set cr = Nothing
'End Sub
'Sub test()
'Dim s As String
'Dim p As String
'Dim c As String
'c = "BIGipServerpool_ecp2_0=IAvrdc+7DAEmsoIqlHEOFwPG8f3ZUbnybRByPBJ3lNL/Jk9gNDnr3zdE0FY82IZh1EGK/RX8Ua/Mog==; JSESSIONID=F29DD20B27C562070248F92FBABD8542"
'p = "{" & Chr(34) & "index" & Chr(34) & ":1," & Chr(34) & "size" & Chr(34) & ":20," & Chr(34) & "firstPageMenuId" & Chr(34) & ":" & Chr(34) & "20180502001" & Chr(34) & "," & Chr(34) & "orgId" & Chr(34) & ":" & Chr(34) & Chr(34) & "," & Chr(34) & "key" & Chr(34) & ":" & Chr(34) & "" & Chr(34) & "}"
'Debug.Print HTTP_GetData("POST", "https://ecp.sgcc.com.cn/ecp2.0/ecpwcmcore//index/noteList", "https://ecp.sgcc.com.cn/ecp2.0/portal/", oRig:="https://ecp.sgcc.com.cn", cType:="application/json", sCookie:=c, sPostData:=p)
'End Sub

'Sub 亚马逊测试() '需要使用cookie ,refer
'Dim s As String
'Dim fl As TextStream
'Dim o As Object
'Dim a As Object
'Dim b As Object
'Dim c As Object
's = "sp-cdn=" & Chr(34) & "L5Z9:CN" & Chr(34) & "; x-wl-uid=1gTFvMWicBAS/OVYTHqeEwkOdtHY16tRkY+izrb+kdthSQslKOHPM4XzVh2fF4qNdqL1rdKFeW/k=; lc-main=en_US; skin=noskin; i18n-prefs=USD; session-token=yVguuun/R0A25bIVJ2Cc+bQIRIApM+vsBmq2YfKVS6MRHiSRYrvr8cEuUh/rs3BlR4hMTb1PkCyvcpuDIzXdizhbU6ZJLApqwbUsmZTZF7mPayy4glsRkns0ZwsloPwa3D47h1n28LeFJ8KNatKMFk+xQWYw6uR6McVG6TNJOOWD6p6AQoRudMB9f5SL9q4KQ7wh5/OJMx1FUflmP90PQjYrG7IaH6YFFgB7BgwSIFscrYbgo0FBxYGwVKJ//d5O; csm-hit=tb:s-D5CXSTJN04N9M8DTKDZ9|1589338963314&t:1589338963888&adb:adblk_no; ubid-main=133-8003022-8217246; session-id-time=2082787201l; session-id=142-8143497-7236122"
's = HTTP_GetData("GET", "https://www.amazon.com/s?k=Car+wash&ref=nb_sb_noss_2", "https://www.amazon.com/?currency=USD&language=en_US", acLang:="en-US", sCookie:=s)
'Set fl = Nothing
'WriteHtml s
'Set o = oHtmlDom.getelementsbyclassname("s-result-list s-search-results sg-row")
'For Each a In o
'Set b = a.getelementsbyclassname("a-size-medium a-color-base a-text-normal")
'Next
'Set oHtmlDom = Nothing
'End Sub
'
'Sub 拉钩()
'Dim s As String
'Dim dic As Object
'Dim cj As New cJson_html
's = "JSESSIONID=ABAAABAABEIABCI09BECE7B7FC2349BB3245D35C4779343; SEARCH_ID=7b220089dd4641ef81e29c52fcc3d2ef; user_trace_token=20200513220956-b8c61f8f-81c3-4614-99e8-370a6ce5db00; X_HTTP_TOKEN=42daf4b72327b2816998739851bf5e71415983ed09; WEBTJ-ID=20200513220950-1720e5ec2f02e0-059afcc45a0616-1d154237-2073600-1720e5ec2f52e9"
's = HTTP_GetData("POST", "https://www.lagou.com/jobs/positionAjax.json?xl=%E6%9C%AC%E7%A7%91&px=default&city=%E9%95%BF%E6%B2%99&needAddtionalResult=false", "https://www.lagou.com/jobs/list_%E4%BA%A7%E5%93%81%E7%BB%8F%E7%90%86/p-city_198?px=default&xl=%E6%9C%AC%E7%A7%91", sCookie:=s, sPostData:="first=true&pn=1&kd=%E4%BA%A7%E5%93%81%E7%BB%8F%E7%90%86")
'
'Set dic = cj.Json_oParse(s)
'Debug.Print
'Dim arr() As String
'ReDim arr(15)
'ReDim arr(15, 6)
'Dim i As Integer
'dic.eval "var t=o['content']['positionResult']['result']"
'For i = 0 To 14
'arr(i, 0) = dic.eval("t[" & i & "]['city']")
'arr(i, 1) = dic.eval("t[" & i & "]['companyFullName']")
'arr(i, 2) = dic.eval("t[" & i & "]['district']")
'arr(i, 3) = dic.eval("t[" & i & "]['education']")
'arr(i, 4) = dic.eval("t[" & i & "]['firstType']")
'arr(i, 5) = dic.eval("t[" & i & "]['thirdType']")
'arr(i, 6) = dic.eval("t[" & i & "]['workYear']")
'Next
'
'Set cj = Nothing
'End Sub
Function jb001(ByVal j As Byte) As String()
Dim arr() As String
If j = 1 Then
ReDim arr(1)
ReDim jb001(1)
For i = 0 To 1
arr(i) = CStr(i)
Next
jb001 = arr
End If
End Function

Sub fkkl()
Dim x As Boolean
Dim j As Boolean
x = False
j = False
Debug.Print x * j
End Sub

Sub dfmkf()
Dim cr As New cRegex
cr.oReg_Initial
cr.oReg_Text = "End Function https://cn.bing.com/&"
cr.oReg_Pattern = "&"
Debug.Print cr.ReplaceText("\.")
Set cr = Nothing
End Sub

Sub klls1233()
Dim arr() As String
arr = Null


Sub dkkkkal()
Dim arr() As New cWinHttpRQ
Dim arrt
Dim url As String
Dim c As String

ReDim arr(25) As New cWinHttpRQ
For i = 0 To 25
    If i Mod 2 = 0 Then c = Cookie_Ch(i \ 2)
    With arr(i)
        .Index = i
        .url = "http://data.10jqka.com.cn/funds/ggzjl/field/zdf/order/desc/page/" & CStr(i + 1) & "/ajax/1/free/1/"
        .StartRe c
        Do Until .isOK = True
        If .IsErr = True Then Exit Do
        DoEvents
        Loop
    End With
Next

arrt = arr(0).Result

arrt = ""
GoTo 100
'
'Set b = arr(1).aResult(1)
'a = arr(0).aResult(1)

'Set a = arr(1).aResult
'Do Until arr(4).Sucount = 5
'DoEvents
'Loop
'Set arrt = arr(2).aResult
'x = arr(2).aResult(1)
'GoTo 100
'
'Do Until arr(9).Sucount = 10
'DoEvents
'Loop
100
Erase arr
'Set arrt = arr(0).aResult
End Sub

'Function ChangProxy(ByVal ProxyAddress As String, ByVal isEnable As Boolean)
''--------------------------https://docs.microsoft.com/zh-cn/previous-versions/windows/desktop/regprov/setdwordvalue-method-in-class-stdregprov
'Dim ModRW_Reg As Object
'Set ModRW_Reg = GetObject("winmgmts:\\.\root\default:StdRegProv")
'If isEnable Then
'    ModRW_Reg.SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings", "ProxyServer", ProxyAddress
'
'    ModRW_Reg.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings", "ProxyEnable", 1
'Else
'    ModRW_Reg.DelValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings", "ProxyServer"
'    ModRW_Reg.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings", "ProxyEnable", 0
'End If
'    Call internetsetoption(0, 39, 0, 0)
'    Set ModRW_Reg = Nothing
'End Function
Sub kdfll()
Dim xjk As New cRegex


'Dim st As Stream
'Set st = fso.OpenTextFile("C:\Users\adobe\Desktop\异常文件-计算md5出现错误值.txt", ForReading, False, TristateUseDefault)
'Set objhash = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
'hash = objhash.ComputeHash_1(st)
'st.Close
'objhash.Clear
'Set st = Nothing
'Set objhash = Nothing
'Debug.Print Mid("abc", 1)
'Dim Hash() As Byte
'Hash = StrConv("我是谁", vbFromUnicode)
'k = UBound(Hash)
'For i = 1 To k
'Debug.Print i + i + 2 + (Hash(i) > 15)
'Next
''hash = "我是谁"
'hash = StrConv("我是谁", vbUnicode)

'    For i = 1 To k                                                     'Returns a String representing the hexadecimal value of a number.-hex
'        Result = Result & Right$("0" & Hex(AscB(MidB(hash, i, 1))), 2) 'http://www.officetanaka.net/excel/vba/function/AscB.htm
'        '---------------------------------------------------------------https://docs.microsoft.com/en-us/office/vba/Language/Reference/user-interface-help/hex-function
'        '---------------------------------------------------------------https://www.engram9.info/visual-basic-vba/mid-mid-midb-midb-functions.html
'    Next
'    Debug.Print Result
'Dim strx As String
'Dim i As Byte, k As Byte
'Dim WD As Object
'
'    On Error Resume Next
'    strx = ThisWorkbook.Path
'    strx = strx & "\LB.docm"
'    If Err.Number > 0 Then Err.Clear
'    Set WD = GetObject(, "word.application")
'    If Err.Number > 0 And WD Is Nothing Then
'        Err.Clear
'        Set WD = CreateObject("word.application")
'    End If
'    With WD
'        i = .Documents.Count
'        For k = 1 To i
'            If strx = .Documents(k).fullname Then GoTo 100
'        Next
'        .Documents.Open (filen)
'        .Visible = True
'        .Activate
'    End With
'100
'    Set WD = Nothing
'
'       Dim objShell
'
'        Set objShell = CreateObject("shell.application")
'
'        objShell.ShellExecute "C:\Users\adobe\Desktop\12164431_《梦幻花》【日】东野圭吾.epub", "", "", "open", 1
'
'        Set objShell = Nothing

'        Dim objShell
'
'        Set objShell = GetObject("System.ServiceProcess.ServiceController")
'        objShell.stop "Spooler"
'
'        Set objShell = Nothing
'With CreateObject("shell.application").Namespace("System.ServiceProcess")
'
'.ServiceController.Start ("Spooler")
'
'End With
'Dim strx2 As String
'Dim wd As Object
'
'SetClipboard "HLAstaticx"
'Set wd = CreateObject(strx2)
'wd.Close savechanges:=False
'strx2 = ThisWorkbook.Path & "\Test.docm"
'SetClipboard "HLAstaticx"
'Set wd = GetObject(strx2)
'wd.Application.Run "Test"
'Debug.Print GetClipboard
'Set wd = Nothing
'Dim obj As Object
'Set obj = CreateObject("System.IO.IsolatedStorage.IsolatedStorageFileStream")
'
'
'Set obj = Nothing
End Sub
Sub StopTim0025er() '计时器 /理论上较高精度
'Debug.Print a.MD5Hash("C:\Users\adobe\Desktop\Windows资源管理器.pdf", True, True)
'For i = 1 To 2
Dim o As Object
Dim s As String
Dim fl As TextStream
Set fl = fso.OpenTextFile("C:\Users\adobe\Desktop\json_test.txt", ForReading, False, TristateUseDefault)
s = fl.ReadAll
fl.Close
Set fl = Nothing
Set o = JsonConverter.ParseJson(s)

Debug.Print


For Each itme In o
Debug.Print itme("title")
Next
Set o = Nothing
End Sub
'Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrcommand As String) As Long
'Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
'Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'Private mciID As Integer
'Sub dkkd()
'GoogleSay "hello"
'End Sub
'Function GoogleSay(ByVal sWord As String)
'    Dim objJs    As Object
'    Dim sFile    As String
'    Dim sTmpPath As String
'
'    sTmpPath = Space(255)
'    GetTempPath 255, sTmpPath
'    sTmpPath = Left(sTmpPath, InStr(sTmpPath, Chr(0)) - 1)
'    sFile = sTmpPath & mciID Mod 2 & ".mp3"
'    mciID = mciID + 1
'    Set objJs = CreateObject("MSScriptControl.ScriptControl")
'    objJs.Language = "JavaScript"
'    sWord = objJs.Eval("encodeURI('" & Replace(sWord, "'", "\'") & "');")
'    URLDownloadToFile 0, "http://translate.google.cn/translate_tts?ie=UTF-8&q=" & sWord & "&tl=zh-CN&prev=input", sFile, 0, 0
'    mciExecute "play " & sFile
'End Function
Sub WordPr02255otection(ByVal FilePath As String, ByVal Password As String, Optional ByVal cmCode As Byte) '设置word文件保护
'Dim obj As Object
'Set obj = CreateObject("word.application")
Dim myDoc As Object
Set myDoc = CreateObject("C:\Users\Master\Desktop\test.docx")
strPassword = "123456"
'Set myDoc = Documents.Open(Filename:="C:\My Documents\Earnings.doc")
myDoc.Password = strPassword
myDoc.Close
Set myDoc = Nothing
End Sub
Sub dk1258l()
'    Dim strx As String
'    Dim objexcel As Object
'    Dim wbs As Object
Debug.Print FileHashes("C:\Users\adobe\Desktop\LBCopy.csv")
'    strx = "C:\Users\adobe\Documents\test.xlsm"
'    Set objexcel = CreateObject("excel.application")
'    Set wbs = objexcel.Workbooks
'    objexcel.Visible = True
'    wbs.Open strx
'    Set objexcel = Nothing
'    Set wbs = Nothing
End Sub
Sub dkkd002()

'AddWordPassword "C:\Users\adobe\Desktop\test004.docx", cmcode:=1
'Dim fl As File, flop As Object
'On Error GoTo errhandle
'Set fl = fso.GetFile("C:\Users\adobe\Desktop\test004.docx")
'Set flop = fl.OpenAsTextStream(ForAppending, TristateUseDefault)
'flop.Close
'Set fl = Nothing
'Set flop = Nothing
'Exit Sub
'errhandle:
'Debug.Print Err.Number
'Err.Clear
'flop.Close
'Set fl = Nothing
'Set flop = Nothing
'newdc.ExportAsFixedFormat OutputFileName:=strx, ExportFormat:=17, OpenAfterExport:=True
Dim wd As Object, wdoc As Object
Dim strx As String
Dim Errcount As Byte
Errcount = 0
On Error GoTo ErrHandle
strx = "C:\Users\adobe\Desktop\test0041.pdf"
Set wd = CreateObject("word.application")
enterpassword:
Set wdoc = wd.documents.Open("C:\Users\adobe\Desktop\test004.docx", ReadOnly:=True, Visible:=False)
wdoc.ExportAsFixedFormat OutputFileName:=strx, ExportFormat:=17
wdoc.Close
wd.Quit
Set wd = Nothing
Set wdoc = Nothing
Exit Sub
ErrHandle:
If Err.Number = 5408 And Errcount < 2 Then
Err.Clear
Errcount = Errcount + 1
Resume enterpassword
End If
wd.Quit
Set wd = Nothing
Set wdoc = Nothing
Debug.Print Err.Number
Err.Clear
End Sub
'arra(0) = 239
'arra(1) = 187
'arra(2) = 191
'                mi = 0
'                p = 1
''                strx2 = Replace(strx, " ", "")
''                strlen2 = Len(strx2)
''                If strlen2 = 0 Then Exit Sub
''                ReDim arrstr(1 To strlen2)
''                For i = 1 To strlen2 '判断输入额内容是否为复合内容
''                    strx1 = Mid(strx2, i, 1)
''                    If strx1 Like "[一-龥]" Then '中文 '可单个处理,可集中处理 "公司财务" 可拆解成单个单词 公/司/财/务, 所以可以不处理空格
''                        k = k + 1
''                        arrstr(p) = strx1
''                        p = p + 1
''                    ElseIf strx1 Like "[a-zA-Z]" Then '英文字母,含大小写 '集中处理,单个拆解单词的含义大幅下降
''                        j = j + 1
''                    ElseIf strx1 Like "[0-9]" Then '数字 '集中处理,单个处理没有单词拆解的影响大
''                        xi = xi + 1
''                    End If
''                Next
'''                If k = 0 And xi = 0 And strlen2 < 3 Then Exit Sub '假如是纯英文字母,长度小于即不作出反应
'''                If j <> strlen2 And k <> strlen2 And xi <> strlen2 Then
''''                    If SafeArrayGetDim(AbtainKeyWord(strx)) = 0 Then Exit Sub '没有获取到有效的值
'''                    keycount = keycount - 1
'''                    If keycount >= 0 Then Exit Sub
'''                    ReDim arr(keycount)
'''                    arr = ObtainKeyWord(strx)
'''                ElseIf k = 0 And j = 0 And xi = strlen2 And strlen2 >= 2 Then '数字
'''                    strx = strx2
'''                End If
''                '-----------------------------------前期关键词分析
'                For j = 1 To 2
'                    For k = 1 To 4
'                    If InStr(1, arrax(k, 1 * j) & "/" & arrax(k, 2 * j), strx, vbTextCompare) > 0 Then
'                        dic(arrax(k, 1)) = arrax(k, 2)
'                        dica(arrax(k, 1)) = arrax(k, 3)
'                        dicb(arrax(k, 1)) = arrax(k, 4)
'                        dicc(arrax(k, 1)) = arrbx(k, 1) '搜索的方式还可以有很大的调整空间'制作一张近似词的表,当搜索英文的某些词语可以同步检索'将拼写错误的词替换掉进行搜索
'                        mi = mi + 1
'                        If mi > 50 Then GoTo 100 '限制搜索结果的数量
'                    Else
'                        For i = 1 To strlen2 '判断输入额内容是否为复合内容
'                            strx1 = Mid(strx2, i, 1)
'                            If strx1 Like "[一-龥]" Then '中文 '可单个处理,可集中处理 "公司财务" 可拆解成单个单词 公/司/财/务, 所以可以不处理空格
'                                k = k + 1
'                                arrstr(p) = strx1
'                                p = p + 1
'                            ElseIf strx1 Like "[a-zA-Z]" Then '英文字母,含大小写 '集中处理,单个拆解单词的含义大幅下降
'                                j = j + 1
'                            ElseIf strx1 Like "[0-9]" Then '数字 '集中处理,单个处理没有单词拆解的影响大
'                                xi = xi + 1
'                            End If
'                        Next
'
'
'                    Next
'                Next
'                '----------------------------------------------------------------
'                ReDim arrstr(strlen)
'                For j = 1 To 4
'                    For k = 1 To blow '注意由于筛选的目标之间有相同的字符,将导致出现多行结果的bug,这里使用字典的方法来解决
'                        If InStr(1, arrax(k, j), strx, vbTextCompare) > 0 Then
'                            dic(arrax(k, 1)) = arrax(k, 2)
'                            dica(arrax(k, 1)) = arrax(k, 3)
'                            dicb(arrax(k, 1)) = arrax(k, 4)
'                            dicc(arrax(k, 1)) = arrbx(k, 1) '搜索的方式还可以有很大的调整空间'制作一张近似词的表,当搜索英文的某些词语可以同步检索'将拼写错误的词替换掉进行搜索
'                            mi = mi + 1
'                            If mi > 50 Then GoTo 100 '限制搜索结果的数量
'                        Else
'                            If InStr(strx, Chr(32)) > 0 Then '输入的内容存在空格 -更模糊的搜索
'                                p = 1
'                                For m = 1 To strlen
'                                    strx1 = Mid(strx, m, 1)
'                                    If strx1 Like "[一-龥]" Then '只针对中文字符
'                                        arrstr(p) = strx1
'                                        p = p + 1
'                                    End If
'                                Next
'                                If p < 2 Then Exit For '内容太少,不再进行检索
'                                xi = 0
'                                For t = 1 To p
'                                    If InStr(1, arrax(k, j), arrstr(t), vbTextCompare) > 0 Then xi = xi + 1
'                                Next
'                                If xi > 2 Then
'                                    dic(arrax(k, 1)) = arrax(k, 2)
'                                    dica(arrax(k, 1)) = arrax(k, 3)
'                                    dicb(arrax(k, 1)) = arrax(k, 4)
'                                    dicc(arrax(k, 1)) = arrbx(k, 1) '搜索的方式还可以有很大的调整空间
'                                    mi = mi + 1
'                                    If mi > 50 Then GoTo 100
'                                End If
'                            Else         '模糊搜索,不包含空格的
'                                p = 0
'                                '-----------------------如果strx="a", "abc",经过转换后,长度将为0,1(注意)
'                                If Len(strx) \ 2 = Len(StrConv(strx, vbFromUnicode)) Then
'                                '--------不包含中文字符,注意这里的不能用工作表中的len/lenb来区分中英文字符的差异,必须进行转换后才能进行比较
'                                    If InStr(1, arrax(k, j), strx, vbTextCompare) > 0 Then
'                                        dic(arrax(k, 1)) = arrax(k, 2)
'                                        dica(arrax(k, 1)) = arrax(k, 3)
'                                        dicb(arrax(k, 1)) = arrax(k, 4)
'                                        dicc(arrax(k, 1)) = arrbx(k, 1)
'                                        Debug.Print "OK"
'                                    End If
'                                Else
'                                For m = 1 To strlen
'                                    strx1 = Mid(strx, m, 1)
'                                    If strx1 Like "[一-龥]" Then '只针对中文字符
'                                        p = p + 1
'                                        arrstr(p) = strx1
'                                    End If
'                                Next
'                                If p < 2 Then Exit For '内容太少,不再进行检索
'                                xi = 0
'                                For t = 1 To p
'                                    If InStr(1, arrax(k, j), arrstr(t), vbTextCompare) > 0 Then xi = xi + 1
'                                Next
'                                If xi > 2 Then
'                                    dic(arrax(k, 1)) = arrax(k, 2)
'                                    dica(arrax(k, 1)) = arrax(k, 3)
'                                    dicb(arrax(k, 1)) = arrax(k, 4)
'                                    dicc(arrax(k, 1)) = arrbx(k, 1) '搜索的方式还可以有很大的调整空间
'                                    mi = mi + 1
'                                    If mi > 50 Then GoTo 100
'                                End If
'                                End If
'                            End If
'                        End If
'                    Next
'                Next
'    '            n = UBound(ltma)

Sub testActiveWindowSize()

'If spyx <> docmx Then  '利用模块级变量保存搜索区域的值在内存中,减少访问表格的需要,只有当表格的数据发生变化才重新获取值,加快访问的速度
'    spyx = blow '初始赋值/变化在进行赋值
'    If SafeArrayGetDim(arrax) <> 0 Then
'    Erase arrax
'    Erase arrbx                        '当数据发生变化的时候,抹掉旧的数据重新获取新的
'    End If
'    With ThisWorkbook.Sheets("书库")
'        arrax = .Range("b6:e" & blow).Value
'        arrbx = .Range("n6:n" & blow).Value '打开次数
'    End With
'End If
'ReDim arrx(1)
'arrx(1) = 1
'Erase arrx
'MsgBox SafeArrayGetDim(arrx)
' Debug.Print Len(StrConv("aaa", vbFromUnicode)), Len("我是谁"), Len(StrConv("我aa", vbFromUnicode)), Len(StrConv("我aaa", vbFromUnicode)), Len("我aaa")
'Dim arr() As String
''If IsArray(arr) = True Then MsgBox 1
''ReDim arr(1)
'Debug.Print SafeArrayGetDim(aaa)
''Dim myreg As Object, match As Object, matches As Object
'Dim arr() As String, k As Byte, xtext As String, i As Byte
'
'xtext = "whot我时100 am"
'Set myreg = CreateObject("VBScript.RegExp")
'With myreg
'    .Pattern = "[a-zA-Z0-9]{2,}" '获取豆瓣评分
'    .Global = True
'    .IgnoreCase = True '不区分大小写
'    Set matches = .execute(xtext)
'    k = matches.Count - 1
'    ReDim arr(k)
'    For Each match In matches
'        arr(i) = match.Value
'        i = i + 1
'    Next
'End With
'Set myreg = Nothing
'Set match = Nothing
'Set matches = Nothing

'[\>]+(.+?)[\(]豆瓣[\)]
'用户评级:+(.+?)(\<)
'[https]+://+book.douban.com/+[a-z]*\/+[0-9]*\/

            '-------------------------------------因为搜索的内容会有多个干扰结果,所以返回单一结果,准确度不一定高
'url = "https://book.douban.com/subject/3259440/"
'With CreateObject("MSXML2.XMLHTTP")
'    .Open "GET", url, False
'    .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'    .send
'    Do While .readyState <> 4 '等待数据的返回
'        DoEvents
'    Loop
'    strx = .responsetext
'Debug.Print strx
'                End
'End With



' Debug.Print "当前窗口可用区域的高度为:" & ActiveWindow.UsableHeight
'
' Debug.Print "当前窗口的高度为:" & ActiveWindow.Height
'
' Debug.Print "当前窗口可用区域的宽度为:" & ActiveWindow.UsableWidth
'
' Debug.Print "当前窗口的宽度为:" & ActiveWindow.Width
'ThisWorkbook.Application.ScreenUpdating = True
'Dim dicx As New Dictionary
'Dim m As Byte
'm = 1
'dicx(m) = ""
'Debug.Print dicx.Keys(0)
'Set dicx = Nothing
''Dim fd As Folder
'
'Set fd = fso.GetFolder("C:\Users\Administrator\Documents\自定义 Office 模板")
'
'Set fd = Nothing

 End Sub






'数据的类型
'byte 符号: - , 占用: 1 数据范围: 0-255
'long 符号: & , 占用: 4,数据范围 -2147483648 ~ 2147483647
'integer 符号: %, 占用: 2 ,数据范围: -32768 ~ 32767
'Private Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)
''Private Declare Function hash Lib "ntdll.dll" Alias "RtlComputeCrc32" (ByVal start As Long, ByVal data As Long, ByVal size As Long) As Long
'Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'ElseIf fso.GetDriveName(strx) = Environ("SYSTEMDRIVE") Then '禁止设置在系统盘一级文件夹设置/只允许将文件夹放在document内
'                c = UBound(Split(strx, "\"))
'                If c = 1 Then
'                    .Label57.Caption = "不允许在系统盘设置"
'                    .TextBox11.text = ""
'                    .TextBox11.SetFocus
'                ElseIf c = 2 Then
'                    If strx <> Environ("UserProfile") Then
'                    .Label57.Caption = "不允许在非用户文件夹"
'                    .TextBox11.text = ""
'                    .TextBox11.SetFocus
'                    End If
'                ElseIf c > 2 Then
'                    strx2 = Environ("UserProfile") & "\Documents"
'                    If InStr(strx, strx2) = 0 Then
'                        .Label57.Caption = "只允许设置在documents文件夹下"
'                        .TextBox11.text = ""
'                        .TextBox11.SetFocus
'                    End If
'i = FileStatus(strx3)
'            Select Case i
'                Case 1: strx5 = "目录不存在"
'                Case 3: strx5 = "文件不存在"
'                Case 6: strx5 = "文件处于打开的状态"
'                Case 7: strx5 = "异常"
'            End Select
'            If i = 1 Or i = 3 Then
'                .Label55.Visible = ture
'                DeleFileOverx strx3
'                .TextBox1.Text = ""
'            End If
'            If i <> 0 Then
'                .Label57.Caption = strx5
'                Set Rng = Nothing
'                Exit Sub
'            End If

'Private Function CRC32(Data() As Byte) As Long
'    Dim Asm(5) As Long, Init As Long
'
'    Asm(0) = &H5B5A5958
'    Asm(1) = &HC033505E
'    Asm(2) = &H3018A36
'    Asm(3) = &H41CED1F0
'    Asm(4) = &HF47ECA3B
'    Asm(5) = &HC3338936
'
'    On Error GoTo Err
'    Init = UBound(Data) - LBound(Data) + 1
' CallWindowProc VarPtr(Asm(0)), VarPtr(Data(LBound(Data))), VarPtr(Data(UBound(Data))), VarPtr(CRC32), Init
'Err:
'End Function
'Dim strdsk As String, strdlw As String, strdcm As String, struserfile As String '文档,下载,桌面,用户文件夹
'If tracenum = 1 Then '该文件为系统盘文件,只允许添加三个位置的文件desktop,downloads,documents '
'                        struserfile = Environ("UserProfile") '用户文件夹
'                        strdsk = struserfile & "\Desktop"
'                        strdlw = struserfile & "Downloads"
'                        strdcm = struserfile & "\Documents"
'                        If c = 3 Then
'                            If strp <> strdsk And strp <> strdcm And strp <> strdlw Then GoTo 1001 '限制三个文件夹
'                            ElseIf c > 3 Then
'                                If InStr(fd.Path, strdsk) = 0 And InStr(fd.Path, strdlw) = 0 And InStr(fd.Path, strdcm) = 0 Then GoTo 1001
'                            ElseIf c < 3 Then
'                                If strp <> Environ("UserProfile") Then GoTo 1001 '限制只允许添加用户文件夹
'                        End If
'                    End If

'Function ZipExtract(ByVal filepath As String, ByVal folderpath As String, Optional ByVal cmcode As Byte) '解压文件
'    Dim exepath As String, strx As String, strx1 As String, strx2 As String, strx3 As String
'    Dim wsh As Object
'
'    exepath = "C:\Program Files\7-Zip\7z.exe" '如果安装到其他位置
'    exepath = """" & exepath & """"
'    Set wsh = CreateObject("WScript.Shell")
'
'    folderpath = """" & folderpath & """"
'    filepath = """" & filepath & """"
'    strx1 = """" & """" & filepath & """" & """"
'    strx2 = """" & """" & folderpath & """" & """"
''    strx = " e " & filepath & " -o" & folderpath
'
'    strx = " e " & strx1 & " -o" & strx2
'
'    If cmcode = 1 Then
'        strx2 = " && del /s " & filepath '执行完成后自动删除掉源文件 'del /s cmd 删除文件命令
'        strx3 = exepath & strx & strx2
'    Else
'        strx3 = exepath & strx
'    End If
'    wsh.Run ("cmd /c " & strx3)
'    Set wsh = Nothing
'End Function
'Function POP(ByVal filepath As String, ByVal folderpath As String)
'Dim exe As String, strx As String, strx1 As String, strx2 As String, strx3 As String
'Dim wsh As Object
'
'Set wsh = CreateObject("WScript.Shell")
'exe = "C:\Program Files\7-Zip\7z.exe "
'
'strx = filepath
'strx1 = folderpath
'With Me
'strx1 = .TextBox1.Text
'strx3 = .TextBox13.Text                        '根据信息的不同区域自动获取搜索关键词-需要修改
'strlen1 = Len(Trim(strx1))
'strlen3 = Len(Trim(strx3))
'If strlen1 = 0 And strlen3 = 0 Then Exit Sub '1
'If strlen1 = 0 And strlen3 > 0 Then             '2
''    keyworda = Left$(strx3, Len(strx3) - Len(Split(strx3, ".")(UBound(Split(strx3, ".")))) - 1)
'    keyworda = strx3
'ElseIf strlen1 > 0 And strlen3 = 0 Then        '3
'       If strx1 Like "HLA*&*" Then                     '3.1
'          keyworda = Mid$(strx1, Len(Split(strx1, "&")(0)) + 3, Len(strx1) - Len(Split(strx1, "&")(0)) - 2 - Len(Split(strx1, ".")(UBound(Split(strx1, ".")))) - 1) '在这里可以注意到split实际上也是一个数组（可以轻松的分离首尾的分割组）
'       Else
'          keyworda = strx1                               '3.2
'       End If
'ElseIf strlen1 > 0 And strlen3 > 0 Then        '4
'       If strx1 Like "HLA*&*" Then
'          keyworda = Mid$(strx1, Len(Split(strx1, "&")(0)) + 3, Len(strx1) - Len(Split(strx1, "&")(0)) - 2 - Len(Split(strx1, ".")(UBound(Split(strx1, ".")))) - 1)
'       Else
'          keyworda = strx1
'       End If
'End If

'strx = """" & strx & """"
'strx1 = """" & strx1 & """"
'exe = """" & exe & """"
'
'wsh.Run (exe & " e " & strx & " -o" & strx1), vbHide
'Set wsh = Nothing
'
'End Function
'Sub kfk()
'POP Cells(6, 1), Cells(7, 1)
'End Sub
'Private Sub TextBox5_Change()
'    Dim i As Long
'    With Me
'    If .TextBox5.Value = "" Then
'        i = 0 ' -0.5
'    Else
'        i = .TextBox5.Value / 1.333 '- 0.5
'    End If
'    With .Label8
'        .Left = Image1.Left + i
'        .Top = Image1.Top + i
'        .Height = Image1.Height - i * 2
'        .Width = Image1.Width - i * 2
'    End With
'    End With
'End Sub
'Private Sub TextBox5_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    If TextBox5.Text = "" Then
'        TextBox5.Text = 0
'    ElseIf Mid(TextBox5.Text, 1, 1) = "0" Then
'        TextBox5.Text = Mid(TextBox5.Text, 2, 3)
'    End If
'    If TextBox5.Text > Image1.Height * 1.33 / 2 Then
'        TextBox5.Text = Int(Image1.Height * 1.33 / 2)
'        TextBox5.SelStart = 0
'        TextBox5.SelLength = Len(TextBox5.Text)
'    End If
'End Sub
'With Me
'strx1 = .TextBox1.Text
'strx3 = .TextBox13.Text                        '根据信息的不同区域自动获取搜索关键词-需要修改
'strlen1 = Len(Trim(strx1))
'strlen3 = Len(Trim(strx3))
'If strlen1 = 0 And strlen3 = 0 Then Exit Sub '1
'If strlen1 = 0 And strlen3 > 0 Then             '2
''    keyworda = Left$(strx3, Len(strx3) - Len(Split(strx3, ".")(UBound(Split(strx3, ".")))) - 1)
'    keyworda = strx3
'ElseIf strlen1 > 0 And strlen3 = 0 Then        '3
'       If strx1 Like "HLA*&*" Then                     '3.1
'          keyworda = Mid$(strx1, Len(Split(strx1, "&")(0)) + 3, Len(strx1) - Len(Split(strx1, "&")(0)) - 2 - Len(Split(strx1, ".")(UBound(Split(strx1, ".")))) - 1) '在这里可以注意到split实际上也是一个数组（可以轻松的分离首尾的分割组）
'       Else
'          keyworda = strx1                               '3.2
'       End If
'ElseIf strlen1 > 0 And strlen3 > 0 Then        '4
'       If strx1 Like "HLA*&*" Then
'          keyworda = Mid$(strx1, Len(Split(strx1, "&")(0)) + 3, Len(strx1) - Len(Split(strx1, "&")(0)) - 2 - Len(Split(strx1, ".")(UBound(Split(strx1, ".")))) - 1)
'       Else
'          keyworda = strx1
'       End If
'End If

'End With
'https://docs.microsoft.com/en-us/windows/win32/api/imagehlp/nf-imagehlp-imageenumeratecertificates


'Private Sub Command1_Click()
'    If SearchFile("WinRAR.exe") <> "没有找到匹配项！" Then '查找WinRAR.exe的路径
'        Shell SearchFile("WinRAR.exe") & "   a   c:\压缩后的文件.rar   c:\被压缩的文件或文件夹", vbHide
'    Else
'        MsgBox "没有找到RAR压缩软件", vbOKOnly, "提示"
'    End If
'End Sub
    
'        With ThisWorkbook
'            blow = .Sheets("书库").[d65536].End(xlUp).Row
'            If blow <> docmx Then
'                docmx = blow
'                Call CwUpdate '更新窗口的数据
'                Call Choicex '更新筛选区的数据
'            End If
'            With .Sheets("主界面") '获取初始值     '更新窗体主界面的4个位置的数据
'                recentfilex = CStr(.Range("w26").Value)
'                prfilex = .Range("i26").Value
'                addfilecx = .[e65536].End(xlUp).Row
'
'                If recentfilex <> Recentfile Then '在表格打开文件后的处理
'                    Recentfile = recentfilex
'                    Call RecentUpdate
'                    With Me
'                        strx1 = .Label1.Caption
'                        strx2 = .Label32.Caption
'                        If Len(strx1) > 0 And strx1 = .ListBox1.List(0, 0) Then '更新编辑搜索中的数据
'                            .Label31.Caption = .ListBox1.List(0, 2)
'                            If Len(strx2) = 0 Then
'                                .Label32.Caption = 1
'                            Else
'                                .Label32.Caption = Int(strx2) + 1
'                            End If
'                        End If
'                    End With
'                End If
'
'                If prfilex <> prfile Then
'                    prfile = prfilex
'                    Call PrReadList
'                End If
'
'                If addfilecx <> addfilec Then
'                    addfilec = addfilecx
'                    Call AddFileListx
'                End If
'            End With
'        End With
'    End If
Sub wbwj()
Dim a() As String, n&, i&, m&
Dim MyPath$, MyName$
MyPath = ThisWorkbook.Path & "\"
MyName = Dir(MyPath & "*.txt")
Sheet1.Activate
[a2:b1000].ClearContents: m = 1
Do While MyName <> ""
    n = 0
    Open MyPath & MyName For Input As #1
    a = Split(StrConv(InputB(LOF(1), 1), vbUnicode), vbCrLf)
    For i = 0 To UBound(a)
        n = n + UBound(Split(a(i))) + 1
    Next
    m = m + 1
    Cells(m, 1) = MyName
    Cells(m, 2) = n
    Close #1
    MyName = Dir
Loop
End Sub
Sub kdkk()
Dim strx As String
strx = Range("p12").Value
strx = Trim(Replace(strx, Chr(10), ""))
strx = Trim(Right(strx, Len(strx) - 3))
Debug.Print strx
'Dim arr() As String, arrx() As String
''ReDim arr(1 To 3)
''arr(1) = 1
''arr(2) = 2
''ReDim arrx(1 To 2)
''arrx = arr
'Dim dic As New Dictionary
'ReDim arr(1 To 2)
'arr(1) = 1
'arr(2) = 2
'
'With dic
'For i = 1 To 3
'
'.key(arr(i)) = ""
'
'Next
'Dim arr() As Integer
'arr = test(3)

'Shell "powershell regedit", vbNormalFocus
''End With
End Sub

Sub Test2020511()
Dim s As String
Dim i As Integer
Dim Matches As Object
Dim match As Object
Dim sPatern
Dim arr() As String
Dim k As Integer, j As Integer
Dim cr As Object
Set cr = CreateObject("VBScript.RegExp")
sPatern = Array("[\d|\u4e00-\u9fa5]{4,}[\d]{2,}[^\x00-\xff]{6,}", "20[0-2]{1}[\d]{1}-([0-1]{1})?[\d]{1}-([0-3]{1})?[\d]{1}", "showProject\([\d]{1,}\)")
s = HTTP_GetData("GET", "http://www.sdhrzb.cn/messagesContentAll?pageNo=2&pageSize=10&status=0&pname=&pno=&pstate=", "http://www.sdhrzb.cn/loginTwoAll?status=0")
s = Replace$(s, " ", "", 1, , vbBinaryCompare)
ReDim arr(10, 2)
For j = 0 To 2
    With cr
        .Pattern = sPatern(j)
        .Global = True
        .IgnoreCase = True
        Set Matches = .Execute(s)
    End With
    i = Matches.Count - 1
    If i >= 0 Then
        i = 0
        For Each match In Matches
            arr(i, j) = match.Value: i = i + 1
        Next
        Set Matches = Nothing
    End If
Next
Set cr = Nothing
End Sub
'Function Test(ByVal k As Byte) As Integer()
'ReDim Test(1 To k)
'Dim arr() As Integer
'ReDim arr(1 To k)
'
'For i = 1 To k
'
'arr(i) = i
'Next
'
'Test = arr
'
'End Function
'Dim arr(1 To 3) As Integer
'Dim arrx() As Integer
'For i = 1 To 3
'arr(i) = i
'Next
'ReDim arrx(1 To 3)
'arrx = arr
'Dim arr() As Byte
'ReDim arr(2)
'For i = 1 To 3
'arr(i) = i
'Next
'ReDim Preserve arrx(3 To 5)
'arrx = arr
'Dim connx As New ADODB.Connection
'Dim rcs As ADODB.Recordset
'Set rcs = New ADODB.Recordset
'
''rcs.Open "C:\Users\adobe\Desktop\异常文件-计算md5出现错误值.txt",
'
'connx.Open "C:\Users\adobe\Desktop\异常文件-计算md5出现错误值.txt"
'
'    rcs.Open sql, conn, adOpenKeyset, adLockOptimistic
    
    
    
''Dim k As Double
'key = "210381195305143915"
'
'
''key = "210381197005123911"
'
'key = StrConv(key, vbFromUnicode)
'i = StrPtr(key)
'p = hash(0, StrPtr(key), LenB(key))
''j = hash(0, StrPtr(key), LenB(key)) And &HFFFFFFFF
'''2147483647
'''-1
''k = hash(0, StrPtr(key), LenB(key)) And &H7FFFFFFF
'
'Debug.Print Hex$(p)


'Set rcs = Nothing


'Sub dkkfj()
'Dim i As Byte, j As Byte, k As Byte, m As Integer
'Dim strx As String, t As Single
'Dim strx2 As String
'Dim wsh As Object
'Dim dic As New Dictionary
'Dim wExec As Object
'
't = Timer
''For i = 0 To 9
''For k = 0 To 9
''For j = 0 To 9
'Set wsh = CreateObject("WScript.Shell")
''strx3 = "123"
'strx = "C:\Program Files\7-Zip\7z.exe"
'strx = """" & strx & """"
'strx2 = "D:\p\1202601175.7z"
'Application.ScreenUpdating = False
'For m = 0 To 1000
'100
'i = Int((9 - 0 + 1) * Rnd + 0)
'j = Int((9 - 0 + 1) * Rnd + 0)
'k = Int((9 - 0 + 1) * Rnd + 0)
''If fso.FileExists("C:\Users\***\Documents\1202601175.PDF") = True Then Debug.Print Timer - t: Exit Sub
'strx3 = CStr(i) & CStr(k) & CStr(j)
'If dic.Exists(strx3) = True Then GoTo 100
'Set wExec = wsh.Exec("cmd /c " & strx & " e " & strx2 & " -p" & strx3)
'result = wExec.StdOut.ReadAll
''dic(strx3) = ""
''Debug.Print strx3
'If InStr(result, "Errors") = 0 Then Debug.Print Timer - t: Exit Sub
''Call dkjf(strx)
''Next
''Next
''Next
'Next
'Application.ScreenUpdating = True
'Set wsh = Nothing
'Set wExec = Nothing
'End Sub
'
'Sub dkjf(ByVal strx3 As String)
'Dim strx As String
'Dim strx2 As String
'Dim wsh As Object
'
'Set wsh = CreateObject("WScript.Shell")
'strx = "C:\Program Files\7-Zip\7z.exe"
'strx = """" & strx & """"
'strx2 = "D:\p\1202601175.7z"
'wsh.Run ("cmd /c " & strx & " e " & strx2 & " -p" & strx3), vbHide
'Set wsh = Nothing
'End Sub
'
'Sub dlk()
'Dim strx As String
'strx = "测试"
'filepath = ThisWorkbook.Path & "\test.xlsx"
'
'conn.Open "Provider = Microsoft.ACE.OLEDB.12.0;Data Source=" & filepath & ";extended properties=""excel 12.0;HDR=YES""" '打开数据存储文件
'sql = "Insert into [测试$] (编号,文件,原因) Values ('" & unicode & "', '" & filen & "', '" & mfilen & "')"
'
'conn.Close
'Set conn = Nothing
'
'End Sub


'If k > x Then k = x
'If i Mod 2 = 0 Then
'strx1 = GetMD5Hash_String(CStr(i))
'Else
'strx1 = LCase(GetMD5Hash_String(CStr(i)))
'End If
'strlen = Len(strx1) - k
'j = Int((strlen - 1 + 1) * Rnd + 1)
'strx = Mid(strx1, j, k)
'password = password & strx
'Next
'
'm = Len(password)
'If m > numx Then
'For q = 1 To numx
'n = Int((m - 1 + 1) * Rnd + 1)
'n = n + 1
'password = password & Chr(97 + n)
'strx3 = strx3 & Mid(password, n, 1)
'Next
'password = strx3
' If filext Like "EPUB" Or filext Like "PDF" Or filext Like "MOBI" Then
'            k = 1
'        ElseIf filext Like "DO*" Or filex Like "XL*" Or filext Like "PP*" Or filext Like "TX*" Or filext Like "AC*" Then
Sub kkx2019()
'Dim drx As Drive
'For i = 3 To 32
'Cells(5, i) = Left(Cells(5, i), Len(Cells(5, i)) - Len(i - 2))
'Next
'Cells(5, "ae") = "文件名异常字符"
'Cells(5, "af") = "文件件位置异常字符"
'For i = 0 To 31
'Cells(3, i + 2) = i
'Next
'Application.DisplayFormulaBar = True
'For Each drx In fso.Drives
''Debug.Print Environ("SYSTEMDRIVE")
'Debug.Print drx.Path
'Shell "notepad.exe " & "C:\Users\Lian\Desktop\contents.txt", vbNormalFocus '打开文件
 'Shell "cmd /c tree D:\L-temp /f >C:\Users\Lian\Desktop\contents.txt", vbHide
Dim strfolder As String
Dim strx As String, strx1 As String
Dim wsh As Object

With Application.FileDialog(msoFileDialogFolderPicker) '文件夹选择窗口(文件夹以此只能选一个)
    .Show
    If .SelectedItems.Count = 0 Then Exit Sub '未选择文件夹则退出sub
    strfolder = .SelectedItems(1)
    strx = Split(strfolder, "\")(UBound(Split(strfolder, "\")))
End With

Set wsh = CreateObject("WScript.Shell")
strx1 = Environ("UserProfile") & "\Desktop\menus.txt"
wsh.Run "cmd /c tree " & strfolder & " /f >" & strx, 0
Sleep 50
Shell "notepad.exe " & strx1, 3

Set wsh = Nothing
'Next
End Sub

Sub kdlop()
Shell "notepad.exe ", vbNormalFocus
End Sub
'Private Sub CommandButton22_Click() '全部文件夹数据更新-按钮已被删除
'Dim arrall()
'If Me.ListBox3.ListCount = 0 Then Exit Sub
'ReDim arral(0 To Me.ListBox3.ListCount - 1)
'For i = 0 To Me.ListBox3.ListCount - 1
'arral(i) = ListBox3.List(i, 0)
'Next
'Call fleshdata(arral)
'
''End Sub

'                    If Not rngad Is Nothing And filedyn = True Then
'                        rngad.Offset(0, 2) = Now '更新所在文件夹的修改时间
'                        filec = rngad.Offset(0, 4).Value '更新文件夹数量
'                        If filec > 1 Then
'                            filec = filec - 1
'                            rngad.Offset(0, 4) = filec
'                        Else
'                            rngad.Offset(0, 4) = 0
'                        End If
'                    End If
'
'Call SearchFile(keyworda)
'
'If rng Is Nothing Then
'Set rng = Nothing  'rng为工程级变量,在执行完毕重置,释放内存
'Call Warning(3)
'Exit Sub
'End If
'
'Me.Label55.Visible = False
'Me.Label23.Caption = rng.Offset(0, 1) '文件名
'Me.Label24.Caption = rng.Offset(0, 2) '文件类型
'Me.Label25.Caption = rng.Offset(0, 3) '文件路径
'Me.Label26.Caption = rng.Offset(0, 4) '文件位置
'
'Me.Label27.Caption = rng.Offset(0, 6) '文件大小
'Me.Label28.Caption = rng.Offset(0, 7) '创建时间
'Me.Label29.Caption = rng.Offset(0, 0) '统一编码
'Me.Label30.Caption = rng.Offset(0, 9) '文件类别
'
'Me.Label31.Caption = rng.Offset(0, 10) '最近打开时间
'Me.Label32.Caption = rng.Offset(0, 11) '打开次数
'Me.Label33.Caption = rng.Offset(0, 8) '标识编码
'
'Me.TextBox5.text = rng.Offset(0, 18) '标签1
'Me.TextBox6.text = rng.Offset(0, 19) '标签2
'
'Me.TextBox4.text = rng.Offset(0, 13) '作者
'
'Me.ComboBox3.text = rng.Offset(0, 14) 'pdf清晰度
'Me.ComboBox4.text = rng.Offset(0, 15) '文本质量
'Me.ComboBox5.text = rng.Offset(0, 16) '内容评分
'
'Me.ComboBox2.text = rng.Offset(0, 17) '推荐评分
'
'Me.Label69.Caption = rng.Offset(0, 23) '豆瓣评分
'
'Me.Label71.Caption = rng.Offset(0, 20) 'MD5
'
'If rng.Offset(0, 12) = "" Then
'Call Aton
'Else
'Me.TextBox3.text = rng.Offset(0, 12) '主文件名
'End If
'
'
'End With
'
'Call Text2a '获取文件的摘要信息
'Call Disabledit
'End If
'If .ListBox1.ListCount = 0 Then
'.ListBox1.AddItem
'GoTo 100
'
'ElseIf .ListBox1.ListCount = 7 Then          '当列表已经存在7个的时候,不断重写第七行的数据
'For k = 6 To 1 Step -1
'.ListBox1.List(k, 0) = .ListBox1.List(k - 1, 0)
'.ListBox1.List(k, 1) = .ListBox1.List(k - 1, 1)
'.ListBox1.List(k, 2) = .ListBox1.List(k - 1, 2)
'Next
'
'GoTo 100
'
'ElseIf .ListBox1.ListCount > 0 And .ListBox1.ListCount < 7 Then '当数据小于7的时候,数据向下移动
'.ListBox1.AddItem
'n = .ListBox1.ListCount
'For k = n To 1 Step -1
'.ListBox1.List(k, 0) = .ListBox1.List(k - 1, 0)
'.ListBox1.List(k, 1) = .ListBox1.List(k - 1, 1)
'.ListBox1.List(k, 2) = .ListBox1.List(k - 1, 2)
'Next
'
'100                                                                 '当列表为空,直接在第一行写入数据
'.ListBox1.List(0, 0) = .Label29.Caption
'.ListBox1.List(0, 1) = .Label23.Caption
'.ListBox1.List(0, 2) = Now
'End If
'
'If .ListView1.ListItems.Count <> 0 Then
'
'Call listvf(.Label29.Caption)
'
'If itemf Is Nothing Then
'GoTo 1001
'Else
'    If .ListView1.ListItems(itemf.Index).SubItems(4) = "" Then .ListView1.ListItems(itemf.Index).SubItems(4) = 0 '空值和"0"还是有区别的
'    .ListView1.ListItems(itemf.Index).SubItems(4) = .ListView1.ListItems(itemf.Index).SubItems(4) + 1 '打开次数+1
'    .Label32.Caption = .ListView1.ListItems(itemf.Index).SubItems(4)
'End If
'
'End If

'For i = 0 To .ListBox2.ListCount - 1
'If .Label29.Caption = .ListBox2.Column(0, i) Then
'Call Warning(4)
'Exit Sub
'End If
'Next
'If .ListBox2.ListCount = 0 Then
'
'GoTo 100
'
'ElseIf .ListBox2.ListCount = 7 Then                               '当列表已经存在7个的时候,不断重写第七行的数据
'
'For k = 6 To 1 Step -1
'.ListBox2.List(k, 0) = .ListBox2.List(k - 1, 0)
'.ListBox2.List(k, 1) = .ListBox2.List(k - 1, 1)
'.ListBox2.List(k, 2) = .ListBox2.List(k - 1, 2)
'Next
'
'GoTo 100
'
'ElseIf .ListBox2.ListCount > 0 And .ListBox2.ListCount < 7 Then '当数据小于7的时候,数据向下移动
'
'n = .ListBox2.ListCount
'.ListBox2.AddItem
'For k = n To 1 Step -1
'.ListBox2.List(k, 0) = .ListBox2.List(k - 1, 0)
'.ListBox2.List(k, 1) = .ListBox2.List(k - 1, 1)
'.ListBox2.List(k, 2) = .ListBox2.List(k - 1, 2)
'Next
'
'100                                                              '当列表为空,直接在第一行写入数据(注意这里的写法, 被纳入了elseif中去了)
'.ListBox2.List(0, 0) = .Label29.Caption
'.ListBox2.List(0, 1) = .Label23.Caption
'.ListBox2.List(0, 2) = Now
'
'End If













'       .Range("k" & .[c65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrmd5)  '文件md5
'        .Range("ab" & .[c65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrcode) '异常字符标记
'        .Range("ac" & .[c65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrcm)   '备注
'        .Range("ae" & .[c65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrfnansi)
'        .Range("ab" & .[c65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrfpansi)
'        .Range("c" & .[c65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrbase) '文件名
'        .Range("d" & .[d65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrextension) '文件扩展名
'        .Range("e" & .[e65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrfiles) '文件路径
'        .Range("f" & .[f65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrparent) '文件所在位置
'        .Range("g" & .[g65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrsizeb) '文件初始大小
'        .Range("h" & .[h65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arredit) '文件修改时间
'        .Range("i" & .[i65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrsize) '文件大小
'        .Range("j" & .[j65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrdate) '文件创建时间
'        .Range("x" & elow & ":" & "x" & flc + elow - 1) = Now '添加目录的时间
'        .Range("ad" & elow & ":" & "ad" & flc + elow - 1) = 1 '标注文件的来源(是通过添加文件夹的方式添加进来的)
'        .Range("b" & elow - 1).AutoFill Destination:=.Range("b" & elow - 1 & ":" & "b" & flc + elow - 1), Type:=xlFillDefault '添加统一编码
'        .Range("b" & elow & ":" & "b" & flc + elow - 1).Interior.Pattern = xlPatternNone
'        .Range("b" & elow & ":" & "b" & flc + elow - 1).Font.ThemeColor = xlThemeColorLight1
'
'i = .ListCount
'        If i = 0 Then
'        .AddItem
'        GoTo 100
'
'        ElseIf i = 7 Then          '当列表已经存在7个的时候,不断重写第七行的数据
'        For k = 6 To 1 Step -1
'        .List(k, 0) = .List(k - 1, 0)
'        .List(k, 1) = .List(k - 1, 1)
'        .List(k, 2) = .List(k - 1, 2)
'        Next
'
'        GoTo 100
'
'        ElseIf i > 0 And i < 7 Then '当数据小于7的时候,数据向下移动
'        .AddItem
'        For k = i To 1 Step -1
'        .List(k, 0) = .List(k - 1, 0)
'        .List(k, 1) = .List(k - 1, 1)
'        .List(k, 2) = .List(k - 1, 2)
'        Next
'
'100                                                                         '当列表为空,直接在第一行写入数据
'        .List(0, 0) = strx
'        .List(0, 1) = strx1
'        .List(0, 2) = recentfile '统一使用这个时间,确保表格和窗口的数据完全一致
'        End If

'    With Me.ListBox1 '最近阅读
'        m = .ListCount
'        If m = 0 Then
'            .AddItem
'            GoTo 100
'        ElseIf m = 7 Then          '当列表已经存在7个的时候,不断重写第七行的数据
'            For k = 6 To 1 Step -1
'            .List(k, 0) = .List(k - 1, 0)
'            .List(k, 1) = .List(k - 1, 1)
'            .List(k, 2) = .List(k - 1, 2)
'            Next
'            GoTo 100
'        ElseIf m > 0 And m < 7 Then '当数据小于7的时候,数据向下移动
'            .AddItem
'            For k = m To 1 Step -1
'                .List(k, 0) = .List(k - 1, 0)
'                .List(k, 1) = .List(k - 1, 1)
'                .List(k, 2) = .List(k - 1, 2)
'            Next
'100                                                                                 '当列表为空,直接在第一行写入数据
'            .List(0, 0) = strx
'            .List(0, 1) = strx1
'            .List(0, 2) = recentfile
'        End If
'    End With

'With Me.ListBox1
'
'If .ListCount = 0 Then
'.AddItem
'GoTo 100
'
'ElseIf .ListCount = 7 Then          '当列表已经存在7个的时候,不断重写第七行的数据 '重新修改
'For k = 6 To 1 Step -1
'.List(k, 0) = .List(k - 1, 0)
'.List(k, 1) = .List(k - 1, 1)
'.List(k, 2) = .List(k - 1, 2)
'Next
'
'GoTo 100
'
'ElseIf .ListCount > 0 And .ListCount < 7 Then '当数据小于7的时候,数据向下移动
'.AddItem
'n = .ListCount
'For k = n To 1 Step -1
'.List(k, 0) = .List(k - 1, 0)
'.List(k, 1) = .List(k - 1, 1)
'.List(k, 2) = .List(k - 1, 2)
'Next
'
'100                                                                 '当列表为空,直接在第一行写入数据
'.List(0, 0) = strx
'.List(0, 1) = Me.ListView1.SelectedItem.ListSubItems(1).text
'.List(0, 2) = recentfile '统一使用这个时间,确保表格和窗口的数据完全一致
'End If
'
'End With

'If f = 4 And elow > 100 And ifilec > 10 Then '文件夹存在于目录-更新文件的时候检查目录中的文件是否还存在'避免不必要的全部检查
'If ckn < 2 Then
'    strxp = fd.Path
'    itemp = elow - 6
'    itemp1 = 1
'    ReDim arrfilenx(1 To itemp)
'    ReDim arrfilesizex(1 To itemp)
'    ReDim arrfilemodyx(1 To itemp)
'    ReDim arrfilepx(1 To itemp)
'    ReDim arrfilefdx(1 To itemp)
'    ReDim arrfilemd5x(1 To itemp)
'    With ThisWorkbook.Sheets("书库")
'    For J = itemp To 1 Step -1 '删除采用倒回删除的方式
'        If arrfilefd(J, 1) = strxp Then
'            If fso.FileExists(arrfilep(J, 1)) = False Then
'                .Rows(J + 5).Delete Shift:=xlShiftUp '如果目录的文件不存在就删除
'            Else
'                itemp2 = itemp2 + 1
'            End If
'        Else
'            arrfilenx(itemp1) = arrfilen(J, 1)
'            arrfilesizex(itemp1) = arrfilesize(J, 1)
'            arrfilemodyx(itemp1) = arrfilemody(J, 1)
'            arrfilepx(itemp1) = arrfilep(J, 1)
'            arrfilefdx(itemp1) = arrfilefd(J, 1)
'            arrfilemd5x(itemp1) = arrfilemd5(J, 1)
'            itemp1 = itemp1 + 1
'        End If
'    Next
'    ckn = 2
'    elow = itemp1 + 6
'    If itemp1 = 0 Or itemp2 = 0 Then f = 0
'    End With
'End If
'Dim arrfilenx() As String, arrfilesizex() As String, arrfilemodyx() As String, arrfilemd5x() As String, arrfilepx() As String, arrfilefdx() As String
'If itemp < 300 And ifilec > 10 Then '当书库的内容非常少的时候'进行全部的检查'如果这里执行了检查后面就不再进行检查,添加的文件足够多
'                ReDim arrfilenx(1 To itemp) '临时数组
'                ReDim arrfilesizex(1 To itemp)
'                ReDim arrfilemodyx(1 To itemp)
'                ReDim arrfilepx(1 To itemp)
'                ReDim arrfilefdx(1 To itemp)
'                ReDim arrfilemd5x(1 To itemp)
'                For itemp1 = itemp To 1 Step -1
'                    If fso.FileExists(arrfilep(itemp1, 1)) = False Then
'                        .Rows(itemp1 + 5).Delete Shift:=xlShiftUp
'                    Else
'                        arrfilenx(itemp2) = arrfilen(itemp1, 1)
'                        arrfilesizex(itemp2) = arrfilesizex(itemp1, 1)
'                        arrfilemodyx(itemp2) = arrfilemody(itemp1, 1)
'                        arrfilemd5x(itemp2) = arrfilemd5(itemp1, 1)
'                        arrfilefdx(itemp2) = arrfilefd(itemp1, 1)
'                        arrfilepx(itmp2) = arrfilep(itemp1, 1)
'                        itemp2 = itemp2 + 1
'                    End If
'                Next
'                elow = itemp2 + 6
'                If itemp2 = 0 Then
'                    If addcode = 0 Then
'                        ClearAll (0)
'                    Else
'                        ClearAll (1) '清除目录内容
'                        GoTo 202
'                    End If
'                Else
'                ckn = 1
'                End If
'            End If
Sub skidj()

Debug.Print Int("1") + 1
'If Len(Selection) = 0 Then MsgBox 1
'Debug.Print fso.GetDrive("e").AvailableSpace
'Dim fl As File
'Dim flop As Object
'Dim strx As String, i As Long, k As String, strx1 As String, t As Single, j As Long
'strx = "11"
'If IsNumeric(strx) = True Then MsgBox 1
'strx = Cells(32, "aa")
't = Timer
''With Me
''strx = .TextBox20.text
''If Len(strx) = 0 Or fso.FileExists(strx) = False Then .Label57.Caption = "文件存在": Exit Sub
'Set fl = fso.GetFile(strx)
'j = fl.Size
'Set flop = fl.OpenAsTextStream(ForWriting, TristateMixed)
'With flop
'For i = 1 To j
'k = CStr(Randnumx(10000))
'strx1 = GetMD5Hash_String(k)
'.Write strx1
'Next
'.Close
'End With
''End With
'Set fl = Nothing
'Set flop = Nothing
'Debug.Print Format(Timer - t, "0.00000")
'On Error GoTo 100
'Dim fl As File
'Set fl = fso.GetFile(Cells(32, "aa"))
'Dim flt As Object
''Debug.Print fl.Name
''Debug.Print fl.ShortName
''Debug.Print fl.ShortPath
''Debug.Print fl.Type
'Set flt = fl.OpenAsTextStream(ForWriting, TristateUseDefault) '不用用forwriting参数,会导致文件彻底损坏
'flt.Write "Nothing"
'flt.Close
'Exit Sub
'100
'MsgBox Err.Number
'
'  Open address For Binary Access Read Write Lock Read Write As #1 '注意open不支持非ansi字符'判断文件是否处于打开的状态,如果是txt文件就直接跳过
'                  Close #1
'  If Len(.Range("ab4").Value) = 0 And .Range("ab9") = 1 Then '1表示IE存在
'            exepath = "C:\Program Files\Internet Explorer\iexplore.exe "
'        ElseIf .Range("ab4") <> "" Then
'            exepath = .Range("ab4")                                                                 'Environ("SYSTEMDRIVE")表示系统所在的盘符
'            If fso.FileExists(Left$(exepath, Len(exepath) - 1)) = False And .Range("ab9") = 1 Then exepath = Environ("SYSTEMDRIVE") & "\Program Files\Internet Explorer\iexplore.exe " '调用前检查程序的存在,如果存在则重新设置为IE
'        End If
'
'Dim Obj As Object
'Dim UF As Object
'
'For Each Obj In ThisWorkbook.VBProject.VBComponents
'  If Obj.Type = 3 Then
'    Set UF = UserForms.Add(Obj.Name)
'    Debug.Print UF.Name & ", " & UF.Controls.Count
'    Unload UF
'  End If
'Next Obj
'Dim k As Byte
'MsgBox k + 1
'Dim str As String
'str = "12334"
'If IsNumeric(str) = True Then MsgBox 1
'Dim strx As String
'strx = ","
'If strx Like "[一-龥]" Then MsgBox 1
'strx = "C:\Users\Lian\Downloads\Compressed"
'MsgBox fso.GetDriveName(strx)
''MsgBox Environ("SYSTEMDRIVE")
'MsgBox ThisWorkbook.VBProject.VBComponents.
''
'fso.CopyFile (ThisWorkbook.FullName), strx & "\", overwritefiles:=True
'Dim strx As String
'strx = Me.ComboBox11.Value
''
'Kill "C:\Users\Lian\Downloads\text.txt"
'MsgBox ThisWorkbook.VBProject.VBComponents("sheet2").CodeModule.ProcCountLines("Worksheet_SelectionChange", vbext_pk_Proc)
'  Dim SFilename As String
'    SFilename = ThisWorkbook.Path & "\test1.vbs " & """hellow""" 'Change the file path
'
'    ' Run VBScript file
'    Set wshShell = CreateObject("Wscript.Shell")
'    wshShell.Run """" & SFilename & """"
'MsgBox Application.Version
'fso.DeleteFile (Cells(1, 1).Value)
'Sub t2()
''         Debug.Print Format(Now, "yyyy-mm-dd")
''         Debug.Print Format(Now, "yyyy年mm月dd日")
'         Debug.Print Format(Now, "yyyy/mm/dd/h:mm:ss")
''         Debug.Print Format(Now, "d-mmm-yy") '英文月份
''         Debug.Print Format(Now, "d-mmmm-yy") '英文月份
''         Debug.Print Format(Now, "aaaa") '中文星期
''         Debug.Print Format(Now, "ddd") '英文星期前三个字母
''         Debug.Print Format(Now, "dddd") '英文星期完整显示
'   End Sub
'书库详情
'Dim elow As Integer
'With ThisWorkbook.Sheets("书库")
'elow = .[e65536].End(xlUp).Row
'If elow < 6 Then
'.Label1.Caption = "无数据"
'Exit Sub
'ElseIf elow > 100 Then
'UserForm6.Show 0
'Call checkfile
'End If
'Me.Label47.Caption = .Range("p37").Value '文件总数
'Me.Label48.Caption = .Range("p38").Value '所有文件大小
'Me.Label49.Caption = .Range("p40").Value 'pdf
'Me.Label50.Caption = .Range("s40").Value 'EPUB
'Me.Label51.Caption = .Range("p42").Value '其他
'Me.Label52.Caption = .Range("p41").Value 'PPT
'Me.Label53.Caption = .Range("v41").Value 'Word
'Me.Label54.Caption = .Range("s41").Value 'Excel
'if strx="工具设置" or len(strx)= then exit sub
'ThisWorkbook.Save
'With ThisWorkbook.Sheets("temp")
'Select Case toolx
'Case 1: xtool = .Range("ab11").Value
'Case 2: xtool = .Range("ab12").Value
'Case 3: xtool = .Range("ab13").Value
'Case 7: xtool = .Range("ab14").Value
'Case 8: xtool = .Range("ab15").Value
'End Select
'End With
'MsgBox ThisWorkbook.FullName
'strx = CStr(Format(Now, "yyyymmddhmmss"))
'fso.GetFile("C:\Users\Lian\Downloads\test.pdf").Name = strx & ".pdf"
'Dim fso As Object
'Set fso = CreateObject("Scripting.FileSystemObject")
'ThisWorkbook.ChangeFileAccess xlReadOnly '变更文件属性
'Kill ThisWorkbook.FullName '删除文件
'MsgBox "初始化成功，请重新打开文件"
'Call openfilelocation(userpath)          '打开新文件所在的文件夹位置
'ThisWorkbook.Close False
'            If ThisWorkbook.Sheets("首页").Range("d3").Value = 1 Or ThisWorkbook.Sheets("书库").Range("d3").Value = 0 Then
'For i = 3 To 35
'Cells(5, i) = Cells(5, i) & i - 2
'Next
'Shell ("PowerShell_ISE "), vbNormalFocus
'strCommand = "Powershell.exe -ExecutionPolicy ByPass ""C:\Users\****\Documents\lb\file.ps1"" " & FilePath
End Sub

'    sql = "select * from [" & TableName & "$] where 统一编码='" & str1 & "'"                                          '查询数据
'    Set rs = New ADODB.Recordset
'    rs.Open sql, conn, adOpenKeyset, adLockOptimistic
'    If rs.BOF And rs.EOF Then '用于判断有无找到数据
 '    Else
''    MsgBox "存在相同编号"
'    End If
    
'    rs.Close
'    Set rs = Nothing

'With Me
'If Len(.TextBox11.text) = 0 And Len(.TextBox12.text) = 0 And Len(.TextBox22.text) = 0 Then '两者都为空的时候
'Call warning(2)
'Me.TextBox11.SetFocus
'Exit Sub
'End If
'
'If Len(.TextBox11.text) > 0 And fso.FileExists(.TextBox11.text) = True Then
'ThisWorkbook.Sheets("temp").Range("ab4") = Trim(.TextBox11.text) & Chr(32) '浏览器 'chr(32)表示空格符号
'ThisWorkbook.Sheets("temp").Range("ac4") = 1 '标记程序已经设置,不在检查chrome浏览器是否存在
'End If
'If InStr(.TextBox12.text, "exe") > 0 And fso.FileExists(.TextBox12.text) = True Then ThisWorkbook.Sheets("temp").Range("ab6") = Trim(.TextBox12.text) & Chr(32) '截图
'If Len(.TextBox22.text) > 0 And fso.FileExists(.TextBox22.text) = True Then ThisWorkbook.Sheets("temp").Range("ab5") = Trim(.TextBox22.text) & Chr(32) 'Pdf编辑
'End With

'创建动态时间

'                .Worksheets(4).Name = "删除备份"
                '.Worksheets(4).Range("a1:u1") =

'Array("统一编码", "文件名", "文件类型", "文件路径", "文件所在位置", "文件初始大小", "文件大小", "文件创建时间", "标识编号", "文件类别", "最近打开时间", "累计打开次数", "主文件名", "作者", "PDF清晰度", "文本质量", "内容评分", "推荐指数", "标签1", "标签2", "添加时间")
'    ThisWorkbook.VBProject.VBComponents.Item(1).CodeModule.ReplaceLine 109, "lastpath=" & """" & lastpath2 & """"
'timest = False '添加时间控件对于程序的运行产生的影响非常大,很容易导致全面的崩溃
'Call atclock
'        ThisWorkbook.VBProject.VBComponents.Item(1).CodeModule.ReplaceLine 109, "lastpath=" & """" & lastpath1 & """"
        
        'MkDir (userpath) '创建文件
        'SetAttr (userpath), vbHidden '设置文件夹的属性为隐藏 '不支持非ansi字符
'Sub comboxclick() '当类别的筛选这行时combox发生点击时间时使用(修改)
'Me.MultiPage1.Value = 0       '选定多标签页
'Me.ListView2.ListItems.Clear  '清除页面上的内容
'Me.ListView2.Visible = True   '可见
'
'End Sub


'If Len(.TextBox14.text) > 0 Then
'   If 2 * Len(Me.TextBox14.text) = LenB(StrConv(Me.TextBox14.text)) Then '需要修正(当输入的是中文字符时,调用其他的搜索)
'    .TextBox13.text = UDF_Translate(Me.TextBox1.text)
'    .TextBox14.text = Me.TextBox1.text
'   Else
'    .Label63.Caption = ""
'    .TextBox13.text = "" '清空
'    Call WriteVocabulary
'   End If
'End If
'End With



Sub 合并表格()
Dim timea As String
timea = Now
timea = Format(timea, "ddd")
Debug.Print timea
' MsgBox "!数据连接异常,程序无法正常运行", vbCritical
' End Sub
'Dim fd As Folder
'Set fd = fso.GetFolder("D:\文件大小")
'With fd
'Debug.Print
'End With
End Sub
'If rng.Offset(0, 2) Like "xl*" Then '如果文件的类型是excel,那么判断打开的文件是否重名
'   For i = 1 To Workbooks.Count
'   If rng.Offset(0, 1).Value = Workbooks(i).Name Then FileExist = False
'   Next
'End If
'If rng Is Nothing Then '文件是否存在
'FileExist = False
'Else
'FileExist = True
'End If
'Public timest As Boolean
'Public Function atclock() '动态时间
'If timest = True Then
'UserForm1.Label1.Caption = Format(Now, "yyyy-mm-dd HH:MM:SS")
'Sleep 25
'ThisWorkbook.Application.OnTime Now + TimeValue("00:00:01"), "atclock"
'End If
'End Function


'Call deleback(Cells(sx(p), 2), Cells(sx(p), 3), Cells(sx(p), 4), Cells(sx(p), 5), Cells(sx(p), 6), Cells(sx(p), 7), Cells(sx(p), 8), Cells(sx(p), 9), Cells(sx(p), 10), Cells(sx(p), 11), Cells(sx(p), 12), Cells(sx(p), 13), Cells(sx(p), 14), Cells(sx(p), 15), Cells(sx(p), 16), Cells(sx(p), 17), Cells(sx(p), 18), Cells(sx(p), 19), Cells(sx(p), 20), Cells(sx(p), 21)) '执行备份
'j = 1
'ReDim arrcolumn(1 To k)

'Rows(sx(p)).Delete Shift:=xlShiftUp
'Call delefileover(arrcolumn(p)) '善后

'If yesno = vbYes Then
'    For p = 1 To k
''    For Each slc In Selection.Rows
'       filedyn = False
''        trow = slc.Row
''        arrcolumn(j) = .Range("f" & trow).Value '暂时存储位置信息 '需要调整
''        j = j + 1
'        tfile = .Range("e" & trow).Value
'        If fso.FileExists(tfile) = True Then
'            If Range("ab" & trow) = "ERC" Then
'                fso.DeleteFile (tfile)
'            Else
'                DeleteFiles (tfile)
'            End If
'            filedyn = True '文件确实是被删掉(只有当文件被删掉的时候才会产生文件夹的修改时间的变化)
'        End If
'    Next
'End If

'Function ReplacePunctuation(ByVal strText As String, Optional ByVal IsCN As Boolean = False) As String '半角(33-126)与全角(65281-65374) '[\-,\/,\|,\$,\+,\%,\&,\',\(,\),\*,\x20-\x2f,\x3a-\x40,\x5b-\x60,\x7b-\x7e,\x80-\xff,\u3000-\u3002,\u300a,\u300b,\u300e-\u3011,\u2014,\u2018,\u2019,\u201c,\u201d,\u2026,\u203b,\u25ce,\uff01-\uff5e,\uffe5]
'Dim i As Long
'Dim strTemp As String * 1
'Dim strLen As Integer
'
'On Error Resume Next
'For i = 1 To strLen
'    strTemp = Mid(strText, i, 1)
'    If lem(strTemp) > 0 Then
'        i = ThisWorkbook.Application.WorksheetFunction.unicode(strTemp)
'        If IsCN = True Then
'            If i < 126 And i > 33 Then Mid(strText, i, 1) = ChrW(i + 65248)
'        Else
'            If i >= 65218 And i <= 65374 Then Mid(strText, i, 1) = ChrW(i - 65248)
'        End If
'    End If
'Next
'ReplacePunctuation = strText
'End Function


'Dim rg As Range

'        ltma = dic.Keys
'        ltmb = dic.Items
'        ltmc = dica.Items
'        ltmd = dicb.Items
'        ltme = dicc.Items

'    .Activate
'    Union(.Range("b6:e" & .[b65536].End(xlUp).Row), .Range("m6:m" & .[b65536].End(xlUp).Row)).Select '这里需要修正(采用union拼接的数组,被分成两个部分)
'    Set rg = Selection
'    ReDim arr(1 To rg.Areas.Count)
'    For i = 1 To UBound(arr)
'        arr(i) = rg.Areas(i)
'    Next

'If f > 0 Then
'flmd = flx.DateLastModified '时间,比较目录文件夹的修改时间和文件的修改时间
'If flmd < fdmd Then GoTo 20
'End If
'Dim fdmd As Date, flmd As Date '文件修改时间,  '文件修改时间 文件剪切不会导致文件createtime发生变化 '只能用于检测下载类文件夹
'Dim t1 As Date
'Dim t2 As Date
'
't1 = Cells(4, 5).Value
't2 = Cells(5, 5).Value '时间比较

'Sub lop()
'Dim t1 As Date
'Dim t2 As Date
't1 = Cells(1, 1).Value
't2 = Cells(2, 1).Value
'Cells(3, 1) = DateDiff("d", t1, t2) 'datediff时间间隔函数,"d"表示间隔的天数day
'End Sub
'
'If t1 > t2 Then MsgBox 1

'With ThisWorkbook.Sheets("temp")                   '将词库放在打开的Excel表格上
'sTest = .Cells(randomx, "ah") '存储试题 '答案-英文
'arrtemp3(listnum) = .Cells(randomx, "ai") '中文,cells也支持range的模式
'arrtemp1(listnum) = sTest
'Me.Label90.Caption = .Cells(randomx, "ai")
'End With

'Alastrow = ThisWorkbook.Sheets("temp").[ah65536].End(xlUp).Row '演示区域 '参数用于生成随机数

'UserForm3.Show 0
'UserForm3.MultiPage1.Value = 1
'InStr(filePath, ChrW(8226)) > 0 Or InStr(filePath, ChrW(12539)) > 0 Or '存在加重号字符/日文的间隔符以及其他的异常字符,进行过滤(根据实际进行调整,这两个字符是个人的文件中经常出现的)
'ThisWorkbook.VBProject.VBComponents.Item(3).CodeModule.InsertLines 125, "exit sub"
'ThisWorkbook.VBProject.VBComponents.Item(3).CodeModule.DeleteLines 125
'With Sheets("主界面")
'For i = 37 To 46
''.Range("d" & i).UnMerge
''.Range("p" & i).UnMerge
'.Range("e" & i & ":" & "h" & i).Merge Across:=False
'.Range("i" & i & ":" & "j" & i).Merge Across:=False
''.Range("w" & i & ":" & "x" & i).Merge Across:=False
'
''Range("u" & i & ":" & "x" & i).Merge Across:=False
'Next
'End With
'
'With ThisWorkbook.Sheets("书库")
''.Range("y5") = "评分"
'  .Range("x5:z5").Select
''    With Selection
''        .Merge Across:=False
''        .HorizontalAlignment = xlHAlignCenter
''    End With
'End With
'
'   With Selection.Borders(xlEdgeLeft)
'        .Weight = xlThin
'        .LineStyle = xlContinuous
'    End With
'    With Selection.Borders(xlEdgeTop)
'        .Weight = xlThin
'        .LineStyle = xlContinuous
'    End With
'    With Selection.Borders(xlEdgeBottom)
'        .Weight = xlThin
'        .LineStyle = xlContinuous
'    End With
'    With Selection.Borders(xlEdgeRight)
'        .Weight = xlThin
'        .LineStyle = xlContinuous
'    End With
'    Selection.Borders(xlEdgeLeft).ColorIndex = xlColorIndexAutomatic
'    Selection.Borders(xlEdgeTop).ColorIndex = xlColorIndexAutomatic
'    Selection.Borders(xlEdgeBottom).ColorIndex = xlColorIndexAutomatic
'    Selection.Borders(xlEdgeRight).ColorIndex = xlColorIndexAutomatic

'End Function
'Public Flagpause As Boolean, FlagStop As Boolean, Flagnext As Boolean '控制执行的变量-单词训练
'Public timenumx As Integer, starx As Boolean
'
'Sub Timeclock()
'timenumx = timenumx + 1
'UserForm3.Label85.Caption = timenumx & "s"
'End Sub
'
''
'Sub startime()
''Dim i As Integer
'''If starx = True And timenumx <= 10 Then
'''Application.OnTime Now + TimeValue("00:00:01"), "startime"
''''DoEvents
''''Sleep 20
'''Timeclock
'''End If
'''MsgBox "本机操作系统的名称和版本为:" & Application.OperatingSystem Environ("OS")
''i = CInt(Application.Version)
'
'Dim wsh As Object
'Dim wExec As Object
'
'Set wsh = CreateObject("WScript.Shell")
'Set wExec = wsh.Exec("powershell Get-Host | Select-Object Version") '获取powershell版本号
'Result = wExec.StdOut.ReadAll
'
'
'End Sub

Sub 调整按钮的位置()

x = 111

For i = 1 To 13

 If i = 12 Then GoTo 100
    ActiveSheet.Shapes.Range(Array("CommandButton" & i)).Select
    With Selection
       .Left = x
    End With
    x = x + 135
100
    Next
    
End Sub

Sub 显示位置()
'
'elow = Sheets("书库").[c65536].End(xlUp).Row + 1
'
'MsgBox elow

'ThisWorkbook.Sheets("书库").Range("d3").ClearContents
'UserForm3.MultiPage1.Pages(6).visibale = True
'UserForm3.Show
'For i = 0 To 5
'MsgBox UserForm3.MultiPage1.Pages(i).Caption
'Next

'MsgBox UserForm3.Controls.Count

'Sheets("书库").Range("f5:f" & Sheets("书库").[f65536].End(xlUp).Row).AutoFilter

'Range("e39") = "c:\users\lian\download\"

'Sheets("书库").Range("f5:f" & Sheets("书库").[f65536].End(xlUp).Row).AutoFilter Field:=1, Criteria1:="c:\users\lian\download\"

'Dim arr()
'With ThisWorkbook.Sheets("书库")
'
'arr = .Range("d6:d" & [d65536].End(xlUp).Row).Value
'
'End With
'ThisWorkbook.Sheets("书库").Range("v4") = "补充"
Application.ScreenUpdating = True
End Sub

Sub 显示选定的行列()

'MsgBox Selection.Column

'If fso.DriveExists("d") = True Then
'
'MsgBox fso.GetDrive("c").DriveType
'End If

'Windows(ThisWorkbook.Name).Visible = True
'ThisWorkbook.Sheets("书库").Visible = False

'With ThisWorkbook
'    For i = 1 To .Worksheets.Count
'    .Worksheets(i).Protect "Password", UserInterfaceOnly:=True
'    Next
'End With
'
'Range("l11") = 11

'MsgBox Application.ThisWorkbook.VBProject.
Application.DisplayFormulaBar = True

End Sub

Sub ceshi()

With ThisWorkbook.Sheets("书库")
    .Select
    Union(.Range("b6:c" & .[b65536].End(xlUp).Row), .Range("r6:s" & .[b65536].End(xlUp).Row)).Select
    Set rg = Selection
    ReDim arr(1 To rg.Areas.Count)
    For i = 1 To UBound(arr)
        arr(i) = rg.Areas(i)
    Next
    
End With

End Sub

Sub arrtest()
'With ThisWorkbook

''最近阅读
'If .Sheets("主界面").Range("u27") <> "" Then '不添加空白的值进来
'
'
'    For j = 33 To 28 Step -1
'        If .Sheets("主界面").Range("p" & j) <> "" Then Exit For
'    Next
'
'
'arrrc = .Sheets("主界面").Range("p27:w" & j).Value
'
'm = UBound(arrrc)
'
''Me.ListBox1.AddItem
'
'End If
'If .Sheets("主界面").Range("e37") <> "" Then '不添加空白的值进来
'
'For j = 38 To 100
'If .Sheets("主界面").Range("e" & j) = "" Then Exit For
'Next
'
'arral = .Sheets("主界面").Range("e37:i" & j - 1).Value
'End If
'
'End With
'With ThisWorkbook.Sheets("书库")
'Dim arrb()
'arrb = .Range("c6" & ":" & "c" & .[b65536].End(xlUp).Row).Value
'End With

'arra = Array(".xlsm", ".docx", ".pptx", ".txt", ".accdb", ".mobi", ".epub")
'MsgBox UBound(arra)

'MsgBox fso.GetParentFolderName(Range("f10"))

'ThisWorkbook.Sheets("书库").Range("w5") = "添加时间"


End Sub

Sub module() '列出模块的位置

'For i = 1 To 25
'
'ThisWorkbook.Sheets("书库").Range("j" & 6 + i) = ThisWorkbook.VBProject.VBComponents.Item(i).Name
'
'Next
ThisWorkbook.VBProject.VBComponents.item(4).CodeModule.InsertLines 144, "exit sub"

'ThisWorkbook.VBProject.VBComponents.Item(4).CodeModule.DeleteLines 144

End Sub

'Sub HideOption()
'With ThisWorkbook.Windows(1)
'.DisplayFormulas = False
'.DisplayHeadings = False
'.DisplayHorizontalScrollBar = False
'.DisplayVerticalScrollBar = False
'.DisplayWorkbookTabs = False
'End With
'End Sub


Sub ats()
'Sheet3.Activate
'Range("e37:j37").ClearContents
Application.ScreenUpdating = True
'If fso.FolderExists("C:\Users\Master\Downloads\Compressed\") = True Then MsgBox "cz"
End Sub

Sub stand()
Dim arra()
Dim arrB()
Dim arrc()

With ThisWorkbook.Sheets("书库")
If .[e65536].End(xlUp).Row > 6 Then
arr = .Range("e6:e" & .[e65536].End(xlUp).Row).Value
Else
ReDim arr(1 To .[e65536].End(xlUp).Row - 5)
For p = 1 To .[e65536].End(xlUp).Row
arr(p) = .Range("e" & p + 5).Value
Next
End If

arra = Array(".xlsm", ".docx", ".pptx", ".txt", ".accdb", ".mobi", ".epub") 'excel,word,ppt,txt,access,电子书
ReDim arrB(1 To UBound(arra) + 1)
For j = 0 To UBound(arra)
arrB(j + 1) = OpenBy(arra(j))
Next

For n = .[e65536].End(xlUp).Row - 5 To 1 Step -1
    For m = 1 To UBound(arrB)
    
    If OpenBy(arr(n, 1)) <> arrB(m) Then .rows(n + 5).Delete Shift:=xlShiftUp
    
    Next
Next
End Sub

Sub kkkds()
Dim k As Integer, n As Integer, i As Integer, p As Integer
Dim strFile As String
Dim arr()

arr = Range("e6:e627").Value
t = Timer
For p = 1 To 622
strFile = arr(p, 1)
k = Len(strFile)
    For n = 1 To k
        For i = 1 To 100
            If i = 14 Then GoTo 1000
        If InStr(Mid$(strFile, n, 1), ChrW(i + 169)) > 0 Then
        Range("j" & p + 5) = "A"
        GoTo 100
        End If
1000
        Next
    Next
100
Next
MsgBox Format(Timer - t, "0,0000")
End Sub





'Sub kij()
'Dim arr()
'Dim k As Integer
't = Timer
'k = [e65536].End(xlUp).Row
'arr = Range("e6:e" & k).Value
'For i = 1 To k - 5
'If Errorcode(arr(i, 1)) > 0 Then Range("ab" & i + 5) = "Error"
'Next
'MsgBox Format(Timer - t, "0,0000")
'End Sub

Sub lkk()
For i = 1 To Len(Cells(1, 1))
Cells(i, 2) = Mid$(Cells(1, 1), i, 1)
Next
End Sub


Sub kok()
'Dim rngmd As Range
'With ThisWorkbook.Sheets("书库")
'Set rngmd = .Range("k6:k" & .[k65536].End(xlUp).Row).Find("filemd5") '检查是否文件已存在
'End With
'Set rngmd = Nothing
If fso.fileexists(Range("h31")) Then MsgBox 1
End Sub


Sub ki()
'Dim fd As Folder
'Dim arr(1 To 10)
'arr(1) = 1
'arr(2) = 2
'For i = 1 To 10
'If 5 = arr(i) Or arr(i) = "" Then Exit For
'Next
'filePath = Cells(7, 3)
'
'Set fd = fso.GetFolder(filePath)          '将fd指定路径对象
'
'If fd.IsRootFolder Then MsgBox 1
'MsgBox fd.Name
'Set fd = Nothing

'Range("f9") = Str(1)
'i = 1
'Dim arr(1 To 3)
'For k = 4 To [b65536].End(xlUp).Row
'If Range("c" & k) <> "" Then
'arr(i) = Split(Range("c" & k), "\")(UBound(Split(Range("c" & k), "\")))
'i = i + 1
'End If
'Next

Dim fd As Folder
Dim rnglists As Range

Set fd = fso.GetFolder("D:\测试\you are")

With ThisWorkbook.Sheets("目录")
Do
   strp = fd.Path & "\"
   Set fd = fd.ParentFolder
   Set rnglists = Nothing
For i = 3 To .Cells.SpecialCells(xlCellTypeLastCell).Column
   Set rnglists = .Cells(4, i).Resize(.[b65536].End(xlUp).Row, 1).Find(strp, lookat:=xlPart)
   If Not rnglists Is Nothing Then
    F = 2
'    If CInt(.Cells(rnglists.Row, 2)) >= c Then
'If x = 1 Then
a = rnglists.Row + 1
'Else
'a = rnglists.Row
'End If
'    Else
'    a = rnglists.Row + 1
'    End If
    .Cells(a, 1).EntireRow.Insert
    GoTo 110
    End If
Next
'    bc = bc - 1 '文件夹层级上升
Loop Until fd.IsRootFolder

End With


'If fd.IsRootFolder Then MsgBox 1

'Cells(4, 3).Resize([b65536].End(xlUp).Row, 1).Select
110
End Sub

'        For i = 1 To 109         '这里的乱码字符需要根据实际来进行调整
'        k = 160 + i
'        If AscB(CharUpper(ChrW(k))) = k Then GoTo 1000 '这里使用charupper将chrw生成的字符转换成为ansi编码，然后对比ascb转换新的ansi编码字符生成的代码值，如果两者的数值相同则意味着这是vba可以处理的字符
'        Errorcode = InStr(strFile, ChrW(k))
'        If Errorcode > 0 Then Exit For
'1000
'        Next

'Sub nin() '测试使用
'Dim rngmd As Range
'For i = 1 To 2
'filemd5 = UCase(Hashpowershell(Range("e" & 619 + i)))
'With ThisWorkbook.Sheets("书库")
'Set rngmd = .Range("k6:k" & .[k65536].End(xlUp).Row).Find(filemd5) '检查是否文件已存在
'End With
'If rngmd Is Nothing Then
'Cells(637 + i, 3) = 1
'Else
'Cells(637 + i) = 0
'End If
'Next
'Set rngmd = Nothing
'Debug.Print Hashpowershell(Range("e523"))
'End Sub

'Sub kin() '测试
'Dim arr()
't = Timer
'Dim i As Integer
'Dim arrs(1 To 1248)
'arr = Range("e6:e1253").Value
'For i = 1 To 1248
'If Errorcode(arr(i, 1)) > 0 Then arrs(i) = errcodex
'Next
'Range("ac" & [ac65536].End(xlUp).Row + 1).Resize(1248) = Application.Transpose(arrs)
'MsgBox Format(Timer - t, "0.0000")
'Debug.Print Errorcode(Cells(14, 1))
'End Sub



'Function UDF_Translate(strText As String) As String '调用有道查询-查询中文专用
'
'    Dim urlDict As String, urlTranslate As String, xmlText As String
'    Dim PhoneticSymbol As String, TranslateArr
'
'    '整理url
'    urlDict = "http://dict.youdao.com/search?q=" & strText & "&doctype=xml"
'    urlTranslate = "http://fanyi.youdao.com/translate?i=" & strText & "&doctype=xml"
'
'    '使用 WebService 和 FilterXML 获取网络数据
'    On Error GoTo Error01
'    xmlText = Application.WorksheetFunction.WebService(urlDict)
'    'Debug.Print xmlText    '调试用
'    PhoneticSymbol = Application.WorksheetFunction.FilterXML(xmlText, "//phonetic-symbol")
'    TranslateArr = Application.WorksheetFunction.FilterXML(xmlText, "//translation/content[1]")
'
'    '整理数据
'    If IsArray(TranslateArr) Then
'        UDF_Translate = "[" & PhoneticSymbol & "] " & Join(Application.WorksheetFunction.Transpose(TranslateArr), "；")
'    Else
'        UDF_Translate = "[" & PhoneticSymbol & "] " & TranslateArr
'    End If
'    'Debug.Print Translate  '调试用
'Exit Function
'Error01:    '若有道词典报错
'    On Error GoTo Error02
'    xmlText = Application.WorksheetFunction.WebService(urlTranslate)
'    'Debug.Print xmlText    '调试用
'    UDF_Translate = Application.WorksheetFunction.FilterXML(xmlText, "//translation")
'    'Debug.Print Translate  '调试用
'Exit Function
'Error02:    '若有道翻译报错
'    UDF_Translate = Err.Description
'End Function


'Option Explicit
'
'Dim arrfiles(1 To 10000)                  '定义一个数组后面用以存放path数据(数值可变)
'Dim arrbase(1 To 10000)                   '存储不包含扩展名的文件名
'Dim arrextension(1 To 10000)              '存储文件扩展名
'Dim arrsize(1 To 10000)                   '存储文件的大小
'Dim arrparent(1 To 10000)                 '存储文件所在位置
'Dim arrdate(1 To 10000)                   '存储文件创建日期
'Dim arrsizeb(1 To 10000)                  '文件的大小,单位比特
'Dim flc As Integer                        '不同sub之间调用相同的变量,注意要使用模块级的定义
'Sub fleshdata(arr()) '更新数据
'
'Dim filepath$
'Dim fd As Folder
'Dim i As Integer, j As Integer
'Dim elow As Integer
'Dim k As Integer
'Dim arrc()
''Dim fso As Object
''Set fso = CreateObject("Scripting.FileSystemObject") '无需设置引用创建fso对象(后期绑定)
'
'Application.ScreenUpdating = False         '关闭屏幕实时更新,加快代码的运行速度
'With ThisWorkbook.Sheets("书库")
'ThisWorkbook.VBProject.VBComponents.Item(3).CodeModule.InsertLines 114, "exit sub"
'If .Range("b6") <> "" Then Call checkfile '检查文件是否存在
'
'For j = 0 To UBound(arr())
'
'If arr(j) = "" Then GoTo 100              '注意此处arr("j")提前等于空
'
'filepath = arr(j)
'
'flc = 0                                   '文件数量的初始值
'
'Set fd = fso.GetFolder(filepath)          '将fd指定路径对象
'
'search fd                                     '调用sf子sub
'
'If flc = 0 Then GoTo 100  '当添加的文件夹无新文件时,进入下一个循环
'
'    .Range("c" & .[c65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrbase) '文件名
'    .Range("d" & .[d65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrextension) '文件扩展名
'    .Range("e" & .[e65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrfiles) '文件路径
'    .Range("f" & .[f65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrparent) '文件所在位置
'    .Range("g" & .[g65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrsizeb) '文件初始大小
'    .Range("h" & .[h65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrsize) '文件大小
'    .Range("i" & .[i65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrdate) '文件创建时间
'    .Range("w" & .[w65536].End(xlUp).Row & ":" & "w" & .[w65536].End(xlUp).Row + flc) = Now
'    .Range("b" & .[b65536].End(xlUp).Row).AutoFill Destination:=.Range("b" & .[b65536].End(xlUp).Row & ":" & "b" & .[b65536].End(xlUp).Row + flc), Type:=xlFillDefault '添加统一编码
'100
'
'Next
'
'Call deletempfile
'End With
'ThisWorkbook.VBProject.VBComponents.Item(3).CodeModule.DeleteLines 114
'Application.ScreenUpdating = True         '运行代码完毕后启用屏幕更新
'
'End Sub
'
'Sub search(ByVal fd As Folder)               'ByVal,为按值传递方式, 有别于byref(按引用方式传递)
'                                         'searchfiles(ByVal fd As Folder)起到function的作用
'Dim fl As File
'Dim sfd As Folder
'Dim arr()
'Dim i As Integer
'Dim arrtemp()
'
'With ThisWorkbook.Sheets("书库")
'
'If .[c65536].End(xlUp).Row - 5 > 1 Then arrtemp = .Range("c6:c" & .[c65536].End(xlUp).Row).Value '尽管value为range或cell的缺省值,但是还是要写上,防止程序有时无法有效识别该数据
'
'For Each fl In fd.Files
'
'    flc = flc + 1
'
'If .[c65536].End(xlUp).Row - 5 > 1 Then
'
'    For i = 1 To .[b65536].End(xlUp).Row - 5
'    If fso.GetFileName(fl.Path) = arrtemp(i, 1) Then '文件名相同的不添加
'
'    flc = flc - 1
'
'    GoTo 100
'
'    End If
'
'    Next
'
'    ElseIf .[c65536].End(xlUp).Row - 5 = 1 Then
'
'    If fso.GetFileName(fl.Path) = .Range("b6").Value Then '文件名相同的不添加
'
'    flc = flc - 1
'
'    GoTo 100
'
'    End If
'
'    Else              '空值,直接跳过
'
'    GoTo 1001
'
'End If
'
'
'1001
'    arrbase(flc) = fso.GetFileName(fl.Path)
'    arrfiles(flc) = fl.Path
'    arrextension(flc) = fso.GetExtensionName(fl.Path)
'    arrparent(flc) = fl.ParentFolder
'    arrdate(flc) = fl.DateCreated
'    arrsizeb(flc) = fl.Size
'
'    If fl.Size < 1048576 Then
'        arrsize(flc) = Format(fl.Size / 1024, "0.00") & "KB"    '文件字节大于1048576显示"MB",否则显示"KB"
'        Else
'        arrsize(flc) = Format(fl.Size / 1048576, "0.00") & "MB"
'    End If
'
'100
'
'Next fl
'
'If fd.SubFolders.Count = 0 Then Exit Sub  '子文件夹数目为零则退出sub
'
'For Each sfd In fd.SubFolders             '搜索子文件夹
'    search sfd
'Next
'End With
'End Sub

'Private Function GetMD5Hash_File(ByVal strFile As String, ByVal lSize As Long) As String
'Dim lFile As Long
'Dim Bytes() As Byte
'If (lSize) Then
'lFile = FreeFile
'ReDim Bytes(lSize - 1)
'Open strFile For Binary As lFile
'Get lFile, , Bytes
'Close lFile
'GetMD5Hash_File = GetMD5Hash_Bytes(Bytes)
'End If
'End Function


'Private Function GetMD5Hash_File(ByVal strfile As String) As String
'Dim lFile As Long
'Dim lSize As Long
'Dim Bytes() As Byte
''Dim arr
''k = Len(strFile)
''ReDim arr(1 To k)
''For i = 1 To k
''Cells(i, 2) = mid$(Cells(1, 1), i, 1)
''Cells(i, 3).FormulaR1C1 = "=UNICODE(RC[-1])"
''arr(i) = ChrW(Cells(i, 3))
''Next
'
'
'lSize = fso.GetFile(strfile).Size
'If (lSize) Then
'lFile = FreeFile
'ReDim Bytes(lSize - 1)
'Open "d:\" & ChrW(112) & "\p.pdf" For Binary As lFile
'Get lFile, , Bytes
'Close lFile
'GetMD5Hash_File = GetMD5Hash_Bytes(Bytes)
'End If
'End Function
'
'
'Sub kjj()
'Dim les As Long
'les = fso.GetFile(Range("e562").Value).Size
'Debug.Print GetMD5Hash_File(Range("e562").Value, les)
'
'End Sub

        If Not rnglist Is Nothing Then 'instr>0
'        frowa = rnglist.Row
'        Set rnglist = .Cells(4, bc).Resize(.[b65536].End(xlUp).Row, 1).FindPrevious(rnglist) '查找前面的值
'        faddr = rnglist.address '记录下地址
'        frow = rnglist.Row
'105
'
'            Set fd = fso.GetFolder(rnglist.Value) 'Sheets(1).Cells.SpecialCells(xlCellTypeLastCell).Row
'            Do
'            Set fd = fd.ParentFolder
'                If fd.Path = strfolder Then '找到匹配值
'103
'               If frowa > frow Then
'               a = frow
'               Else
'               a = frowa
'               End If
'                'a = rnglist.Row
'                .Cells(a, 1).EntireRow.Insert
'                f = 2
'                GoTo 104
'                Else
'                    If c = UBound(Split(rnglist.Value, "\")) Then '同级文件比较
'                       If fdp = fd.Path Then GoTo 103
'                    End If
'                End If
'            Loop Until fd.Drive & "\" = fd.ParentFolder '循环到第一层文件
'            Set rnglist = .Cells(4, bc).Resize(.[b65536].End(xlUp).Row, 1).FindPrevious(rnglist) '没有匹配到值
'                If rnglist.address = faddr Then '循环到起始值
'                a = .[c65536].End(xlUp).Row + 1
'                GoTo 104
'                End If
'            GoTo 105
'          End If
'          a = .[c65536].End(xlUp).Row + 1
'     End If

'    .Activate
'    Union(.Range("b6:c" & .[b65536].End(xlUp).Row), .Range("r6:s" & .[b65536].End(xlUp).Row)).Select
'    Set rg = Selection
'    ReDim arr(1 To rg.Areas.Count)
'    For i = 1 To UBound(arr)
'        arr(i) = rg.Areas(i)
'    Next

On Error GoTo 100 '防止读取不到注册表信息
With CreateObject("wscript.shell") '读取注册表的浏览器信息
cversion = .RegRead("HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon\version") '此位置存储有浏览器的版本信息
If cversion = "" Then Exit Sub '没有获取到信息


Sub popoio()
''Range("b6:b8").Interior.Pattern = xlPatternNone
'''Range("b6:b8").Interior.Pattern = xlPatternNone
''Range("b6:b8").Font.ThemeColor = xlThemeColorLight1
''Range("b6:b8").Borders(xlDiagonalUp).LineStyle = xlNone
''Range("b6:b8").Borders(xlEdgeTop).LineStyle = xlNone
''                If j >= 49 Then
''                    .Range("e" & j + 1 & "h" & j + 1).Merge
''                    .Range("i" & j + 1 & "j" & j + 1).Merge
''                End If
'Range("b7:b8").Borders(xlEdgeBottom).LineStyle = xlNone
    With CreateObject("wscript.shell")
        Result = .RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\7-Zip\DisplayIcon") '获取后缀名对应的注册类型
    End With
''Dim fd  As Folder
''Set fd = fso.GetFolder("D:\b") 'Environ("HOMEDRIVE")
'''Debug.Print fd.ParentFolder.Path
'''Debug.Print Environ("SYSTEMDRIVE")
''Debug.Print fd.Files.Count
'''Debug.Print Environ("UserProfile") 'Environ("ALLUSERSPROFILE")
''Set fd = Nothing
''If Environ("HOMEPATH") = "C:\Users\Lian" Then MsgBox 1
End Sub



Sub kjjk()
'Dim i As Integer
'i = 1
'Call lkklo(i)
'for i =1 to
Dim k As Integer, i As Integer
For i = 1 To ThisWorkbook.VBProject.VBComponents.Count
Cells(i, 1) = ThisWorkbook.VBProject.VBComponents.item(i).CodeModule.CountOfLines
Next
End Sub



Sub lkklo(ByVal x As Integer)

Dim k As Integer

If x > 0 Then k = k + 1

lkklo (k)

End Sub

Sub EnumSEVars() '文件夹变量
        Dim strVar As String
        Dim i As Long
        For i = 1 To 255
            strVar = Environ$(i)
            If LenB(strVar) = 0& Then Exit For
            Cells(i, 1) = strVar
        Next
End Sub

'注意变量使用时用到的双引号"" ""
'If Me.TextBox11.text <> "" And Me.TextBox12 <> "" Then                                                                               '两个设置都设置时
'   If fso.FileExists(mid$(Me.TextBox11.text, 2, Len(Me.TextBox11.text) - 3)) = False Or fso.FileExists(mid$(Me.TextBox12.text, 2, Len(Me.TextBox12.text) - 3)) = False Then
'   Call warning(2)
'   Me.TextBox1.SetFocus
'   Exit Sub
'   Else
'   exepatha1 = Me.TextBox11.text
'   exepatha2 = Me.TextBox12.text
'   ThisWorkbook.VBProject.VBComponents.Item(8).CodeModule.ReplaceLine 33, "exepath = " & "" & exepatha1 & ""          '注意这里代码的位置在后续修改中可能出现的位置变化
'   ThisWorkbook.VBProject.VBComponents.Item(9).CodeModule.ReplaceLine 256, "exepath = " & "" & exepatha1 & ""
'   ThisWorkbook.VBProject.VBComponents.Item(9).CodeModule.ReplaceLine 467, "exepath = " & "" & exepatha2 & ""
'   Call warning(1)
'   End If
'End If
'
'If Me.TextBox11.text <> "" And Me.TextBox12.text = "" Then
'   If fso.FileExists(mid$(Me.TextBox11.text, 2, Len(Me.TextBox11.text) - 3)) = False Then
'   Call warning(2)
'   Me.TextBox11.SetFocus
'   Exit Sub
'   Else
'   exepatha1 = Me.TextBox11.text
'   ThisWorkbook.VBProject.VBComponents.Item(8).CodeModule.ReplaceLine 33, "exepath = " & "" & exepatha1 & ""
'   ThisWorkbook.VBProject.VBComponents.Item(9).CodeModule.ReplaceLine 256, "exepath = " & "" & exepatha1 & ""
'   Call warning(1)
'   End If
'End If
'
'If Me.TextBox12.text <> "" And Me.TextBox11.text = "" Then
'   If fso.FileExists(mid$(Me.TextBox12.text, 2, Len(Me.TextBox12.text) - 3)) = False Then
'   Call warning(2)
'   Me.TextBox12.SetFocus
'   Exit Sub
'   Else
'   exepatha2 = Me.TextBox12.text
'   ThisWorkbook.VBProject.VBComponents.Item(9).CodeModule.ReplaceLine 467, "exepath = " & "" & exepatha2 & ""
'   Call warning(1)
'   End If
'End If

'Private Sub UserForm_Initialize()
'Dim i As Integer, m As Integer, k As Integer, s As Integer
'Dim arr(1 To 50)
's = 1
'l = 3
'For m = 4 To [b65536].End(xlUp).Row
'For n = 3 To Cells.SpecialCells(xlCellTypeLastCell).Column - 1
'If Cells(m, n) <> "" Then
'p = p + 1
'Next
'Next
'
'With Me.TreeView1.Nodes
'.Add , , "Menus", "Menus" '根目录
'For i = 4 To [b65536].End(xlUp).Row
'
'    If Cells(i, 3) <> "" Then
'        If arr(s) <> Cells(i, 3) Then
'        arr(s) = Cells(i, 3)
'        .Add "Menus", 4, arr(s), Split(arr(s), "\")(UBound(Split(arr(s), "\")))
'        End If
'    End If
'm = i
'
'
'For k = 3 To Cells.SpecialCells(xlCellTypeLastCell).Column - 1
'If InStr(Cells(m + 1, k + 1), "\") > 0 Then
'If arr(s + 1) <> Cells(m + 1, k + 1) Then
'arr(s + 1) = Cells(m + 1, k + 1)
'If IsEmpty(arr(s)) Then
'arr(s) = "Null" & i
'.Add "Menus", 4, arr(s), "Null"
'End If
'.Add arr(s), 4, arr(s + 1), Split(arr(s + 1), "\")(UBound(Split(arr(s + 1), "\")))
's = s + 1
'End If
'End If
'm = m + 1
'Next
's = 1   '重置
'Next
'    End With
'    Me.TreeView1.Nodes(1).Expanded = True
'End Sub

Sub aoso()
MsgBox Environ("ProgramW6432")
End Sub
'已添加文件夹
'If .Sheets("书库").Range("b6") = "" Then GoTo 1005
'If .Sheets("书库").[d65536].End(xlUp).Row = 6 Then
'Me.ListBox3.AddItem
'Me.ListBox3.List(Me.ListBox3.ListCount - 1, 0) = .Sheets("书库").Range("f6").Value
'Me.ListBox3.List(Me.ListBox3.ListCount - 1, 0) = .Sheets("书库").Range("w6").Value
'Else
'Dim dicadd As New Dictionary
'Dim foldnum As Integer
'arral = .Sheets("书库").Range("f6:f" & .Sheets("书库").[f65536].End(xlUp).Row).Value
'For foldnum = 1 To .Sheets("书库").[f65536].End(xlUp).Row - 5
'dicadd(arral(foldnum, 1)) = ""
'Next
'End If
'Me.ListBox3.List = dicadd.Keys

'已添加文件夹

Sub kl()
Dim k As Integer, i As Integer, j As Integer, m As Integer, lc As Integer
'lc = ThisWorkbook.VBProject.VBComponents.Count


'For i = 1 To lc
'With ThisWorkbook.VBProject.VBComponents.Item(i).CodeModule
'k = k + .CountOfLines
'j = j + .CountOfDeclarationLines
''m = m + .ProcCountLines(
'End With
'Next

'With Me
'.Label1.Caption = k
'End With
m = m + Sheet2.OLEObjects.Count
'Next
'ThisWorkbook.VBProject.VBComponents.
Debug.Print m
End Sub

                
'                .Range("d" & rngad.Row & ":" & "j" & rngad.Row).ClearContents '清除掉内容
'                If rngad.Row <> .[d65536].End(xlUp).Row Then '如果是最后一行,就直接处理
'                    k = rngad.Row
'                    j = .[d65536].End(xlUp).Row
'                    For i = k To j - 1
'                    .Range("d" & i) = .Range("d" & i + 1)
'                    Next
'                End If
'            End With
'        End If


'If ThisWorkbook.Sheets("书库").Range("d" & addrowx).Value <> "txt" And Len(ThisWorkbook.Sheets("书库").Range("ab" & addrowx).Value) = 0 Then '不为txt文件'路径不存在特殊字符
'
'Open address For Binary Access Read Write Lock Read Write As #1  '判断文件是否处于打开的状态不支持非ansi编码, 如果是txt文档,那么就不需要判断是否打开(txt文件在打开的状态下不会被锁定(notepad))
'Close #1
'
'If Err.Number <> 0 Then '利用文件打开状态操作文件出现错误来判断文件是否处于打开的状态
'      Me.Label1.Caption = "文件正在使用中,请关闭后再试"
'      Err.Clear
'      Exit Sub
'End If
'
'End If
'
'On Error GoTo 101
'
'newname = .TextBox1.text
'
'With ThisWorkbook.Sheets("书库")
'str1 = newname & "." & .Range("d" & addrowx)
'If fso.FileExists(.Range("f" & addrowx) & "\" & str1) = True Then
'Me.Label1.Caption = "文件重名" '判断文件是否重名
'Me.TextBox1.text = ""
'Exit Sub
'Else
'
'If Errcode(newname, "N", 1) > 0 Then '当文件不重名就去检查是否存在异常字符
'Me.Label1.Caption = "输入的文件名存在异常字符" '修改文件名的是否存在异常字符''如果输入的名称出现重名/存在空格,将重新执行输入-检查循环(inputbox)
'Me.TextBox1.text = ""
'Exit Sub
'Else
'
'If Len(ThisWorkbook.Sheets("书库").Range("ab" & addrowx).Value) > 0 Then '文件路径存在异常字符
'
'Shell "cmd /c rename " & address & Chr(32) & str1, 0   '调用cmd 的rename命令去执行'可以处理异常字符'但是无法反馈文件是否处于打开的状态
'Sleep 100 '这里需要非常注意,cmd执行的速度,先暂时暂缓后续代码的执行,等待cmd命令的执行,否则fso无法判断新的文件名的存在
'DoEvents
'str1 = .Range("f" & addrowx) & "\" & str1 '新的文件路径
'If fso.FileExists(str1) = True Then '检查文件是否已经修改成功
'.Range("e" & addrowx) = str1
'.Range("c" & addrowx) = newname & "." & Range("d" & addrowx)
'Else
'Me.Label1.Caption = "文件处于打开状态,改名失败"
'End If
'
'Else
'Name address As .Range("f" & addrowx) & "\" & newname & "." & .Range("d" & addrowx) 'name不支持非ansi编码的字符串
'.Range("e" & addrowx) = .Range("f" & addrowx) & "\" & newname & "." & .Range("d" & addrowx)
'.Range("c" & addrowx) = newname & "." & Range("d" & addrowx)
'End If
'
'
'
'
'End If
'.Range("ae" & addrowx) = "" '如果原来的名字存在异常字符那么就去掉这个标记
'.Range("ab" & addrowx) = ""
'
'If .Range("d" & addrowx) = "txt" Then Me.Label2.Caption = "txt文档可以在打开的状态下移动\重命名\删除"      '提醒用户,txt文档可以在打开的状态下移动和重命名等操作
'str = .Range("f" & addrowx).Value & "\"
'End With
'If ThisWorkbook.Sheets("书库").Range("ad" & addrowx).Value = 1 Then
'Set rngtime = ThisWorkbook.Sheets("目录").Cells(4, 3).Resize(ThisWorkbook.Sheets("目录").[b65536].End(xlUp).Row, ThisWorkbook.Sheets("目录").Cells.SpecialCells(xlCellTypeLastCell).Column).Find(str, lookat:=xlWhole)
'If Not rngtime Is Nothing Then rngtime.Offset(0, 2) = Now '文件名修改成功后所在的文件夹的修改时间发生变更
'End If
'.Label1.Caption = "修改成功!"
'.CommandButton2.Enabled = True '
'End With
'
'Set rngtime = Nothing
'
'Exit Sub
'
'101
'Err.Clear
'Me.Label1.Caption = "修改文件名失败"


'    ElseIf c = 2 Then
'    tfolder = tfolder & "\"
''    Set fd = fso.GetFolder(tfolder)
''    Set fd = fd.ParentFolder
'    tfolderp = fd.Path & "\"
'    Set rngad = .Range("f6:f" & .[f65536].End(xlUp).Row).Find(tfolderp, lookat:=xlPart) '检查文件夹是否还有其他的文件存在目录
'    If rngad Is Nothing Then '清除所有这个文件夹的关联文件夹
'    With ThisWorkbook.Sheets("目录")
'    Set rngad = .Cells(4, 3).Resize(.[b65536].End(xlUp).Row, .Cells.SpecialCells(xlCellTypeLastCell).Column).Find(tfolder, lookat:=xlWhole) '精确搜索
'    If Not rngad Is Nothing Then
'    If .AutoFilterMode = True Then .AutoFilterMode = False '筛选如果处于开启状态则关闭
'   .Range("a3:a" & .[a65536].End(xlUp).Row).AutoFilter Field:=1, Criteria1:=rngad.Offset(0, -2).Value
'   .Range("a4").Resize(.[a65536].End(xlUp).Row - 3).SpecialCells(xlCellTypeVisible).Delete Shift:=xlShiftUp '删除掉筛选出来的结果
'   .Range("a3:a" & .[a65536].End(xlUp).Row).AutoFilter
'    End If
'    End With
'    End If
    
'    Else
'    Set fd = fso.GetFolder(tfolder) '注意不要放进do循环里面
'    Do
'    tfolderp = fd.Path & "\" '注意这里的位置和文件夹的层级
'    Set fd = fd.ParentFolder '层级上升
'    Set rngad = .Range("f4:f" & .[f65536].End(xlUp).Row).Find(tfolderp, lookat:=xlPart) '检查文件夹是否还有其他的文件存在目录
'    If Not rngad Is Nothing Then n = n + 1 '清除所有这个文件夹的关联文件夹
'    Loop Until fd.IsRootFolder
'    If n = 0 Then  '清除所有关联文件
    
    
    
'
'    End If
'    End With
    
'
'    tfolderp = tfolder & "\"
'    With .Sheets("目录")
'    Set rngad = .Cells(4, 3).Resize(.[b65536].End(xlUp).Row, .Cells.SpecialCells(xlCellTypeLastCell).Column).Find(tfolderp, lookat:=xlWhole) '精确搜索
'    If Not rngad Is Nothing Then
'    rngadr1 = rngad.address '第一个位置
'    rngadrowx = rngad.Row
'    rngadcolumnx = rngad.Column
'    Set rngad = .Cells(4, rngadcolumnx + 1).Resize(.[b65536].End(xlUp).Row, .Cells.SpecialCells(xlCellTypeLastCell).Column).Find(tfolderp, lookat:=xlPart) '模糊搜索,是否有子文件夹
'        If rngad Is Nothing Then '如果没有
'        Rows(rngadrowx).Delete Shift:=xlShiftUp
'        Else
'        If rngad.address <> rngadr1 And yesno = vbYes Then rngad.Offset(0, 2) = Now '删除后文件夹的时间发生变化' '存在有子文件夹
'        End If
'    End If
'    End With
'
'End With
'Set fd = Nothing


'
'Function tempdele() '需要注意行号的变化
'    ThisWorkbook.VBProject.VBComponents.Item(1).CodeModule.DeleteLines 4  '删除掉部分代码（删除后产生的位置变化）'首次运行的使用
'    ThisWorkbook.VBProject.VBComponents.Item(1).CodeModule.DeleteLines 11
'    ThisWorkbook.VBProject.VBComponents.Item(1).CodeModule.DeleteLines 36
'    ThisWorkbook.VBProject.VBComponents.Item(1).CodeModule.DeleteLines 169
'    ThisWorkbook.VBProject.VBComponents.Item(1).CodeModule.DeleteLines 146 '删除掉执行
'    ThisWorkbook.VBProject.VBComponents.Item(24).CodeModule.DeleteLines 29
'End Function

'Public starx As Boolean, Flagpause As Boolean, FlagStop As Boolean
'Public i As Integer, k As Integer
'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
' Sub startime()
' Dim p As Integer
'If starx = True And i <= 24 Then
'If FlagStop = True Then Exit Sub
'If Flagpause = True Then GoSub 1
'Application.OnTime Now + TimeValue("00:00:01"), "startime"
'sk
'End If
'
'Exit Sub
'1:
'For p = 0 To 1 Step 0
'Sleep 25
'        DoEvents
'        If Flagpause = False Then Return '继续执行
'        If FlagStop Then Exit Sub
'Next
'
'End Sub
'
'Sub sk()
'i = i + 1
'UserForm1.Label1.Caption = i & "s"
''If Int(i / 5) = k + 1 Then
''DoEvents
'''Application.Speech.Speak ("water")
'''heloop
''speakvs ("water")
''DoEvents
''k = k + 1
''End If
'End Sub

'Function speakvs(ByVal strx As String)
'Dim wshShell As Object
'
'SFilename = "C:\Users\Lian\Documents\lb\test1.vbs" & """" & strx & """"
'Set wshShell = CreateObject("Wscript.Shell")
'    wshShell.Run """" & SFilename & ""
'End Function

'Sub opopooio()
'  Dim SFilename As String
'    SFilename = ThisWorkbook.Path & "\test1.vbs " & """hellow""" 'Change the file path
'
'    ' Run VBScript file
'    Set wshShell = CreateObject("Wscript.Shell")
'    wshShell.Run """" & SFilename & """"
'End Sub


'Sub kdjfj()
'speakvs ("water")
'End Sub

' Sub heloop()
' Call opo("hellow")
' End Sub
'
' Sub opo(strtxt As String)
' Dim strx As String
'Dim scr As ScriptControl: Set scr = New ScriptControl
'strx = "SAPI.SpVoice"
'scr.Language = "VBScript"
'scr.AddCode "sub T: " & "CreateObject(" & """" & strx & """" & ").Speak" & """" & strtxt & """" & ": end sub"
'scr.Run "T"
' End Sub


'Private Sub CommandButton1_Click()
'i = 0
'k = 0
'starx = True
'FlagStop = False
'Flagpause = False
'Call startime
'
'End Sub
'
'Private Sub CommandButton2_Click()
'starx = False
'End Sub
'
'Private Sub CommandButton3_Click()
'Flagpause = True
'End Sub
'
'Private Sub CommandButton4_Click()
'FlagStop = True
'End Sub
'
'Private Sub CommandButton5_Click()
'Flagpause = False
'End Sub



'Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
'Private Sub Form_Load()
'Dim LocaleID As Long
'LocaleID = GetSystemDefaultLCID
'Select Case LocaleID
'Case &H404
'MsgBox "当前系统为：繁体中文", , "语言"
'Case &H804
'MsgBox "当前系统为：简体中文", , "语言"
'Case &H409
'MsgBox "当前系统为：英文", , "语言"
'End Select
'End Sub






Sub searchWordFromdict(tmpWord As String, tmpTrans As String, tmpPhonetic As String)
'
'''''"http://www.dictionary.com/browse/" & tmpWord &"?s=t"
'
'    '"http://www.iciba.com/" & tmpWord
'    Dim XH As Object
'    Dim s() As String
'    Dim str_tmp As String
'    Dim str_base As String
'    Dim URL As String
'    tmpTrans = ""
'    tmpPhonetic = ""
'
'
'   ' tmpWord = Replace(tmpWord, " ", "_")
''        URL = "https://www.dictionary.com/browse/" & tmpWord
'    URL = "http://dict.cn/" & tmpWord
'
'    '开启网页
'    Set XH = CreateObject("Msxml2.XMLHTTP") 'Microsoft.XMLHTTP")
''       On Error Resume Next
'    XH.Open "get", URL, 0 'True
'    XH.Send (Null)
''       On Error Resume Next
'    While XH.readyState <> 4
'        DoEvents
'    Wend
'    str_base = XH.responsetext
''Debug.Print Mid(str_base, 4001, 3000)
'    'XH.Close
'    Set XH = Nothing
'
'        '取得音标部分
'    ybEN = "英 " & Split(Split(str_base, "EN-US"">")(1), "<")(0)
'    ybUS = " 美 " & Split(Split(str_base, "EN-US"">")(2), "<")(0) ' & "]"
'
'    str_base1 = Split(Split(str_base, "<ol slider=""2"">")(1), "</ol>")(0)
'    v = Split(str_base1, "<li>")
'
'        '取得中文含义部分
'    For i = LBound(v) + 1 To UBound(v) - 1
'        hytmp = hytmp & "；" & Split(v(i), "<")(0)
'    Next i
'    tmpPhonetic = ybEN & " " & ybUS
'    tmpTrans = Mid(hytmp, 2)
End Sub
'http://www.iciba.com/branch

Sub searchWordFromhjxd(tmpWord As String, tmpTrans As String, tmpPhonetic As String) '''***
'        Dim XH As Object
'        Dim s() As String
'        Dim str_tmp As String
'        Dim str_base As String
'        Dim URL As String
'        tmpTrans = ""
'        tmpPhonetic = ""
'        'https://dict.hjenglish.com/notfound/es/knowledge
'        URL = "https://dict.hjenglish.com/notfound/es/" & tmpWord
'        '开启网页
'        Set XH = CreateObject("Msxml2.XMLHTTP") 'Microsoft.XMLHTTP")
' '       On Error Resume Next
'        XH.Open "get", URL, 0 'True
'        XH.Send '(Null)
' '       On Error Resume Next
'        While XH.readyState <> 4
'            DoEvents
'        Wend
'        str_base = XH.responsetext
'
'        'XH.Close
'        Set XH = Nothing
'
' Debug.Print URL 'Mid(str_base, 6001, 3000)
'            '取得音标部分
'    ybEN = "英 " & Split(Split(str_base, "value-en"">")(1), "<")(0)
'    ybUS = " 美 " & Split(Split(str_base, "-us"">")(1), "<")(0)
'
'        str_base1 = Split(Split(str_base, "<ul>")(2), "</ul>")(0)
'        v = Split(str_base1, "<span")
'
'            '取得中文含义部分
'        For i = LBound(v) + 1 To UBound(v)
'            hytmp = hytmp & Split(v(i), "</span>")(0)
'        Next i
'        tmpPhonetic = ybEN & " " & ybUS
'tmpTrans = Mid(Replace(Replace(hytmp, "class=""attr"">", Chr(10)), ">", ""), 3)
'
End Sub
'http://www.iciba.com/branch

Sub searchWordFromBaidu(tmpWord As String, tmpTrans As String, tmpPhonetic As String)
'    'http://dict.baidu.com/s?wd=单词  'https://fanyi.baidu.com/?aldtype=85#en/zh/truths
'    'https://fanyi.baidu.com/?#en/zh/truths
'    Dim XH As Object
'    Dim str_base As String, URL
'    If Len(tmpWord) = 0 Then Exit Sub
'    tmpTrans = "": tmpPhonetic = ""
'    '开启网页
'    Set XH = CreateObject("Microsoft.XMLHTTP")
'    URL = "http://www.baidu.com/s?ie=UTF-8&wd=" & tmpWord
'    XH.Open "GET", URL, 0 'True
'    XH.Send (Null)
'    While XH.readyState <> 4
'        DoEvents
'    Wend
'    str_base = XH.responsetext
'    'XH.Close
'    Set XH = Nothing
'   str_base = Replace(str_base, "&#039;", "'")
'    '取得音标部分’
'    ybEN = "英 [" & Split(Split(str_base, ">[")(1), "]<")(0) & "] "
'    ybUS = "美 [" & Split(Split(str_base, ">[")(2), "]<")(0) & "]"
'    '中文含义
'    ybZEHY = Replace(Replace(Mid(Split(Split(str_base, "text2"">")(1), "<")(0), 4), " ", ""), Chr(10), "")
'    tmpPhonetic = ybEN & ybUS
'    tmpTrans = ybZEHY
End Sub


Sub searchWordFromidt(tmpWord As String, tmpTrans As String, tmpPhonetic As String)
'        Dim XH As Object
'        Dim s() As String
'        Dim str_tmp As String
'        Dim str_base As String
'        Dim URL As String
'        tmpTrans = ""
'        tmpPhonetic = ""
'        URL = "https://www.dreye.com.cn/dict_new/dict.php?w=" & tmpWord
'        '开启网页
'        Set XH = CreateObject("Msxml2.XMLHTTP") 'Microsoft.XMLHTTP")
' '       On Error Resume Next
'        XH.Open "get", URL, 0 'True
'        XH.Send (Null)
' '       On Error Resume Next
'        While XH.readyState <> 4
'            DoEvents
'        Wend
'        str_base = XH.responsetext
'        'XH.Close
'        Set XH = Nothing
'
'            '取得音标部分
'    ybEN = "英 " & Split(Split(str_base, "KK:")(1), " ")(0)
'    ybUS = " 美 " & Split(Split(str_base, "DJ:")(1), "<")(0)
'
'        str_base1 = Split(Split(str_base, "<ul>")(2), "</ul>")(0)
' 'Debug.Print Mid(str_base1, 1, 3000)
'        v = Split(str_base1, "<span")
'
'            '取得中文含义部分
'        For i = LBound(v) + 1 To UBound(v)
'            hytmp = hytmp & Split(v(i), "</span>")(0)
'        Next i
'        tmpPhonetic = ybEN & " " & ybUS
'tmpTrans = Mid(Replace(Replace(hytmp, "class=""attr"">", Chr(10)), ">", ""), 3)
'
End Sub
