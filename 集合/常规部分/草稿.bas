Attribute VB_Name = "�ݸ�"
'***********��Դ�ϼ�*********"
'https://developer.microsoft.com/zh-cn/windows/downloads/virtual-machines               'win10 ���ɰ汾�����
'https://docs.microsoft.com/zh-cn/cpp/mfc/mfc-desktop-applications?view=vs-2019         'MFC
'https://docs.microsoft.com/zh-cn/windows/win32/index                                   'Win32 API
'https://www.qqxiuzi.cn/                                                                '����

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
'        arr(i, 2) = item.Children.item(1).innertext '����
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
'                If strT = "listenFiles" Then  '��ʺ���������
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

'Sub ����ѷ����() '��Ҫʹ��cookie ,refer
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
'Sub ����()
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
'Set st = fso.OpenTextFile("C:\Users\adobe\Desktop\�쳣�ļ�-����md5���ִ���ֵ.txt", ForReading, False, TristateUseDefault)
'Set objhash = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
'hash = objhash.ComputeHash_1(st)
'st.Close
'objhash.Clear
'Set st = Nothing
'Set objhash = Nothing
'Debug.Print Mid("abc", 1)
'Dim Hash() As Byte
'Hash = StrConv("����˭", vbFromUnicode)
'k = UBound(Hash)
'For i = 1 To k
'Debug.Print i + i + 2 + (Hash(i) > 15)
'Next
''hash = "����˭"
'hash = StrConv("����˭", vbUnicode)

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
'        objShell.ShellExecute "C:\Users\adobe\Desktop\12164431_���λû������ա���Ұ����.epub", "", "", "open", 1
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
Sub StopTim0025er() '��ʱ�� /�����Ͻϸ߾���
'Debug.Print a.MD5Hash("C:\Users\adobe\Desktop\Windows��Դ������.pdf", True, True)
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
Sub WordPr02255otection(ByVal FilePath As String, ByVal Password As String, Optional ByVal cmCode As Byte) '����word�ļ�����
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
''                For i = 1 To strlen2 '�ж�����������Ƿ�Ϊ��������
''                    strx1 = Mid(strx2, i, 1)
''                    If strx1 Like "[һ-��]" Then '���� '�ɵ�������,�ɼ��д��� "��˾����" �ɲ��ɵ������� ��/˾/��/��, ���Կ��Բ�����ո�
''                        k = k + 1
''                        arrstr(p) = strx1
''                        p = p + 1
''                    ElseIf strx1 Like "[a-zA-Z]" Then 'Ӣ����ĸ,����Сд '���д���,������ⵥ�ʵĺ������½�
''                        j = j + 1
''                    ElseIf strx1 Like "[0-9]" Then '���� '���д���,��������û�е��ʲ���Ӱ���
''                        xi = xi + 1
''                    End If
''                Next
'''                If k = 0 And xi = 0 And strlen2 < 3 Then Exit Sub '�����Ǵ�Ӣ����ĸ,����С�ڼ���������Ӧ
'''                If j <> strlen2 And k <> strlen2 And xi <> strlen2 Then
''''                    If SafeArrayGetDim(AbtainKeyWord(strx)) = 0 Then Exit Sub 'û�л�ȡ����Ч��ֵ
'''                    keycount = keycount - 1
'''                    If keycount >= 0 Then Exit Sub
'''                    ReDim arr(keycount)
'''                    arr = ObtainKeyWord(strx)
'''                ElseIf k = 0 And j = 0 And xi = strlen2 And strlen2 >= 2 Then '����
'''                    strx = strx2
'''                End If
''                '-----------------------------------ǰ�ڹؼ��ʷ���
'                For j = 1 To 2
'                    For k = 1 To 4
'                    If InStr(1, arrax(k, 1 * j) & "/" & arrax(k, 2 * j), strx, vbTextCompare) > 0 Then
'                        dic(arrax(k, 1)) = arrax(k, 2)
'                        dica(arrax(k, 1)) = arrax(k, 3)
'                        dicb(arrax(k, 1)) = arrax(k, 4)
'                        dicc(arrax(k, 1)) = arrbx(k, 1) '�����ķ�ʽ�������кܴ�ĵ����ռ�'����һ�Ž��ƴʵı�,������Ӣ�ĵ�ĳЩ�������ͬ������'��ƴд����Ĵ��滻����������
'                        mi = mi + 1
'                        If mi > 50 Then GoTo 100 '�����������������
'                    Else
'                        For i = 1 To strlen2 '�ж�����������Ƿ�Ϊ��������
'                            strx1 = Mid(strx2, i, 1)
'                            If strx1 Like "[һ-��]" Then '���� '�ɵ�������,�ɼ��д��� "��˾����" �ɲ��ɵ������� ��/˾/��/��, ���Կ��Բ�����ո�
'                                k = k + 1
'                                arrstr(p) = strx1
'                                p = p + 1
'                            ElseIf strx1 Like "[a-zA-Z]" Then 'Ӣ����ĸ,����Сд '���д���,������ⵥ�ʵĺ������½�
'                                j = j + 1
'                            ElseIf strx1 Like "[0-9]" Then '���� '���д���,��������û�е��ʲ���Ӱ���
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
'                    For k = 1 To blow 'ע������ɸѡ��Ŀ��֮������ͬ���ַ�,�����³��ֶ��н����bug,����ʹ���ֵ�ķ��������
'                        If InStr(1, arrax(k, j), strx, vbTextCompare) > 0 Then
'                            dic(arrax(k, 1)) = arrax(k, 2)
'                            dica(arrax(k, 1)) = arrax(k, 3)
'                            dicb(arrax(k, 1)) = arrax(k, 4)
'                            dicc(arrax(k, 1)) = arrbx(k, 1) '�����ķ�ʽ�������кܴ�ĵ����ռ�'����һ�Ž��ƴʵı�,������Ӣ�ĵ�ĳЩ�������ͬ������'��ƴд����Ĵ��滻����������
'                            mi = mi + 1
'                            If mi > 50 Then GoTo 100 '�����������������
'                        Else
'                            If InStr(strx, Chr(32)) > 0 Then '��������ݴ��ڿո� -��ģ��������
'                                p = 1
'                                For m = 1 To strlen
'                                    strx1 = Mid(strx, m, 1)
'                                    If strx1 Like "[һ-��]" Then 'ֻ��������ַ�
'                                        arrstr(p) = strx1
'                                        p = p + 1
'                                    End If
'                                Next
'                                If p < 2 Then Exit For '����̫��,���ٽ��м���
'                                xi = 0
'                                For t = 1 To p
'                                    If InStr(1, arrax(k, j), arrstr(t), vbTextCompare) > 0 Then xi = xi + 1
'                                Next
'                                If xi > 2 Then
'                                    dic(arrax(k, 1)) = arrax(k, 2)
'                                    dica(arrax(k, 1)) = arrax(k, 3)
'                                    dicb(arrax(k, 1)) = arrax(k, 4)
'                                    dicc(arrax(k, 1)) = arrbx(k, 1) '�����ķ�ʽ�������кܴ�ĵ����ռ�
'                                    mi = mi + 1
'                                    If mi > 50 Then GoTo 100
'                                End If
'                            Else         'ģ������,�������ո��
'                                p = 0
'                                '-----------------------���strx="a", "abc",����ת����,���Ƚ�Ϊ0,1(ע��)
'                                If Len(strx) \ 2 = Len(StrConv(strx, vbFromUnicode)) Then
'                                '--------�����������ַ�,ע������Ĳ����ù������е�len/lenb��������Ӣ���ַ��Ĳ���,�������ת������ܽ��бȽ�
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
'                                    If strx1 Like "[һ-��]" Then 'ֻ��������ַ�
'                                        p = p + 1
'                                        arrstr(p) = strx1
'                                    End If
'                                Next
'                                If p < 2 Then Exit For '����̫��,���ٽ��м���
'                                xi = 0
'                                For t = 1 To p
'                                    If InStr(1, arrax(k, j), arrstr(t), vbTextCompare) > 0 Then xi = xi + 1
'                                Next
'                                If xi > 2 Then
'                                    dic(arrax(k, 1)) = arrax(k, 2)
'                                    dica(arrax(k, 1)) = arrax(k, 3)
'                                    dicb(arrax(k, 1)) = arrax(k, 4)
'                                    dicc(arrax(k, 1)) = arrbx(k, 1) '�����ķ�ʽ�������кܴ�ĵ����ռ�
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

'If spyx <> docmx Then  '����ģ�鼶�����������������ֵ���ڴ���,���ٷ��ʱ�����Ҫ,ֻ�е��������ݷ����仯�����»�ȡֵ,�ӿ���ʵ��ٶ�
'    spyx = blow '��ʼ��ֵ/�仯�ڽ��и�ֵ
'    If SafeArrayGetDim(arrax) <> 0 Then
'    Erase arrax
'    Erase arrbx                        '�����ݷ����仯��ʱ��,Ĩ���ɵ��������»�ȡ�µ�
'    End If
'    With ThisWorkbook.Sheets("���")
'        arrax = .Range("b6:e" & blow).Value
'        arrbx = .Range("n6:n" & blow).Value '�򿪴���
'    End With
'End If
'ReDim arrx(1)
'arrx(1) = 1
'Erase arrx
'MsgBox SafeArrayGetDim(arrx)
' Debug.Print Len(StrConv("aaa", vbFromUnicode)), Len("����˭"), Len(StrConv("��aa", vbFromUnicode)), Len(StrConv("��aaa", vbFromUnicode)), Len("��aaa")
'Dim arr() As String
''If IsArray(arr) = True Then MsgBox 1
''ReDim arr(1)
'Debug.Print SafeArrayGetDim(aaa)
''Dim myreg As Object, match As Object, matches As Object
'Dim arr() As String, k As Byte, xtext As String, i As Byte
'
'xtext = "whot��ʱ100 am"
'Set myreg = CreateObject("VBScript.RegExp")
'With myreg
'    .Pattern = "[a-zA-Z0-9]{2,}" '��ȡ��������
'    .Global = True
'    .IgnoreCase = True '�����ִ�Сд
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

'[\>]+(.+?)[\(]����[\)]
'�û�����:+(.+?)(\<)
'[https]+://+book.douban.com/+[a-z]*\/+[0-9]*\/

            '-------------------------------------��Ϊ���������ݻ��ж�����Ž��,���Է��ص�һ���,׼ȷ�Ȳ�һ����
'url = "https://book.douban.com/subject/3259440/"
'With CreateObject("MSXML2.XMLHTTP")
'    .Open "GET", url, False
'    .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'    .send
'    Do While .readyState <> 4 '�ȴ����ݵķ���
'        DoEvents
'    Loop
'    strx = .responsetext
'Debug.Print strx
'                End
'End With



' Debug.Print "��ǰ���ڿ�������ĸ߶�Ϊ:" & ActiveWindow.UsableHeight
'
' Debug.Print "��ǰ���ڵĸ߶�Ϊ:" & ActiveWindow.Height
'
' Debug.Print "��ǰ���ڿ�������Ŀ��Ϊ:" & ActiveWindow.UsableWidth
'
' Debug.Print "��ǰ���ڵĿ��Ϊ:" & ActiveWindow.Width
'ThisWorkbook.Application.ScreenUpdating = True
'Dim dicx As New Dictionary
'Dim m As Byte
'm = 1
'dicx(m) = ""
'Debug.Print dicx.Keys(0)
'Set dicx = Nothing
''Dim fd As Folder
'
'Set fd = fso.GetFolder("C:\Users\Administrator\Documents\�Զ��� Office ģ��")
'
'Set fd = Nothing

 End Sub






'���ݵ�����
'byte ����: - , ռ��: 1 ���ݷ�Χ: 0-255
'long ����: & , ռ��: 4,���ݷ�Χ -2147483648 ~ 2147483647
'integer ����: %, ռ��: 2 ,���ݷ�Χ: -32768 ~ 32767
'Private Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)
''Private Declare Function hash Lib "ntdll.dll" Alias "RtlComputeCrc32" (ByVal start As Long, ByVal data As Long, ByVal size As Long) As Long
'Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'ElseIf fso.GetDriveName(strx) = Environ("SYSTEMDRIVE") Then '��ֹ������ϵͳ��һ���ļ�������/ֻ�����ļ��з���document��
'                c = UBound(Split(strx, "\"))
'                If c = 1 Then
'                    .Label57.Caption = "��������ϵͳ������"
'                    .TextBox11.text = ""
'                    .TextBox11.SetFocus
'                ElseIf c = 2 Then
'                    If strx <> Environ("UserProfile") Then
'                    .Label57.Caption = "�������ڷ��û��ļ���"
'                    .TextBox11.text = ""
'                    .TextBox11.SetFocus
'                    End If
'                ElseIf c > 2 Then
'                    strx2 = Environ("UserProfile") & "\Documents"
'                    If InStr(strx, strx2) = 0 Then
'                        .Label57.Caption = "ֻ����������documents�ļ�����"
'                        .TextBox11.text = ""
'                        .TextBox11.SetFocus
'                    End If
'i = FileStatus(strx3)
'            Select Case i
'                Case 1: strx5 = "Ŀ¼������"
'                Case 3: strx5 = "�ļ�������"
'                Case 6: strx5 = "�ļ����ڴ򿪵�״̬"
'                Case 7: strx5 = "�쳣"
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
'Dim strdsk As String, strdlw As String, strdcm As String, struserfile As String '�ĵ�,����,����,�û��ļ���
'If tracenum = 1 Then '���ļ�Ϊϵͳ���ļ�,ֻ�����������λ�õ��ļ�desktop,downloads,documents '
'                        struserfile = Environ("UserProfile") '�û��ļ���
'                        strdsk = struserfile & "\Desktop"
'                        strdlw = struserfile & "Downloads"
'                        strdcm = struserfile & "\Documents"
'                        If c = 3 Then
'                            If strp <> strdsk And strp <> strdcm And strp <> strdlw Then GoTo 1001 '���������ļ���
'                            ElseIf c > 3 Then
'                                If InStr(fd.Path, strdsk) = 0 And InStr(fd.Path, strdlw) = 0 And InStr(fd.Path, strdcm) = 0 Then GoTo 1001
'                            ElseIf c < 3 Then
'                                If strp <> Environ("UserProfile") Then GoTo 1001 '����ֻ��������û��ļ���
'                        End If
'                    End If

'Function ZipExtract(ByVal filepath As String, ByVal folderpath As String, Optional ByVal cmcode As Byte) '��ѹ�ļ�
'    Dim exepath As String, strx As String, strx1 As String, strx2 As String, strx3 As String
'    Dim wsh As Object
'
'    exepath = "C:\Program Files\7-Zip\7z.exe" '�����װ������λ��
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
'        strx2 = " && del /s " & filepath 'ִ����ɺ��Զ�ɾ����Դ�ļ� 'del /s cmd ɾ���ļ�����
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
'strx3 = .TextBox13.Text                        '������Ϣ�Ĳ�ͬ�����Զ���ȡ�����ؼ���-��Ҫ�޸�
'strlen1 = Len(Trim(strx1))
'strlen3 = Len(Trim(strx3))
'If strlen1 = 0 And strlen3 = 0 Then Exit Sub '1
'If strlen1 = 0 And strlen3 > 0 Then             '2
''    keyworda = Left$(strx3, Len(strx3) - Len(Split(strx3, ".")(UBound(Split(strx3, ".")))) - 1)
'    keyworda = strx3
'ElseIf strlen1 > 0 And strlen3 = 0 Then        '3
'       If strx1 Like "HLA*&*" Then                     '3.1
'          keyworda = Mid$(strx1, Len(Split(strx1, "&")(0)) + 3, Len(strx1) - Len(Split(strx1, "&")(0)) - 2 - Len(Split(strx1, ".")(UBound(Split(strx1, ".")))) - 1) '���������ע�⵽splitʵ����Ҳ��һ�����飨�������ɵķ�����β�ķָ��飩
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
'strx3 = .TextBox13.Text                        '������Ϣ�Ĳ�ͬ�����Զ���ȡ�����ؼ���-��Ҫ�޸�
'strlen1 = Len(Trim(strx1))
'strlen3 = Len(Trim(strx3))
'If strlen1 = 0 And strlen3 = 0 Then Exit Sub '1
'If strlen1 = 0 And strlen3 > 0 Then             '2
''    keyworda = Left$(strx3, Len(strx3) - Len(Split(strx3, ".")(UBound(Split(strx3, ".")))) - 1)
'    keyworda = strx3
'ElseIf strlen1 > 0 And strlen3 = 0 Then        '3
'       If strx1 Like "HLA*&*" Then                     '3.1
'          keyworda = Mid$(strx1, Len(Split(strx1, "&")(0)) + 3, Len(strx1) - Len(Split(strx1, "&")(0)) - 2 - Len(Split(strx1, ".")(UBound(Split(strx1, ".")))) - 1) '���������ע�⵽splitʵ����Ҳ��һ�����飨�������ɵķ�����β�ķָ��飩
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
'    If SearchFile("WinRAR.exe") <> "û���ҵ�ƥ���" Then '����WinRAR.exe��·��
'        Shell SearchFile("WinRAR.exe") & "   a   c:\ѹ������ļ�.rar   c:\��ѹ�����ļ����ļ���", vbHide
'    Else
'        MsgBox "û���ҵ�RARѹ�����", vbOKOnly, "��ʾ"
'    End If
'End Sub
    
'        With ThisWorkbook
'            blow = .Sheets("���").[d65536].End(xlUp).Row
'            If blow <> docmx Then
'                docmx = blow
'                Call CwUpdate '���´��ڵ�����
'                Call Choicex '����ɸѡ��������
'            End If
'            With .Sheets("������") '��ȡ��ʼֵ     '���´����������4��λ�õ�����
'                recentfilex = CStr(.Range("w26").Value)
'                prfilex = .Range("i26").Value
'                addfilecx = .[e65536].End(xlUp).Row
'
'                If recentfilex <> Recentfile Then '�ڱ����ļ���Ĵ���
'                    Recentfile = recentfilex
'                    Call RecentUpdate
'                    With Me
'                        strx1 = .Label1.Caption
'                        strx2 = .Label32.Caption
'                        If Len(strx1) > 0 And strx1 = .ListBox1.List(0, 0) Then '���±༭�����е�����
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
''rcs.Open "C:\Users\adobe\Desktop\�쳣�ļ�-����md5���ִ���ֵ.txt",
'
'connx.Open "C:\Users\adobe\Desktop\�쳣�ļ�-����md5���ִ���ֵ.txt"
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
'strx = "����"
'filepath = ThisWorkbook.Path & "\test.xlsx"
'
'conn.Open "Provider = Microsoft.ACE.OLEDB.12.0;Data Source=" & filepath & ";extended properties=""excel 12.0;HDR=YES""" '�����ݴ洢�ļ�
'sql = "Insert into [����$] (���,�ļ�,ԭ��) Values ('" & unicode & "', '" & filen & "', '" & mfilen & "')"
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
'Cells(5, "ae") = "�ļ����쳣�ַ�"
'Cells(5, "af") = "�ļ���λ���쳣�ַ�"
'For i = 0 To 31
'Cells(3, i + 2) = i
'Next
'Application.DisplayFormulaBar = True
'For Each drx In fso.Drives
''Debug.Print Environ("SYSTEMDRIVE")
'Debug.Print drx.Path
'Shell "notepad.exe " & "C:\Users\Lian\Desktop\contents.txt", vbNormalFocus '���ļ�
 'Shell "cmd /c tree D:\L-temp /f >C:\Users\Lian\Desktop\contents.txt", vbHide
Dim strfolder As String
Dim strx As String, strx1 As String
Dim wsh As Object

With Application.FileDialog(msoFileDialogFolderPicker) '�ļ���ѡ�񴰿�(�ļ����Դ�ֻ��ѡһ��)
    .Show
    If .SelectedItems.Count = 0 Then Exit Sub 'δѡ���ļ������˳�sub
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
'Private Sub CommandButton22_Click() 'ȫ���ļ������ݸ���-��ť�ѱ�ɾ��
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
'                        rngad.Offset(0, 2) = Now '���������ļ��е��޸�ʱ��
'                        filec = rngad.Offset(0, 4).Value '�����ļ�������
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
'Set rng = Nothing  'rngΪ���̼�����,��ִ���������,�ͷ��ڴ�
'Call Warning(3)
'Exit Sub
'End If
'
'Me.Label55.Visible = False
'Me.Label23.Caption = rng.Offset(0, 1) '�ļ���
'Me.Label24.Caption = rng.Offset(0, 2) '�ļ�����
'Me.Label25.Caption = rng.Offset(0, 3) '�ļ�·��
'Me.Label26.Caption = rng.Offset(0, 4) '�ļ�λ��
'
'Me.Label27.Caption = rng.Offset(0, 6) '�ļ���С
'Me.Label28.Caption = rng.Offset(0, 7) '����ʱ��
'Me.Label29.Caption = rng.Offset(0, 0) 'ͳһ����
'Me.Label30.Caption = rng.Offset(0, 9) '�ļ����
'
'Me.Label31.Caption = rng.Offset(0, 10) '�����ʱ��
'Me.Label32.Caption = rng.Offset(0, 11) '�򿪴���
'Me.Label33.Caption = rng.Offset(0, 8) '��ʶ����
'
'Me.TextBox5.text = rng.Offset(0, 18) '��ǩ1
'Me.TextBox6.text = rng.Offset(0, 19) '��ǩ2
'
'Me.TextBox4.text = rng.Offset(0, 13) '����
'
'Me.ComboBox3.text = rng.Offset(0, 14) 'pdf������
'Me.ComboBox4.text = rng.Offset(0, 15) '�ı�����
'Me.ComboBox5.text = rng.Offset(0, 16) '��������
'
'Me.ComboBox2.text = rng.Offset(0, 17) '�Ƽ�����
'
'Me.Label69.Caption = rng.Offset(0, 23) '��������
'
'Me.Label71.Caption = rng.Offset(0, 20) 'MD5
'
'If rng.Offset(0, 12) = "" Then
'Call Aton
'Else
'Me.TextBox3.text = rng.Offset(0, 12) '���ļ���
'End If
'
'
'End With
'
'Call Text2a '��ȡ�ļ���ժҪ��Ϣ
'Call Disabledit
'End If
'If .ListBox1.ListCount = 0 Then
'.ListBox1.AddItem
'GoTo 100
'
'ElseIf .ListBox1.ListCount = 7 Then          '���б��Ѿ�����7����ʱ��,������д�����е�����
'For k = 6 To 1 Step -1
'.ListBox1.List(k, 0) = .ListBox1.List(k - 1, 0)
'.ListBox1.List(k, 1) = .ListBox1.List(k - 1, 1)
'.ListBox1.List(k, 2) = .ListBox1.List(k - 1, 2)
'Next
'
'GoTo 100
'
'ElseIf .ListBox1.ListCount > 0 And .ListBox1.ListCount < 7 Then '������С��7��ʱ��,���������ƶ�
'.ListBox1.AddItem
'n = .ListBox1.ListCount
'For k = n To 1 Step -1
'.ListBox1.List(k, 0) = .ListBox1.List(k - 1, 0)
'.ListBox1.List(k, 1) = .ListBox1.List(k - 1, 1)
'.ListBox1.List(k, 2) = .ListBox1.List(k - 1, 2)
'Next
'
'100                                                                 '���б�Ϊ��,ֱ���ڵ�һ��д������
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
'    If .ListView1.ListItems(itemf.Index).SubItems(4) = "" Then .ListView1.ListItems(itemf.Index).SubItems(4) = 0 '��ֵ��"0"�����������
'    .ListView1.ListItems(itemf.Index).SubItems(4) = .ListView1.ListItems(itemf.Index).SubItems(4) + 1 '�򿪴���+1
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
'ElseIf .ListBox2.ListCount = 7 Then                               '���б��Ѿ�����7����ʱ��,������д�����е�����
'
'For k = 6 To 1 Step -1
'.ListBox2.List(k, 0) = .ListBox2.List(k - 1, 0)
'.ListBox2.List(k, 1) = .ListBox2.List(k - 1, 1)
'.ListBox2.List(k, 2) = .ListBox2.List(k - 1, 2)
'Next
'
'GoTo 100
'
'ElseIf .ListBox2.ListCount > 0 And .ListBox2.ListCount < 7 Then '������С��7��ʱ��,���������ƶ�
'
'n = .ListBox2.ListCount
'.ListBox2.AddItem
'For k = n To 1 Step -1
'.ListBox2.List(k, 0) = .ListBox2.List(k - 1, 0)
'.ListBox2.List(k, 1) = .ListBox2.List(k - 1, 1)
'.ListBox2.List(k, 2) = .ListBox2.List(k - 1, 2)
'Next
'
'100                                                              '���б�Ϊ��,ֱ���ڵ�һ��д������(ע�������д��, ��������elseif��ȥ��)
'.ListBox2.List(0, 0) = .Label29.Caption
'.ListBox2.List(0, 1) = .Label23.Caption
'.ListBox2.List(0, 2) = Now
'
'End If













'       .Range("k" & .[c65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrmd5)  '�ļ�md5
'        .Range("ab" & .[c65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrcode) '�쳣�ַ����
'        .Range("ac" & .[c65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrcm)   '��ע
'        .Range("ae" & .[c65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrfnansi)
'        .Range("ab" & .[c65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrfpansi)
'        .Range("c" & .[c65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrbase) '�ļ���
'        .Range("d" & .[d65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrextension) '�ļ���չ��
'        .Range("e" & .[e65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrfiles) '�ļ�·��
'        .Range("f" & .[f65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrparent) '�ļ�����λ��
'        .Range("g" & .[g65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrsizeb) '�ļ���ʼ��С
'        .Range("h" & .[h65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arredit) '�ļ��޸�ʱ��
'        .Range("i" & .[i65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrsize) '�ļ���С
'        .Range("j" & .[j65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrdate) '�ļ�����ʱ��
'        .Range("x" & elow & ":" & "x" & flc + elow - 1) = Now '���Ŀ¼��ʱ��
'        .Range("ad" & elow & ":" & "ad" & flc + elow - 1) = 1 '��ע�ļ�����Դ(��ͨ������ļ��еķ�ʽ��ӽ�����)
'        .Range("b" & elow - 1).AutoFill Destination:=.Range("b" & elow - 1 & ":" & "b" & flc + elow - 1), Type:=xlFillDefault '���ͳһ����
'        .Range("b" & elow & ":" & "b" & flc + elow - 1).Interior.Pattern = xlPatternNone
'        .Range("b" & elow & ":" & "b" & flc + elow - 1).Font.ThemeColor = xlThemeColorLight1
'
'i = .ListCount
'        If i = 0 Then
'        .AddItem
'        GoTo 100
'
'        ElseIf i = 7 Then          '���б��Ѿ�����7����ʱ��,������д�����е�����
'        For k = 6 To 1 Step -1
'        .List(k, 0) = .List(k - 1, 0)
'        .List(k, 1) = .List(k - 1, 1)
'        .List(k, 2) = .List(k - 1, 2)
'        Next
'
'        GoTo 100
'
'        ElseIf i > 0 And i < 7 Then '������С��7��ʱ��,���������ƶ�
'        .AddItem
'        For k = i To 1 Step -1
'        .List(k, 0) = .List(k - 1, 0)
'        .List(k, 1) = .List(k - 1, 1)
'        .List(k, 2) = .List(k - 1, 2)
'        Next
'
'100                                                                         '���б�Ϊ��,ֱ���ڵ�һ��д������
'        .List(0, 0) = strx
'        .List(0, 1) = strx1
'        .List(0, 2) = recentfile 'ͳһʹ�����ʱ��,ȷ�����ʹ��ڵ�������ȫһ��
'        End If

'    With Me.ListBox1 '����Ķ�
'        m = .ListCount
'        If m = 0 Then
'            .AddItem
'            GoTo 100
'        ElseIf m = 7 Then          '���б��Ѿ�����7����ʱ��,������д�����е�����
'            For k = 6 To 1 Step -1
'            .List(k, 0) = .List(k - 1, 0)
'            .List(k, 1) = .List(k - 1, 1)
'            .List(k, 2) = .List(k - 1, 2)
'            Next
'            GoTo 100
'        ElseIf m > 0 And m < 7 Then '������С��7��ʱ��,���������ƶ�
'            .AddItem
'            For k = m To 1 Step -1
'                .List(k, 0) = .List(k - 1, 0)
'                .List(k, 1) = .List(k - 1, 1)
'                .List(k, 2) = .List(k - 1, 2)
'            Next
'100                                                                                 '���б�Ϊ��,ֱ���ڵ�һ��д������
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
'ElseIf .ListCount = 7 Then          '���б��Ѿ�����7����ʱ��,������д�����е����� '�����޸�
'For k = 6 To 1 Step -1
'.List(k, 0) = .List(k - 1, 0)
'.List(k, 1) = .List(k - 1, 1)
'.List(k, 2) = .List(k - 1, 2)
'Next
'
'GoTo 100
'
'ElseIf .ListCount > 0 And .ListCount < 7 Then '������С��7��ʱ��,���������ƶ�
'.AddItem
'n = .ListCount
'For k = n To 1 Step -1
'.List(k, 0) = .List(k - 1, 0)
'.List(k, 1) = .List(k - 1, 1)
'.List(k, 2) = .List(k - 1, 2)
'Next
'
'100                                                                 '���б�Ϊ��,ֱ���ڵ�һ��д������
'.List(0, 0) = strx
'.List(0, 1) = Me.ListView1.SelectedItem.ListSubItems(1).text
'.List(0, 2) = recentfile 'ͳһʹ�����ʱ��,ȷ�����ʹ��ڵ�������ȫһ��
'End If
'
'End With

'If f = 4 And elow > 100 And ifilec > 10 Then '�ļ��д�����Ŀ¼-�����ļ���ʱ����Ŀ¼�е��ļ��Ƿ񻹴���'���ⲻ��Ҫ��ȫ�����
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
'    With ThisWorkbook.Sheets("���")
'    For J = itemp To 1 Step -1 'ɾ�����õ���ɾ���ķ�ʽ
'        If arrfilefd(J, 1) = strxp Then
'            If fso.FileExists(arrfilep(J, 1)) = False Then
'                .Rows(J + 5).Delete Shift:=xlShiftUp '���Ŀ¼���ļ������ھ�ɾ��
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
'If itemp < 300 And ifilec > 10 Then '���������ݷǳ��ٵ�ʱ��'����ȫ���ļ��'�������ִ���˼�����Ͳ��ٽ��м��,��ӵ��ļ��㹻��
'                ReDim arrfilenx(1 To itemp) '��ʱ����
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
'                        ClearAll (1) '���Ŀ¼����
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
''If Len(strx) = 0 Or fso.FileExists(strx) = False Then .Label57.Caption = "�ļ�����": Exit Sub
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
'Set flt = fl.OpenAsTextStream(ForWriting, TristateUseDefault) '������forwriting����,�ᵼ���ļ�������
'flt.Write "Nothing"
'flt.Close
'Exit Sub
'100
'MsgBox Err.Number
'
'  Open address For Binary Access Read Write Lock Read Write As #1 'ע��open��֧�ַ�ansi�ַ�'�ж��ļ��Ƿ��ڴ򿪵�״̬,�����txt�ļ���ֱ������
'                  Close #1
'  If Len(.Range("ab4").Value) = 0 And .Range("ab9") = 1 Then '1��ʾIE����
'            exepath = "C:\Program Files\Internet Explorer\iexplore.exe "
'        ElseIf .Range("ab4") <> "" Then
'            exepath = .Range("ab4")                                                                 'Environ("SYSTEMDRIVE")��ʾϵͳ���ڵ��̷�
'            If fso.FileExists(Left$(exepath, Len(exepath) - 1)) = False And .Range("ab9") = 1 Then exepath = Environ("SYSTEMDRIVE") & "\Program Files\Internet Explorer\iexplore.exe " '����ǰ������Ĵ���,�����������������ΪIE
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
'If strx Like "[һ-��]" Then MsgBox 1
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
''         Debug.Print Format(Now, "yyyy��mm��dd��")
'         Debug.Print Format(Now, "yyyy/mm/dd/h:mm:ss")
''         Debug.Print Format(Now, "d-mmm-yy") 'Ӣ���·�
''         Debug.Print Format(Now, "d-mmmm-yy") 'Ӣ���·�
''         Debug.Print Format(Now, "aaaa") '��������
''         Debug.Print Format(Now, "ddd") 'Ӣ������ǰ������ĸ
''         Debug.Print Format(Now, "dddd") 'Ӣ������������ʾ
'   End Sub
'�������
'Dim elow As Integer
'With ThisWorkbook.Sheets("���")
'elow = .[e65536].End(xlUp).Row
'If elow < 6 Then
'.Label1.Caption = "������"
'Exit Sub
'ElseIf elow > 100 Then
'UserForm6.Show 0
'Call checkfile
'End If
'Me.Label47.Caption = .Range("p37").Value '�ļ�����
'Me.Label48.Caption = .Range("p38").Value '�����ļ���С
'Me.Label49.Caption = .Range("p40").Value 'pdf
'Me.Label50.Caption = .Range("s40").Value 'EPUB
'Me.Label51.Caption = .Range("p42").Value '����
'Me.Label52.Caption = .Range("p41").Value 'PPT
'Me.Label53.Caption = .Range("v41").Value 'Word
'Me.Label54.Caption = .Range("s41").Value 'Excel
'if strx="��������" or len(strx)= then exit sub
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
'ThisWorkbook.ChangeFileAccess xlReadOnly '����ļ�����
'Kill ThisWorkbook.FullName 'ɾ���ļ�
'MsgBox "��ʼ���ɹ��������´��ļ�"
'Call openfilelocation(userpath)          '�����ļ����ڵ��ļ���λ��
'ThisWorkbook.Close False
'            If ThisWorkbook.Sheets("��ҳ").Range("d3").Value = 1 Or ThisWorkbook.Sheets("���").Range("d3").Value = 0 Then
'For i = 3 To 35
'Cells(5, i) = Cells(5, i) & i - 2
'Next
'Shell ("PowerShell_ISE "), vbNormalFocus
'strCommand = "Powershell.exe -ExecutionPolicy ByPass ""C:\Users\****\Documents\lb\file.ps1"" " & FilePath
End Sub

'    sql = "select * from [" & TableName & "$] where ͳһ����='" & str1 & "'"                                          '��ѯ����
'    Set rs = New ADODB.Recordset
'    rs.Open sql, conn, adOpenKeyset, adLockOptimistic
'    If rs.BOF And rs.EOF Then '�����ж������ҵ�����
 '    Else
''    MsgBox "������ͬ���"
'    End If
    
'    rs.Close
'    Set rs = Nothing

'With Me
'If Len(.TextBox11.text) = 0 And Len(.TextBox12.text) = 0 And Len(.TextBox22.text) = 0 Then '���߶�Ϊ�յ�ʱ��
'Call warning(2)
'Me.TextBox11.SetFocus
'Exit Sub
'End If
'
'If Len(.TextBox11.text) > 0 And fso.FileExists(.TextBox11.text) = True Then
'ThisWorkbook.Sheets("temp").Range("ab4") = Trim(.TextBox11.text) & Chr(32) '����� 'chr(32)��ʾ�ո����
'ThisWorkbook.Sheets("temp").Range("ac4") = 1 '��ǳ����Ѿ�����,���ڼ��chrome������Ƿ����
'End If
'If InStr(.TextBox12.text, "exe") > 0 And fso.FileExists(.TextBox12.text) = True Then ThisWorkbook.Sheets("temp").Range("ab6") = Trim(.TextBox12.text) & Chr(32) '��ͼ
'If Len(.TextBox22.text) > 0 And fso.FileExists(.TextBox22.text) = True Then ThisWorkbook.Sheets("temp").Range("ab5") = Trim(.TextBox22.text) & Chr(32) 'Pdf�༭
'End With

'������̬ʱ��

'                .Worksheets(4).Name = "ɾ������"
                '.Worksheets(4).Range("a1:u1") =

'Array("ͳһ����", "�ļ���", "�ļ�����", "�ļ�·��", "�ļ�����λ��", "�ļ���ʼ��С", "�ļ���С", "�ļ�����ʱ��", "��ʶ���", "�ļ����", "�����ʱ��", "�ۼƴ򿪴���", "���ļ���", "����", "PDF������", "�ı�����", "��������", "�Ƽ�ָ��", "��ǩ1", "��ǩ2", "���ʱ��")
'    ThisWorkbook.VBProject.VBComponents.Item(1).CodeModule.ReplaceLine 109, "lastpath=" & """" & lastpath2 & """"
'timest = False '���ʱ��ؼ����ڳ�������в�����Ӱ��ǳ���,�����׵���ȫ��ı���
'Call atclock
'        ThisWorkbook.VBProject.VBComponents.Item(1).CodeModule.ReplaceLine 109, "lastpath=" & """" & lastpath1 & """"
        
        'MkDir (userpath) '�����ļ�
        'SetAttr (userpath), vbHidden '�����ļ��е�����Ϊ���� '��֧�ַ�ansi�ַ�
'Sub comboxclick() '������ɸѡ����ʱcombox�������ʱ��ʱʹ��(�޸�)
'Me.MultiPage1.Value = 0       'ѡ�����ǩҳ
'Me.ListView2.ListItems.Clear  '���ҳ���ϵ�����
'Me.ListView2.Visible = True   '�ɼ�
'
'End Sub


'If Len(.TextBox14.text) > 0 Then
'   If 2 * Len(Me.TextBox14.text) = LenB(StrConv(Me.TextBox14.text)) Then '��Ҫ����(��������������ַ�ʱ,��������������)
'    .TextBox13.text = UDF_Translate(Me.TextBox1.text)
'    .TextBox14.text = Me.TextBox1.text
'   Else
'    .Label63.Caption = ""
'    .TextBox13.text = "" '���
'    Call WriteVocabulary
'   End If
'End If
'End With



Sub �ϲ����()
Dim timea As String
timea = Now
timea = Format(timea, "ddd")
Debug.Print timea
' MsgBox "!���������쳣,�����޷���������", vbCritical
' End Sub
'Dim fd As Folder
'Set fd = fso.GetFolder("D:\�ļ���С")
'With fd
'Debug.Print
'End With
End Sub
'If rng.Offset(0, 2) Like "xl*" Then '����ļ���������excel,��ô�жϴ򿪵��ļ��Ƿ�����
'   For i = 1 To Workbooks.Count
'   If rng.Offset(0, 1).Value = Workbooks(i).Name Then FileExist = False
'   Next
'End If
'If rng Is Nothing Then '�ļ��Ƿ����
'FileExist = False
'Else
'FileExist = True
'End If
'Public timest As Boolean
'Public Function atclock() '��̬ʱ��
'If timest = True Then
'UserForm1.Label1.Caption = Format(Now, "yyyy-mm-dd HH:MM:SS")
'Sleep 25
'ThisWorkbook.Application.OnTime Now + TimeValue("00:00:01"), "atclock"
'End If
'End Function


'Call deleback(Cells(sx(p), 2), Cells(sx(p), 3), Cells(sx(p), 4), Cells(sx(p), 5), Cells(sx(p), 6), Cells(sx(p), 7), Cells(sx(p), 8), Cells(sx(p), 9), Cells(sx(p), 10), Cells(sx(p), 11), Cells(sx(p), 12), Cells(sx(p), 13), Cells(sx(p), 14), Cells(sx(p), 15), Cells(sx(p), 16), Cells(sx(p), 17), Cells(sx(p), 18), Cells(sx(p), 19), Cells(sx(p), 20), Cells(sx(p), 21)) 'ִ�б���
'j = 1
'ReDim arrcolumn(1 To k)

'Rows(sx(p)).Delete Shift:=xlShiftUp
'Call delefileover(arrcolumn(p)) '�ƺ�

'If yesno = vbYes Then
'    For p = 1 To k
''    For Each slc In Selection.Rows
'       filedyn = False
''        trow = slc.Row
''        arrcolumn(j) = .Range("f" & trow).Value '��ʱ�洢λ����Ϣ '��Ҫ����
''        j = j + 1
'        tfile = .Range("e" & trow).Value
'        If fso.FileExists(tfile) = True Then
'            If Range("ab" & trow) = "ERC" Then
'                fso.DeleteFile (tfile)
'            Else
'                DeleteFiles (tfile)
'            End If
'            filedyn = True '�ļ�ȷʵ�Ǳ�ɾ��(ֻ�е��ļ���ɾ����ʱ��Ż�����ļ��е��޸�ʱ��ı仯)
'        End If
'    Next
'End If

'Function ReplacePunctuation(ByVal strText As String, Optional ByVal IsCN As Boolean = False) As String '���(33-126)��ȫ��(65281-65374) '[\-,\/,\|,\$,\+,\%,\&,\',\(,\),\*,\x20-\x2f,\x3a-\x40,\x5b-\x60,\x7b-\x7e,\x80-\xff,\u3000-\u3002,\u300a,\u300b,\u300e-\u3011,\u2014,\u2018,\u2019,\u201c,\u201d,\u2026,\u203b,\u25ce,\uff01-\uff5e,\uffe5]
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
'    Union(.Range("b6:e" & .[b65536].End(xlUp).Row), .Range("m6:m" & .[b65536].End(xlUp).Row)).Select '������Ҫ����(����unionƴ�ӵ�����,���ֳ���������)
'    Set rg = Selection
'    ReDim arr(1 To rg.Areas.Count)
'    For i = 1 To UBound(arr)
'        arr(i) = rg.Areas(i)
'    Next

'If f > 0 Then
'flmd = flx.DateLastModified 'ʱ��,�Ƚ�Ŀ¼�ļ��е��޸�ʱ����ļ����޸�ʱ��
'If flmd < fdmd Then GoTo 20
'End If
'Dim fdmd As Date, flmd As Date '�ļ��޸�ʱ��,  '�ļ��޸�ʱ�� �ļ����в��ᵼ���ļ�createtime�����仯 'ֻ�����ڼ���������ļ���
'Dim t1 As Date
'Dim t2 As Date
'
't1 = Cells(4, 5).Value
't2 = Cells(5, 5).Value 'ʱ��Ƚ�

'Sub lop()
'Dim t1 As Date
'Dim t2 As Date
't1 = Cells(1, 1).Value
't2 = Cells(2, 1).Value
'Cells(3, 1) = DateDiff("d", t1, t2) 'datediffʱ��������,"d"��ʾ���������day
'End Sub
'
'If t1 > t2 Then MsgBox 1

'With ThisWorkbook.Sheets("temp")                   '���ʿ���ڴ򿪵�Excel�����
'sTest = .Cells(randomx, "ah") '�洢���� '��-Ӣ��
'arrtemp3(listnum) = .Cells(randomx, "ai") '����,cellsҲ֧��range��ģʽ
'arrtemp1(listnum) = sTest
'Me.Label90.Caption = .Cells(randomx, "ai")
'End With

'Alastrow = ThisWorkbook.Sheets("temp").[ah65536].End(xlUp).Row '��ʾ���� '�����������������

'UserForm3.Show 0
'UserForm3.MultiPage1.Value = 1
'InStr(filePath, ChrW(8226)) > 0 Or InStr(filePath, ChrW(12539)) > 0 Or '���ڼ��غ��ַ�/���ĵļ�����Լ��������쳣�ַ�,���й���(����ʵ�ʽ��е���,�������ַ��Ǹ��˵��ļ��о������ֵ�)
'ThisWorkbook.VBProject.VBComponents.Item(3).CodeModule.InsertLines 125, "exit sub"
'ThisWorkbook.VBProject.VBComponents.Item(3).CodeModule.DeleteLines 125
'With Sheets("������")
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
'With ThisWorkbook.Sheets("���")
''.Range("y5") = "����"
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
'Public Flagpause As Boolean, FlagStop As Boolean, Flagnext As Boolean '����ִ�еı���-����ѵ��
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
'''MsgBox "��������ϵͳ�����ƺͰ汾Ϊ:" & Application.OperatingSystem Environ("OS")
''i = CInt(Application.Version)
'
'Dim wsh As Object
'Dim wExec As Object
'
'Set wsh = CreateObject("WScript.Shell")
'Set wExec = wsh.Exec("powershell Get-Host | Select-Object Version") '��ȡpowershell�汾��
'Result = wExec.StdOut.ReadAll
'
'
'End Sub

Sub ������ť��λ��()

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

Sub ��ʾλ��()
'
'elow = Sheets("���").[c65536].End(xlUp).Row + 1
'
'MsgBox elow

'ThisWorkbook.Sheets("���").Range("d3").ClearContents
'UserForm3.MultiPage1.Pages(6).visibale = True
'UserForm3.Show
'For i = 0 To 5
'MsgBox UserForm3.MultiPage1.Pages(i).Caption
'Next

'MsgBox UserForm3.Controls.Count

'Sheets("���").Range("f5:f" & Sheets("���").[f65536].End(xlUp).Row).AutoFilter

'Range("e39") = "c:\users\lian\download\"

'Sheets("���").Range("f5:f" & Sheets("���").[f65536].End(xlUp).Row).AutoFilter Field:=1, Criteria1:="c:\users\lian\download\"

'Dim arr()
'With ThisWorkbook.Sheets("���")
'
'arr = .Range("d6:d" & [d65536].End(xlUp).Row).Value
'
'End With
'ThisWorkbook.Sheets("���").Range("v4") = "����"
Application.ScreenUpdating = True
End Sub

Sub ��ʾѡ��������()

'MsgBox Selection.Column

'If fso.DriveExists("d") = True Then
'
'MsgBox fso.GetDrive("c").DriveType
'End If

'Windows(ThisWorkbook.Name).Visible = True
'ThisWorkbook.Sheets("���").Visible = False

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

With ThisWorkbook.Sheets("���")
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

''����Ķ�
'If .Sheets("������").Range("u27") <> "" Then '����ӿհ׵�ֵ����
'
'
'    For j = 33 To 28 Step -1
'        If .Sheets("������").Range("p" & j) <> "" Then Exit For
'    Next
'
'
'arrrc = .Sheets("������").Range("p27:w" & j).Value
'
'm = UBound(arrrc)
'
''Me.ListBox1.AddItem
'
'End If
'If .Sheets("������").Range("e37") <> "" Then '����ӿհ׵�ֵ����
'
'For j = 38 To 100
'If .Sheets("������").Range("e" & j) = "" Then Exit For
'Next
'
'arral = .Sheets("������").Range("e37:i" & j - 1).Value
'End If
'
'End With
'With ThisWorkbook.Sheets("���")
'Dim arrb()
'arrb = .Range("c6" & ":" & "c" & .[b65536].End(xlUp).Row).Value
'End With

'arra = Array(".xlsm", ".docx", ".pptx", ".txt", ".accdb", ".mobi", ".epub")
'MsgBox UBound(arra)

'MsgBox fso.GetParentFolderName(Range("f10"))

'ThisWorkbook.Sheets("���").Range("w5") = "���ʱ��"


End Sub

Sub module() '�г�ģ���λ��

'For i = 1 To 25
'
'ThisWorkbook.Sheets("���").Range("j" & 6 + i) = ThisWorkbook.VBProject.VBComponents.Item(i).Name
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

With ThisWorkbook.Sheets("���")
If .[e65536].End(xlUp).Row > 6 Then
arr = .Range("e6:e" & .[e65536].End(xlUp).Row).Value
Else
ReDim arr(1 To .[e65536].End(xlUp).Row - 5)
For p = 1 To .[e65536].End(xlUp).Row
arr(p) = .Range("e" & p + 5).Value
Next
End If

arra = Array(".xlsm", ".docx", ".pptx", ".txt", ".accdb", ".mobi", ".epub") 'excel,word,ppt,txt,access,������
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
'With ThisWorkbook.Sheets("���")
'Set rngmd = .Range("k6:k" & .[k65536].End(xlUp).Row).Find("filemd5") '����Ƿ��ļ��Ѵ���
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
'Set fd = fso.GetFolder(filePath)          '��fdָ��·������
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

Set fd = fso.GetFolder("D:\����\you are")

With ThisWorkbook.Sheets("Ŀ¼")
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
'    bc = bc - 1 '�ļ��в㼶����
Loop Until fd.IsRootFolder

End With


'If fd.IsRootFolder Then MsgBox 1

'Cells(4, 3).Resize([b65536].End(xlUp).Row, 1).Select
110
End Sub

'        For i = 1 To 109         '����������ַ���Ҫ����ʵ�������е���
'        k = 160 + i
'        If AscB(CharUpper(ChrW(k))) = k Then GoTo 1000 '����ʹ��charupper��chrw���ɵ��ַ�ת����Ϊansi���룬Ȼ��Ա�ascbת���µ�ansi�����ַ����ɵĴ���ֵ��������ߵ���ֵ��ͬ����ζ������vba���Դ�����ַ�
'        Errorcode = InStr(strFile, ChrW(k))
'        If Errorcode > 0 Then Exit For
'1000
'        Next

'Sub nin() '����ʹ��
'Dim rngmd As Range
'For i = 1 To 2
'filemd5 = UCase(Hashpowershell(Range("e" & 619 + i)))
'With ThisWorkbook.Sheets("���")
'Set rngmd = .Range("k6:k" & .[k65536].End(xlUp).Row).Find(filemd5) '����Ƿ��ļ��Ѵ���
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

'Sub kin() '����
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



'Function UDF_Translate(strText As String) As String '�����е���ѯ-��ѯ����ר��
'
'    Dim urlDict As String, urlTranslate As String, xmlText As String
'    Dim PhoneticSymbol As String, TranslateArr
'
'    '����url
'    urlDict = "http://dict.youdao.com/search?q=" & strText & "&doctype=xml"
'    urlTranslate = "http://fanyi.youdao.com/translate?i=" & strText & "&doctype=xml"
'
'    'ʹ�� WebService �� FilterXML ��ȡ��������
'    On Error GoTo Error01
'    xmlText = Application.WorksheetFunction.WebService(urlDict)
'    'Debug.Print xmlText    '������
'    PhoneticSymbol = Application.WorksheetFunction.FilterXML(xmlText, "//phonetic-symbol")
'    TranslateArr = Application.WorksheetFunction.FilterXML(xmlText, "//translation/content[1]")
'
'    '��������
'    If IsArray(TranslateArr) Then
'        UDF_Translate = "[" & PhoneticSymbol & "] " & Join(Application.WorksheetFunction.Transpose(TranslateArr), "��")
'    Else
'        UDF_Translate = "[" & PhoneticSymbol & "] " & TranslateArr
'    End If
'    'Debug.Print Translate  '������
'Exit Function
'Error01:    '���е��ʵ䱨��
'    On Error GoTo Error02
'    xmlText = Application.WorksheetFunction.WebService(urlTranslate)
'    'Debug.Print xmlText    '������
'    UDF_Translate = Application.WorksheetFunction.FilterXML(xmlText, "//translation")
'    'Debug.Print Translate  '������
'Exit Function
'Error02:    '���е����뱨��
'    UDF_Translate = Err.Description
'End Function


'Option Explicit
'
'Dim arrfiles(1 To 10000)                  '����һ������������Դ��path����(��ֵ�ɱ�)
'Dim arrbase(1 To 10000)                   '�洢��������չ�����ļ���
'Dim arrextension(1 To 10000)              '�洢�ļ���չ��
'Dim arrsize(1 To 10000)                   '�洢�ļ��Ĵ�С
'Dim arrparent(1 To 10000)                 '�洢�ļ�����λ��
'Dim arrdate(1 To 10000)                   '�洢�ļ���������
'Dim arrsizeb(1 To 10000)                  '�ļ��Ĵ�С,��λ����
'Dim flc As Integer                        '��ͬsub֮�������ͬ�ı���,ע��Ҫʹ��ģ�鼶�Ķ���
'Sub fleshdata(arr()) '��������
'
'Dim filepath$
'Dim fd As Folder
'Dim i As Integer, j As Integer
'Dim elow As Integer
'Dim k As Integer
'Dim arrc()
''Dim fso As Object
''Set fso = CreateObject("Scripting.FileSystemObject") '�����������ô���fso����(���ڰ�)
'
'Application.ScreenUpdating = False         '�ر���Ļʵʱ����,�ӿ����������ٶ�
'With ThisWorkbook.Sheets("���")
'ThisWorkbook.VBProject.VBComponents.Item(3).CodeModule.InsertLines 114, "exit sub"
'If .Range("b6") <> "" Then Call checkfile '����ļ��Ƿ����
'
'For j = 0 To UBound(arr())
'
'If arr(j) = "" Then GoTo 100              'ע��˴�arr("j")��ǰ���ڿ�
'
'filepath = arr(j)
'
'flc = 0                                   '�ļ������ĳ�ʼֵ
'
'Set fd = fso.GetFolder(filepath)          '��fdָ��·������
'
'search fd                                     '����sf��sub
'
'If flc = 0 Then GoTo 100  '����ӵ��ļ��������ļ�ʱ,������һ��ѭ��
'
'    .Range("c" & .[c65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrbase) '�ļ���
'    .Range("d" & .[d65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrextension) '�ļ���չ��
'    .Range("e" & .[e65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrfiles) '�ļ�·��
'    .Range("f" & .[f65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrparent) '�ļ�����λ��
'    .Range("g" & .[g65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrsizeb) '�ļ���ʼ��С
'    .Range("h" & .[h65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrsize) '�ļ���С
'    .Range("i" & .[i65536].End(xlUp).Row + 1).Resize(flc) = Application.Transpose(arrdate) '�ļ�����ʱ��
'    .Range("w" & .[w65536].End(xlUp).Row & ":" & "w" & .[w65536].End(xlUp).Row + flc) = Now
'    .Range("b" & .[b65536].End(xlUp).Row).AutoFill Destination:=.Range("b" & .[b65536].End(xlUp).Row & ":" & "b" & .[b65536].End(xlUp).Row + flc), Type:=xlFillDefault '���ͳһ����
'100
'
'Next
'
'Call deletempfile
'End With
'ThisWorkbook.VBProject.VBComponents.Item(3).CodeModule.DeleteLines 114
'Application.ScreenUpdating = True         '���д�����Ϻ�������Ļ����
'
'End Sub
'
'Sub search(ByVal fd As Folder)               'ByVal,Ϊ��ֵ���ݷ�ʽ, �б���byref(�����÷�ʽ����)
'                                         'searchfiles(ByVal fd As Folder)��function������
'Dim fl As File
'Dim sfd As Folder
'Dim arr()
'Dim i As Integer
'Dim arrtemp()
'
'With ThisWorkbook.Sheets("���")
'
'If .[c65536].End(xlUp).Row - 5 > 1 Then arrtemp = .Range("c6:c" & .[c65536].End(xlUp).Row).Value '����valueΪrange��cell��ȱʡֵ,���ǻ���Ҫд��,��ֹ������ʱ�޷���Чʶ�������
'
'For Each fl In fd.Files
'
'    flc = flc + 1
'
'If .[c65536].End(xlUp).Row - 5 > 1 Then
'
'    For i = 1 To .[b65536].End(xlUp).Row - 5
'    If fso.GetFileName(fl.Path) = arrtemp(i, 1) Then '�ļ�����ͬ�Ĳ����
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
'    If fso.GetFileName(fl.Path) = .Range("b6").Value Then '�ļ�����ͬ�Ĳ����
'
'    flc = flc - 1
'
'    GoTo 100
'
'    End If
'
'    Else              '��ֵ,ֱ������
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
'        arrsize(flc) = Format(fl.Size / 1024, "0.00") & "KB"    '�ļ��ֽڴ���1048576��ʾ"MB",������ʾ"KB"
'        Else
'        arrsize(flc) = Format(fl.Size / 1048576, "0.00") & "MB"
'    End If
'
'100
'
'Next fl
'
'If fd.SubFolders.Count = 0 Then Exit Sub  '���ļ�����ĿΪ�����˳�sub
'
'For Each sfd In fd.SubFolders             '�������ļ���
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
'        Set rnglist = .Cells(4, bc).Resize(.[b65536].End(xlUp).Row, 1).FindPrevious(rnglist) '����ǰ���ֵ
'        faddr = rnglist.address '��¼�µ�ַ
'        frow = rnglist.Row
'105
'
'            Set fd = fso.GetFolder(rnglist.Value) 'Sheets(1).Cells.SpecialCells(xlCellTypeLastCell).Row
'            Do
'            Set fd = fd.ParentFolder
'                If fd.Path = strfolder Then '�ҵ�ƥ��ֵ
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
'                    If c = UBound(Split(rnglist.Value, "\")) Then 'ͬ���ļ��Ƚ�
'                       If fdp = fd.Path Then GoTo 103
'                    End If
'                End If
'            Loop Until fd.Drive & "\" = fd.ParentFolder 'ѭ������һ���ļ�
'            Set rnglist = .Cells(4, bc).Resize(.[b65536].End(xlUp).Row, 1).FindPrevious(rnglist) 'û��ƥ�䵽ֵ
'                If rnglist.address = faddr Then 'ѭ������ʼֵ
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

On Error GoTo 100 '��ֹ��ȡ����ע�����Ϣ
With CreateObject("wscript.shell") '��ȡע�����������Ϣ
cversion = .RegRead("HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon\version") '��λ�ô洢��������İ汾��Ϣ
If cversion = "" Then Exit Sub 'û�л�ȡ����Ϣ


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
        Result = .RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\7-Zip\DisplayIcon") '��ȡ��׺����Ӧ��ע������
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

Sub EnumSEVars() '�ļ��б���
        Dim strVar As String
        Dim i As Long
        For i = 1 To 255
            strVar = Environ$(i)
            If LenB(strVar) = 0& Then Exit For
            Cells(i, 1) = strVar
        Next
End Sub

'ע�����ʹ��ʱ�õ���˫����"" ""
'If Me.TextBox11.text <> "" And Me.TextBox12 <> "" Then                                                                               '�������ö�����ʱ
'   If fso.FileExists(mid$(Me.TextBox11.text, 2, Len(Me.TextBox11.text) - 3)) = False Or fso.FileExists(mid$(Me.TextBox12.text, 2, Len(Me.TextBox12.text) - 3)) = False Then
'   Call warning(2)
'   Me.TextBox1.SetFocus
'   Exit Sub
'   Else
'   exepatha1 = Me.TextBox11.text
'   exepatha2 = Me.TextBox12.text
'   ThisWorkbook.VBProject.VBComponents.Item(8).CodeModule.ReplaceLine 33, "exepath = " & "" & exepatha1 & ""          'ע����������λ���ں����޸��п��ܳ��ֵ�λ�ñ仯
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
'.Add , , "Menus", "Menus" '��Ŀ¼
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
's = 1   '����
'Next
'    End With
'    Me.TreeView1.Nodes(1).Expanded = True
'End Sub

Sub aoso()
MsgBox Environ("ProgramW6432")
End Sub
'������ļ���
'If .Sheets("���").Range("b6") = "" Then GoTo 1005
'If .Sheets("���").[d65536].End(xlUp).Row = 6 Then
'Me.ListBox3.AddItem
'Me.ListBox3.List(Me.ListBox3.ListCount - 1, 0) = .Sheets("���").Range("f6").Value
'Me.ListBox3.List(Me.ListBox3.ListCount - 1, 0) = .Sheets("���").Range("w6").Value
'Else
'Dim dicadd As New Dictionary
'Dim foldnum As Integer
'arral = .Sheets("���").Range("f6:f" & .Sheets("���").[f65536].End(xlUp).Row).Value
'For foldnum = 1 To .Sheets("���").[f65536].End(xlUp).Row - 5
'dicadd(arral(foldnum, 1)) = ""
'Next
'End If
'Me.ListBox3.List = dicadd.Keys

'������ļ���

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

                
'                .Range("d" & rngad.Row & ":" & "j" & rngad.Row).ClearContents '���������
'                If rngad.Row <> .[d65536].End(xlUp).Row Then '��������һ��,��ֱ�Ӵ���
'                    k = rngad.Row
'                    j = .[d65536].End(xlUp).Row
'                    For i = k To j - 1
'                    .Range("d" & i) = .Range("d" & i + 1)
'                    Next
'                End If
'            End With
'        End If


'If ThisWorkbook.Sheets("���").Range("d" & addrowx).Value <> "txt" And Len(ThisWorkbook.Sheets("���").Range("ab" & addrowx).Value) = 0 Then '��Ϊtxt�ļ�'·�������������ַ�
'
'Open address For Binary Access Read Write Lock Read Write As #1  '�ж��ļ��Ƿ��ڴ򿪵�״̬��֧�ַ�ansi����, �����txt�ĵ�,��ô�Ͳ���Ҫ�ж��Ƿ��(txt�ļ��ڴ򿪵�״̬�²��ᱻ����(notepad))
'Close #1
'
'If Err.Number <> 0 Then '�����ļ���״̬�����ļ����ִ������ж��ļ��Ƿ��ڴ򿪵�״̬
'      Me.Label1.Caption = "�ļ�����ʹ����,��رպ�����"
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
'With ThisWorkbook.Sheets("���")
'str1 = newname & "." & .Range("d" & addrowx)
'If fso.FileExists(.Range("f" & addrowx) & "\" & str1) = True Then
'Me.Label1.Caption = "�ļ�����" '�ж��ļ��Ƿ�����
'Me.TextBox1.text = ""
'Exit Sub
'Else
'
'If Errcode(newname, "N", 1) > 0 Then '���ļ���������ȥ����Ƿ�����쳣�ַ�
'Me.Label1.Caption = "������ļ��������쳣�ַ�" '�޸��ļ������Ƿ�����쳣�ַ�''�����������Ƴ�������/���ڿո�,������ִ������-���ѭ��(inputbox)
'Me.TextBox1.text = ""
'Exit Sub
'Else
'
'If Len(ThisWorkbook.Sheets("���").Range("ab" & addrowx).Value) > 0 Then '�ļ�·�������쳣�ַ�
'
'Shell "cmd /c rename " & address & Chr(32) & str1, 0   '����cmd ��rename����ȥִ��'���Դ����쳣�ַ�'�����޷������ļ��Ƿ��ڴ򿪵�״̬
'Sleep 100 '������Ҫ�ǳ�ע��,cmdִ�е��ٶ�,����ʱ�ݻ����������ִ��,�ȴ�cmd�����ִ��,����fso�޷��ж��µ��ļ����Ĵ���
'DoEvents
'str1 = .Range("f" & addrowx) & "\" & str1 '�µ��ļ�·��
'If fso.FileExists(str1) = True Then '����ļ��Ƿ��Ѿ��޸ĳɹ�
'.Range("e" & addrowx) = str1
'.Range("c" & addrowx) = newname & "." & Range("d" & addrowx)
'Else
'Me.Label1.Caption = "�ļ����ڴ�״̬,����ʧ��"
'End If
'
'Else
'Name address As .Range("f" & addrowx) & "\" & newname & "." & .Range("d" & addrowx) 'name��֧�ַ�ansi������ַ���
'.Range("e" & addrowx) = .Range("f" & addrowx) & "\" & newname & "." & .Range("d" & addrowx)
'.Range("c" & addrowx) = newname & "." & Range("d" & addrowx)
'End If
'
'
'
'
'End If
'.Range("ae" & addrowx) = "" '���ԭ�������ִ����쳣�ַ���ô��ȥ��������
'.Range("ab" & addrowx) = ""
'
'If .Range("d" & addrowx) = "txt" Then Me.Label2.Caption = "txt�ĵ������ڴ򿪵�״̬���ƶ�\������\ɾ��"      '�����û�,txt�ĵ������ڴ򿪵�״̬���ƶ����������Ȳ���
'str = .Range("f" & addrowx).Value & "\"
'End With
'If ThisWorkbook.Sheets("���").Range("ad" & addrowx).Value = 1 Then
'Set rngtime = ThisWorkbook.Sheets("Ŀ¼").Cells(4, 3).Resize(ThisWorkbook.Sheets("Ŀ¼").[b65536].End(xlUp).Row, ThisWorkbook.Sheets("Ŀ¼").Cells.SpecialCells(xlCellTypeLastCell).Column).Find(str, lookat:=xlWhole)
'If Not rngtime Is Nothing Then rngtime.Offset(0, 2) = Now '�ļ����޸ĳɹ������ڵ��ļ��е��޸�ʱ�䷢�����
'End If
'.Label1.Caption = "�޸ĳɹ�!"
'.CommandButton2.Enabled = True '
'End With
'
'Set rngtime = Nothing
'
'Exit Sub
'
'101
'Err.Clear
'Me.Label1.Caption = "�޸��ļ���ʧ��"


'    ElseIf c = 2 Then
'    tfolder = tfolder & "\"
''    Set fd = fso.GetFolder(tfolder)
''    Set fd = fd.ParentFolder
'    tfolderp = fd.Path & "\"
'    Set rngad = .Range("f6:f" & .[f65536].End(xlUp).Row).Find(tfolderp, lookat:=xlPart) '����ļ����Ƿ����������ļ�����Ŀ¼
'    If rngad Is Nothing Then '�����������ļ��еĹ����ļ���
'    With ThisWorkbook.Sheets("Ŀ¼")
'    Set rngad = .Cells(4, 3).Resize(.[b65536].End(xlUp).Row, .Cells.SpecialCells(xlCellTypeLastCell).Column).Find(tfolder, lookat:=xlWhole) '��ȷ����
'    If Not rngad Is Nothing Then
'    If .AutoFilterMode = True Then .AutoFilterMode = False 'ɸѡ������ڿ���״̬��ر�
'   .Range("a3:a" & .[a65536].End(xlUp).Row).AutoFilter Field:=1, Criteria1:=rngad.Offset(0, -2).Value
'   .Range("a4").Resize(.[a65536].End(xlUp).Row - 3).SpecialCells(xlCellTypeVisible).Delete Shift:=xlShiftUp 'ɾ����ɸѡ�����Ľ��
'   .Range("a3:a" & .[a65536].End(xlUp).Row).AutoFilter
'    End If
'    End With
'    End If
    
'    Else
'    Set fd = fso.GetFolder(tfolder) 'ע�ⲻҪ�Ž�doѭ������
'    Do
'    tfolderp = fd.Path & "\" 'ע�������λ�ú��ļ��еĲ㼶
'    Set fd = fd.ParentFolder '�㼶����
'    Set rngad = .Range("f4:f" & .[f65536].End(xlUp).Row).Find(tfolderp, lookat:=xlPart) '����ļ����Ƿ����������ļ�����Ŀ¼
'    If Not rngad Is Nothing Then n = n + 1 '�����������ļ��еĹ����ļ���
'    Loop Until fd.IsRootFolder
'    If n = 0 Then  '������й����ļ�
    
    
    
'
'    End If
'    End With
    
'
'    tfolderp = tfolder & "\"
'    With .Sheets("Ŀ¼")
'    Set rngad = .Cells(4, 3).Resize(.[b65536].End(xlUp).Row, .Cells.SpecialCells(xlCellTypeLastCell).Column).Find(tfolderp, lookat:=xlWhole) '��ȷ����
'    If Not rngad Is Nothing Then
'    rngadr1 = rngad.address '��һ��λ��
'    rngadrowx = rngad.Row
'    rngadcolumnx = rngad.Column
'    Set rngad = .Cells(4, rngadcolumnx + 1).Resize(.[b65536].End(xlUp).Row, .Cells.SpecialCells(xlCellTypeLastCell).Column).Find(tfolderp, lookat:=xlPart) 'ģ������,�Ƿ������ļ���
'        If rngad Is Nothing Then '���û��
'        Rows(rngadrowx).Delete Shift:=xlShiftUp
'        Else
'        If rngad.address <> rngadr1 And yesno = vbYes Then rngad.Offset(0, 2) = Now 'ɾ�����ļ��е�ʱ�䷢���仯' '���������ļ���
'        End If
'    End If
'    End With
'
'End With
'Set fd = Nothing


'
'Function tempdele() '��Ҫע���кŵı仯
'    ThisWorkbook.VBProject.VBComponents.Item(1).CodeModule.DeleteLines 4  'ɾ�������ִ��루ɾ���������λ�ñ仯��'�״����е�ʹ��
'    ThisWorkbook.VBProject.VBComponents.Item(1).CodeModule.DeleteLines 11
'    ThisWorkbook.VBProject.VBComponents.Item(1).CodeModule.DeleteLines 36
'    ThisWorkbook.VBProject.VBComponents.Item(1).CodeModule.DeleteLines 169
'    ThisWorkbook.VBProject.VBComponents.Item(1).CodeModule.DeleteLines 146 'ɾ����ִ��
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
'        If Flagpause = False Then Return '����ִ��
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
'MsgBox "��ǰϵͳΪ����������", , "����"
'Case &H804
'MsgBox "��ǰϵͳΪ����������", , "����"
'Case &H409
'MsgBox "��ǰϵͳΪ��Ӣ��", , "����"
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
'    '������ҳ
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
'        'ȡ�����겿��
'    ybEN = "Ӣ " & Split(Split(str_base, "EN-US"">")(1), "<")(0)
'    ybUS = " �� " & Split(Split(str_base, "EN-US"">")(2), "<")(0) ' & "]"
'
'    str_base1 = Split(Split(str_base, "<ol slider=""2"">")(1), "</ol>")(0)
'    v = Split(str_base1, "<li>")
'
'        'ȡ�����ĺ��岿��
'    For i = LBound(v) + 1 To UBound(v) - 1
'        hytmp = hytmp & "��" & Split(v(i), "<")(0)
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
'        '������ҳ
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
'            'ȡ�����겿��
'    ybEN = "Ӣ " & Split(Split(str_base, "value-en"">")(1), "<")(0)
'    ybUS = " �� " & Split(Split(str_base, "-us"">")(1), "<")(0)
'
'        str_base1 = Split(Split(str_base, "<ul>")(2), "</ul>")(0)
'        v = Split(str_base1, "<span")
'
'            'ȡ�����ĺ��岿��
'        For i = LBound(v) + 1 To UBound(v)
'            hytmp = hytmp & Split(v(i), "</span>")(0)
'        Next i
'        tmpPhonetic = ybEN & " " & ybUS
'tmpTrans = Mid(Replace(Replace(hytmp, "class=""attr"">", Chr(10)), ">", ""), 3)
'
End Sub
'http://www.iciba.com/branch

Sub searchWordFromBaidu(tmpWord As String, tmpTrans As String, tmpPhonetic As String)
'    'http://dict.baidu.com/s?wd=����  'https://fanyi.baidu.com/?aldtype=85#en/zh/truths
'    'https://fanyi.baidu.com/?#en/zh/truths
'    Dim XH As Object
'    Dim str_base As String, URL
'    If Len(tmpWord) = 0 Then Exit Sub
'    tmpTrans = "": tmpPhonetic = ""
'    '������ҳ
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
'    'ȡ�����겿�֡�
'    ybEN = "Ӣ [" & Split(Split(str_base, ">[")(1), "]<")(0) & "] "
'    ybUS = "�� [" & Split(Split(str_base, ">[")(2), "]<")(0) & "]"
'    '���ĺ���
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
'        '������ҳ
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
'            'ȡ�����겿��
'    ybEN = "Ӣ " & Split(Split(str_base, "KK:")(1), " ")(0)
'    ybUS = " �� " & Split(Split(str_base, "DJ:")(1), "<")(0)
'
'        str_base1 = Split(Split(str_base, "<ul>")(2), "</ul>")(0)
' 'Debug.Print Mid(str_base1, 1, 3000)
'        v = Split(str_base1, "<span")
'
'            'ȡ�����ĺ��岿��
'        For i = LBound(v) + 1 To UBound(v)
'            hytmp = hytmp & Split(v(i), "</span>")(0)
'        Next i
'        tmpPhonetic = ybEN & " " & ybUS
'tmpTrans = Mid(Replace(Replace(hytmp, "class=""attr"">", Chr(10)), ">", ""), 3)
'
End Sub
