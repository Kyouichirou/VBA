Attribute VB_Name = "翻译"
Option Explicit
Private Declare Function timeGetTime Lib "winmm.dll" () As Long '时间api可以精确到毫秒

Function GetdicMeaning(ByVal keywordx As String, ByVal excode As Byte) As String '获取翻译
    Dim i As Integer
    Dim strx As String
    
    If excode = 1 Then '英文
        strx = SearchWordFromYoudao(keywordx)
    ElseIf excode = 2 Or excode = 3 Then '中文/存在空格的英文
        strx = SearchWordFromCiba(keywordx)
    End If

    i = Len(Trim(strx))
    If i = 0 Then
        GetdicMeaning = "未获取到信息"
        Exit Function
    ElseIf i > 150 Then
        GetdicMeaning = "返回的信息处理失败"
        Exit Function
    End If
    If strx = "xnothingx" Then GetdicMeaning = "程序异常!": Exit Function
    GetdicMeaning = strx
End Function

Private Function SearchWordFromCiba(tmpWord As String) As String '金山
     Dim XH As Object
     Dim url As String
     Dim s() As String
     Dim str_tmp As String
     Dim str_base As String
     Dim hytmp As String, hy As Variant, i As Integer, j As Integer, k As Integer
     Dim time1 As Long, strx As String
     
     tmpWord = Replace(tmpWord, " ", "%20") '空格连接符是  " % "
     strx = encodeURI(tmpWord) '转码处理
     '------------------------------------------------------------------office2013后支持 Application.EncodeURL
     If strx = "xnothingx" Then SearchWordFromCiba = strx: Exit Function
     url = "http://www.iciba.com/" & strx
     '开启网页
     Set XH = CreateObject("Microsoft.XMLHTTP")
     If XH Is Nothing Then UserForm3.Label57.Caption = "创建对象失败": Exit Function
     On Error Resume Next
     XH.Open "get", url, True '0
     XH.Send (Null)
     On Error Resume Next
     time1 = timeGetTime
     While XH.readyState <> 4
         DoEvents
         If timeGetTime - time1 > 5000 Then '毫秒
            XH.Close
            Set XH = Nothing
            Exit Function
         End If
     Wend
     str_base = XH.responseText
     XH.Close
     Set XH = Nothing
     
     '对中文含义分解
     hy = Replace(Split(Split(str_base, "s="""">")(1), "</ul>")(0), "prop"">", ">" & Chr(10))
     hy = Split(hy, "<span")
     j = LBound(hy)
     k = UBound(hy)
     For i = j + 1 To k
         hytmp = hytmp & Split(Split(hy(i), ">")(1), "<")(0)  'vbCrLf &
     Next i
     
     SearchWordFromCiba = Mid$(hytmp, 2)
End Function

Function SearchWordFromBing(tmpWord As String) As String '必应----备用
'    http://cn.bing.com/dict/search?q=about+to&go=%E6%8F%90%E4%BA%A4&qs=bs&form=CM
    'http://cn.bing.com/dict/search?q=about+to&go=提交&qs=bs&form=CM
    Dim XH As Object
    Dim s() As String
    Dim str_tmp As String, url As String, hytmp As String
    Dim str_base As String, hy As Variant
    Dim i As Integer, j As Integer, k As Integer
    
    tmpWord = Replace(tmpWord, " ", "+") '出现空格的情况
    url = "http://cn.bing.com/dict/search?q=" & tmpWord
    Set XH = CreateObject("Msxml2.XMLHTTP") 'Microsoft.XMLHTTP")
    If XH Is Nothing Then UserForm3.Label57.Caption = "创建对象失败": Exit Function
    On Error Resume Next
    XH.Open "GET", url, 0 'True
    XH.Send '(Null)
     While XH.readyState <> 4
         DoEvents
     Wend
     str_base = XH.responseText
     XH.Close
     Set XH = Nothing
     '对中文含义分解
     hy = Split(hy, "<span class=""pos"">")
     j = LBound(hy)
     k = UBound(hy)
     For i = j + 1 To k
         hytmp = hytmp & DelHtml(Split(hy(i), "</span></span>")(0)) & vbCrLf
     Next i
     If UBound(hy) = 0 Then hytmp = ""
     SearchWordFromBing = Left$(hytmp, Len(hytmp) - 1) '
End Function

Function SearchWordFromYoudao(ByVal tmpWord As String) As String '有道
    'http://dict.youdao.com/search?q=单词&keyfrom=dict.index
    Dim XH As Object
    Dim s() As String, i As Integer, j As Integer, k As Integer
    Dim str_tmp As String, url As String
    Dim str_base As String
    Dim tmpTrans As String, time1 As Long
    
    Set XH = CreateObject("Microsoft.XMLHTTP")
    If XH Is Nothing Then UserForm3.Label57.Caption = "创建对象失败": Exit Function
'    tmpWord = Replace(tmpWord, " ", "%20") '出现空格的情况-无效

    url = "http://dict.youdao.com/search?q=" & tmpWord
    On Error Resume Next
    XH.Open "GET", url, True     '开启网页
    XH.Send
    On Error Resume Next
    time1 = timeGetTime
    While XH.readyState <> 4
        DoEvents
        If timeGetTime - time1 > 5000 Then '毫秒
           XH.Close
           Set XH = Nothing
           Exit Function
        End If
    Wend
    str_base = XH.responseText
    XH.Close
    Set XH = Nothing

    str_tmp = Split((Split(str_base, "<ul>")(1)), "</ul>")(0)
    s = Split(str_tmp, "<li>")
    k = UBound(s)
    j = LBound(s)
    For i = j + 1 To k
        tmpTrans = tmpTrans & Chr(10) & Split(s(i), "</li")(0)
    Next
    SearchWordFromYoudao = Mid$(tmpTrans, 2)
End Function

Function encodeURI(strText As String) As String 'js字符转换 /需要对中文部分进行转码
    Dim obj As Object
    '---------------------Excel也有内置的转换函数 Application.EncodeURL
    Set obj = CreateObject("msscriptcontrol.scriptcontrol")
    If obj Is Nothing Then MsgBox "字符转码Sub异常", vbCritical, "Warning!!!": encodeURI = "xnothingx": Exit Function
    With obj
        .Language = "JavaScript"
        encodeURI = .eval("encodeURIComponent('" & strText & "');")
    End With
    Set obj = Nothing
End Function

Function DelHtml(ByVal strh As String) As String '正则提取字符串
    Dim a As String
    Dim regEx As Object
    'Dim mMatch As Match
    'Dim Matches As matchcollection
    
    a = strh
    a = Replace(a, Chr(13) & Chr(10), "")
'    A = Replace(A, Chr(32), "")
    a = Replace(a, Chr(9), "")
    a = Replace(a, "</p>", vbCrLf)   '给段落后加上回车
    Set regEx = CreateObject("vbscript.regexp")    '引入正则表达式
    With regEx
        .Global = True
        .Pattern = "\<[^<>]*?\>"   '用<>括起来的html符号
        .MultiLine = True  '多行有效
        .IgnoreCase = True  '忽略大小写(网页处理时这个参数比较重要)
        a = .Replace(a, "")   '将html符号全部替换为空
    End With
    a = Trim(a)
    
    '特殊符号处理
    a = Replace(a, "&lt;", "<")
    a = Replace(a, "&gt;", ">")
    a = Replace(a, "&amp;", "&")
    a = Replace(a, "&quot;", "\")
    a = Replace(a, "&-->", vbCrLf)
    a = Replace(a, "&#230;", ChrW(230)) '&#230;
    a = Replace(a, "&#160;", ChrW(160)) '&#160;
    a = Replace(a, "&nbsp;", " ")  '&nbsp;?
    DelHtml = a
End Function






