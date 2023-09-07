Attribute VB_Name = "模块8"
'Option Explicit
Sub fkkf()

'Dim fl As TextStream
'Set fl = fso.OpenTextFile("C:\Users\adobe\Desktop\x.txt", ForReading, False, TristateUseDefault)
'sResult = fl.ReadAll
''fl.Close
Set fl = fso.OpenTextFile("C:\Users\adobe\Desktop\p.txt", ForReading, False, TristateUseDefault)
s = fl.ReadAll
fl.Close
Set fl = Nothing
Dim x As Object
Set x = GetObject("C:\Users\adobe\Desktop\doubanjs.html")
'strWord = "hello"
'ObtainObjInfo x.parentWindow.decrypt
a = x.parentwindow.xdecrypt(s)
a = CallByName(x.parentwindow, "decrypt", VbMethod, s)
'Dim js As Object
'Dim oDom As Object, oWin As Object
'Set oDom = CreateObject("htmlfile")
''oDom.DesignMode = "on" ' 开启编辑模式(不要直接使用.body.innerhtml=shtml,这样会导致IE浏览器打开)
'oDom.Write sResult ' 写入数据
'Set oWin = oDom.parentWindow
'oWin.execScript
''s = oWin.Result
'x = oWin.eval("decrypt")
'Set oDom = Nothing
'Set oWin = Nothing
'Debug.Print x
Stop
End Sub

Sub QQ1720002187970()
    Dim sJS As String
    Dim arrElement
    Dim obj()
    'JS对象
Dim fl As TextStream
Set fl = fso.OpenTextFile("C:\Users\adobe\Desktop\x.txt", ForReading, False, TristateUseDefault)
sResult = fl.ReadAll
fl.Close
''Set fl = fso.OpenTextFile("C:\Users\adobe\Desktop\p.txt", ForReading, False, TristateUseDefault)
''s = fl.ReadAll
''fl.Close
Set fl = Nothing
's = " var Result = " & sResult

    Dim oHtml As Object
    '定义HtmlDocument对象
    Set oHtml = CreateObject("htmlfile")
    Dim oWindow As Object
    Set oWindow = oHtml.parentwindow
    With oWindow
        .execScript sResult
       '用callbyname获取js对象的属性值
       x = CallByName(.obj, "decrypt", VbGet)
    End With
End Sub

Function BitShiftRight(ByVal i As Variant)
    Dim sJS As String
    sJS = " var Result = " & i & " >> 5"
    Dim oHtml As Object
    '定义HtmlDocument对象
    Set oHtml = CreateObject("htmlfile")
    Dim oWindow As Object
    Set oWindow = oHtml.parentwindow
    With oWindow
        .execScript sJS
      BitShiftRight = .Result
    End With
    Set oWindow = Nothing
    Set oHtml = Nothing
End Function
Sub QQ1722187970()
    Debug.Print
    Debug.Print BitShiftRight(64)
End Sub

Function JSEval(s As String) As String
With CreateObject("MSScriptControl.ScriptControl")
    .Language = "javascript"
    JSEval = .eval(s)
End With
End Function
Private Function ObtainObjInfo(ByVal objx As Object) As String() '获取对象属性名称
    Dim i As Integer, k As Integer
    Dim arr() As String
    Dim objinfo As Object

    Set oTli = CreateObject("TLI.TLIApplication")
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

Function GoogleTranslate(strWord As String, Optional Mode As Boolean = False) As String
    'Mode为TRUE则为汉译英，为FALSE则为英译汉，默认是FALSE
    Dim strURL As String
    Dim strText As String
    Dim strJSScript As String
    Dim objHTTP As Object
    Dim TKKFunc As String
    Dim OtherFunc As String
    Dim objHTML As Object
    Dim DataFunc As String
    Dim tkValue As String
    Dim EncodeWord As String
    Dim strMode As String
    
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    Set objHTML = CreateObject("htmlfile")
    
    '获取TKK函数
    strURL = "http://translate.google.cn/"
    strText = GetReponseText(objHTTP, strURL)
    TKKFunc = "TKK=" & Split(Split(strText, "TKK=")(1), "');")(0) & "');"
    
    '获取其他函数
    strURL = "http://translate.google.cn/translate/releases/twsfe_20161212_RC00/r/js/desktop_module_main.js" 'tkk:'440673.3623724348'
    strText = GetReponseText(objHTTP, strURL)
    OtherFunc = "var gk=" & Split(Split(strText, "var gk=")(1), "var kk=")(0)
    
    '合成完整的tk算法函数，并加上html代码：
    strJSScript = "<html><script>" & TKKFunc & OtherFunc & "</script></html>"
    
    '计算单词的tk值
    objHTML.Write strJSScript
    tkValue = CallByName(objHTML.parentwindow, "jk", VbMethod, strWord)
    
    '将单词进行编码
    EncodeWord = CallByName(objHTML.parentwindow, "encodeURIComponent", VbMethod, strWord)
    
    '从服务器获取翻译数据
    If Mode Then
        strMode = "&sl=zh-CN&tl=en"
    Else
        strMode = "&sl=en&tl=zh-CN"
    End If
    strURL = "http://translate.google.cn/translate_a/single?client=t" _
        & strMode & "&hl=zh-CN" _
        & "&dt=at&dt=bd&dt=ex&dt=ld&dt=md&dt=qca&dt=rw&dt=rm&dt=ss&dt=t" _
        & "&ie=UTF-8&oe=UTF-8&source=bh&ssel=0&tsel=0&kc=1" _
        & tkValue _
        & "&q=" & EncodeWord
    strText = GetReponseText(objHTTP, strURL)
    
    '自定义处理数据的js函数
    DataFunc = "getdata=function(a){var s='';a=eval(a);for(var i=0;i<a[0].length-1;i++)s+=a[0][i][0];return s}"
    strJSScript = "<html><script>" & DataFunc & "</script></html>"
    
    '获取翻译
    objHTML.Write strJSScript
    GoogleTranslate = CallByName(objHTML.parentwindow, "getdata", VbMethod, strText)
    
    Set objHTTP = Nothing
    Set objHTML = Nothing
End Function

Private Function GetReponseText(objHTTP As Object, strURL As String)
    With objHTTP
        .Open "GET", strURL, False
        .setRequestHeader "User-Agent", "Mozilla/4.0"
        .Send
        GetReponseText = .responseText
    End With
End Function


 Sub dkkf()
 Debug.Print xUnisCode2Charter("\xE5\x85\x84\xE5\xBC\x9F\xE9\x9A\xBE\xE5\xBD\x93\xE6\x9D\x9C\xE6\xAD\x8C")
 End Sub


