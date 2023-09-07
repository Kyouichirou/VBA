Attribute VB_Name = "豆瓣"
Option Explicit
'private const
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) '时间 -单词训练
Private Declare Function timeGetTime Lib "winmm.dll" () As Long '时间 -单词训练
Private Const BdUrl As String = "https://www.baidu.com"
Private Const ByUrl As String = "https://cn.bing.com/search?q=site:book.douban.com%20"
'-------处理网页内容,使用正则的好处
'可以匹配复杂的规则-因为网页反馈的信息非常混乱,仅仅通过split,left等函数处理起来过于麻烦,而且在处理信息量较大的内容时,速度较慢

Function DoubanBook(ByVal strWord As String) '获取豆瓣读书的评分'由于豆瓣对书籍的信息进行了加密处理，借道必应来获取数据(360搜索也能够够抓取豆瓣的评分)
    Dim strx As String, url As String
    Dim xtemp As Variant, arr() As String
    Dim t As Long
    
    With UserForm3
        If TestURL(BdUrl) Then '检查网络链接是否处于正常状态
            url = ByUrl & encodeURI(strWord) & "&count=1" '指定必应去搜索豆瓣读书的信息
            '------------------------------------------------------------------------------------count=1为搜索结果参数,表示返回一个搜索结果
            'site:book.douban.com-表示指定搜索豆瓣读书的信息,%20表示 "+"
            '-------------------------------------因为搜索的内容会有多个干扰结果,所以返回单一结果,准确度不一定高
            With CreateObject("MSXML2.XMLHTTP")
                .Open "GET", url, True
                .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                .Send
                t = timeGetTime
                Do While .readyState <> 4 And timeGetTime - t < 5000 '等待数据的返回
                    DoEvents
                Loop
                strx = .responseText
            End With
'                If InStr(strx, "https://book.douban.com/review/") > 0 Then GoTo 100 '需要进一步细化返回的结果-部分旧的页面采用的是5分制的评分,新版本的评分为10分制
                If InStr(strx, "用户评级") > 0 Then '表示获取到有效的搜索结果
                   ReDim arr(2)
                   arr = DoubanData(strx)
                   If Len(arr(0)) = 0 Or Len(arr(1)) = 0 Or Len(arr(2)) = 0 Then GoTo 100
                    .TextBox16.Text = arr(0)
                    .TextBox17.Text = arr(2)
                    .TextBox15.Text = arr(1)
                    .TextBox16.Visible = True
                    .TextBox17.Visible = True
                    .TextBox15.Visible = True
                    .CommandButton54.Visible = True
                    .Label56.Caption = "" '清空信息
                    .Label57.Caption = "操作成功"
                Else
100
                    .Label57.Caption = "未找到书籍信息"
                    .Label56.Caption = "未找到书籍信息"
                End If
        Else
            .Label57.Caption = "网络连接异常"
        End If
    End With
End Function

Private Function DoubanData(ByVal xtext As String) As String()
    Dim myreg As Object, match As Object, Matches As Object
    Dim arr() As String
    Dim arreg(), i As Byte
    '------------------------------------------------------https://tool.oschina.net/regex/,正则表达式在线测试
    '----------------------------------------------------- https://www.runoob.com/regexp/regexp-metachar.html 规则
    ReDim DoubanData(2)
    ReDim arr(2)
    ReDim arreg(2)
    arreg = Array("用户评级:+(.+?)(\<)", "[\>]+(.+?)[\(]豆瓣[\)]", "[https]+://+book.douban.com/+[a-z]*\/+[0-9]*\/")
    Set myreg = CreateObject("VBScript.RegExp") '正则表达式
    For i = 0 To 2
        With myreg
            .Pattern = arreg(i) '获取豆瓣评分
            .Global = True
            .IgnoreCase = True '不区分大小写
            Set Matches = .Execute(xtext)
            For Each match In Matches
                arr(i) = match.Value: Exit For
            Next
            If i = 1 Then xtext = arr(i) '二次处理结果
        End With
        Set match = Nothing
        Set Matches = Nothing
    Next
    arr(0) = Trim(Replace(Replace(arr(0), "用户评级:", ""), "<", "")) '获取用户评级
    If Len(arr(1)) > 0 Then arr(1) = Trim(Right$(arr(1), Len(arr(1)) - InStrRev(arr(1), ">")))
    DoubanData = arr
    Set myreg = Nothing
End Function

Function ObtainDoubanPicture(ByVal url As String) As String() '获取豆瓣封面链接/国籍,作者
    Dim myreg As Object, match As Object, Matches As Object
    Dim IE As Object, i As Byte
    Dim strx As String, strx1 As String, strx2 As String
    Dim arreg(), arr(2) As String, arrTemp(1)
    Dim t As Long
    '---------------https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa752093(v=vs.85)?redirectedfrom=MSDN
    Set IE = CreateObject("InternetExplorer.Application") '创建ie浏览器对象
    IE.Visible = False
    IE.Navigate url
    t = timeGetTime
    Do While IE.readyState <> 4 And timeGetTime - t < 7000 '等待最长时间不超过7s,必须增加时间控制, ie打开豆瓣, 防止出现页面假死形成的死循环
        DoEvents
    Loop
    If IE.Busy = True Then IE.Stop '停止继续加载网页
    With IE.Document
        arrTemp(0) = .getElementById("mainpic").InnerHtml '书名+封面
        arrTemp(1) = .getElementById("info").InnerHtml '作者+国籍
    End With
    ReDim arreg(1)
    arreg = Array("[a-zA-z]+://[^\s]*", "[\>]+(\s.+?)+[\<]\/a[\>]")
    Set myreg = CreateObject("VBScript.RegExp") '正则表达式
    IE.Quit
    Set IE = Nothing
    For i = 0 To 1
        With myreg
            .Pattern = arreg(i) '获取豆瓣图片链接
            .Global = True
            .IgnoreCase = True '不区分大小写
            Set Matches = .Execute(arrTemp(i))
            For Each match In Matches
                arr(i) = match.Value: Exit For
            Next
        End With
        Set match = Nothing
        Set Matches = Nothing
    Next
    strx1 = Trim(arr(0))
    arr(0) = Left$(strx1, Len(strx1) - 2) '图片链接
    strx2 = arr(1)
    strx2 = Trim(Replace(Replace(strx2, "</a>", ""), ">", ""))
    strx2 = Trim(Replace(strx2, Chr(10), "", 1, 2))
    If InStr(strx2, "[") > 0 And InStr(strx2, "]") > 0 Then
        arr(2) = AuthorNT(strx2) '国籍
        strx2 = Trim(Right(strx2, Len(strx2) - Len(arr(2))))
    End If
    arr(1) = strx2 '作者
    ReDim ObtainDoubanPicture(2)
    ObtainDoubanPicture = arr
    Set myreg = Nothing
End Function

Function DoubanTreat(ByVal gradex As String, ByVal authorx As String, ByVal bookx As String) As String() '豆瓣信息处理
    Dim strx As String, strx1 As String, strx2 As String, strx3 As String, arr() As String
    Dim strx4 As String
    
    On Error Resume Next
    ReDim DoubanTreat(1 To 5)
    ReDim arr(1 To 5)
    If InStr(gradex, "v:average") > 0 Then
        strx = Split(Split(gradex, "v:average")(1), "</")(0)
        arr(1) = Trim(Mid$(strx, 3, Len(strx) - 2)) '评分
    End If
    If InStr(authorx, "</a>") > 0 Then
        strx1 = Split(authorx, "</a>")(0)
        strx4 = Trim(Replace(Trim(Right$(strx1, Len(strx1) - InStrRev(strx1, ">"))), Chr(34), "", 1, 2)) '作者
        strx4 = Trim(Replace(strx4, Chr(34), "", 1, 2))
        strx4 = Trim(Replace(strx4, Chr(10), "", 1, 2)) 'chr(10)换行符,注意和vbcr等常量有所区别
        If InStr(strx4, "[") > 0 And InStr(strx4, "]") > 0 Then '[日]东野圭吾
            arr(5) = AuthorNT(strx4) '作者国籍
            strx4 = Trim(Right(strx4, Len(strx4) - Len(arr(5))))
        End If
        arr(2) = strx4 '作者
    End If
    
    If InStr(bookx, "title=") > 0 Then
        strx2 = Split(Split(bookx, "title=")(1), Chr(34))(1)
        arr(3) = Trim(strx2)
    End If
    If InStr(bookx, "src") > 0 Then
        strx3 = Trim(Split(Split(bookx, "src")(1), " ")(0))
        arr(4) = Trim(Replace(Trim(Mid$(strx3, 2, Len(strx3) - 1)), Chr(34), "", 1, 2)) '封面
    End If
    DoubanTreat = arr
End Function

Private Function AuthorNT(ByVal textx As String) As String '豆瓣-获取作者的国籍-正则表达式
    Dim myreg As Object, match As Object, Matches As Object
    '------------------------------------------------------https://tool.oschina.net/regex/,正则表达式在线测试
    '----------------------------------------------------- https://www.runoob.com/regexp/regexp-metachar.html 规则
    Set myreg = CreateObject("VBScript.RegExp") '正则表达式
    With myreg
        .Pattern = "(\[)+(.+?)+(\])" '匹配[日]这种数据但不包含单独的"["或者"]",也不匹配"[]",也不匹配"[日"或者"日]" ,\[表示转义, 用于目标是"["这个符号
        .Global = True
        .IgnoreCase = True '不区分大小写
        Set Matches = .Execute(textx)
        For Each match In Matches
            AuthorNT = match.Value: Exit For
        Next
    End With
    Set myreg = Nothing
    Set match = Nothing
    Set Matches = Nothing
End Function
