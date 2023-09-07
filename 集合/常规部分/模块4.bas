Attribute VB_Name = "模块4"
'Option Explicit
Private Declare Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Sub dkkxxld()
'Set sapiX = CreateObject("SAPI.SpVoice")
'sapiX.Volume = 100 '音量
'sapiX.Rate = 0 '语音速率 越大越快
'FlagsAsync = 1 '同步或异步，1是异步
''下面这段是选系统已安装的语音，可以不运行，用控制面板人工手选好的
''set colVoice=sapiX.getVoices() '安装有多少个语音集合可选
''set sapiX.Voice=colVoice(0) '选第1种语音
'
'strText = "I asked God for a bike, but I know God doesn't work that way."
'strText = strText & "So I stole a bike and asked for forgiveness."
'strText = strText & "开始我直接求上帝赐辆自行车。后来我琢磨上帝办事儿不是这个路数。 于是老子偷了一辆然后求上帝宽恕。"
'
'sapiX.Speak strText, FlagsAsync '同步或异步
'MsgBox "调试窗口，注意同步和异步效果"
'sapiX.Skip 20
''sapiX.Pause '暂停SPEAK，异步有效，如果MSGBOX按下去，还在扯的一段就卡察
''MsgBox "PAUSE，继续发音"
''sapiX.Volume = 30
'sapiX.Resume
'Randomize (Timer)
'For i = 0 To 99
'Debug.Print RandomNumx(10)
'Next
''Debug.Print Application.EncodeUrl("&")
Debug.Print DeleteUrlCacheEntry("https://aip.baidubce.com/")
End Sub


Private Function RandomNumx(ByVal randomnum As Integer) As Integer '随机数
    Dim RndNumber, i As Byte
    
    Randomize (Timer) '初始化rnd
100
    RandomNumx = Int(randomnum * Rnd) + 1
'    If listnum > 1 Then
'        For i = 1 To listnum
'            If RandomNumx = arrTemp(i) Then GoTo 100 '出现重复的就重新执行
'        Next
'    End If
End Function

'i=%E5%B0%91%E5%A5%B3&from=zh-CHS&to=en&smartresult=dict&client=fanyideskweb&salt=15867649459408&sign=0a6bcd1042be82297221c2dd1478525c&ts=1586764945940&bv=9d1e6a4f9d4241fb7947f623cc9e4efa&doctype=json&version=2.1&keyfrom=fanyi.web&action=FY_BY_REALTlME
Sub 按钮2_Click()                                                    '对suwenkai老师的代码进行了补充，感谢 much 的帮助
'Dim pd$, exp$, i%, j%
'Dim dic As New Collection
'Dim a As Variant
'"美国股市暴跌,每个家庭的财富都将出现严重的损失.美国总统说:" &
''exp = ""
'pd = "&type=AUTO2AUTO"
'pd = "&i=" & Replace(Replace(Replace(encodeURIx("美国经济危机"), "%", "%C2%"), "%C2%E", "%C3%A"), "%C2%20", " ") & pd
'pd = pd & "&doctype=json"
'Debug.Print pd
'fanyideskweb+加拿大+15858962462879+Nw(nmmbP%A-r6U3EUn]Aj


'strx = CStr(Val(GetLongTime) + 125)

Set oDom = CreateObject("htmlfile")
Set oWin = oDom.parentwindow
oWin.execScript
str1 = "美国申请失业金人数连续第四周回落"
strx = oWin.eval("new Date().getTime()")
a = oWin.eval("parseInt(10*Math.random(),10)")
strx1 = strx & a
strx2 = "fanyideskweb" & str1 & strx1 & "Nw(nmmbP%A-r6U3EUn]Aj"
strx3 = LCase(GetMD5Hash_String(strx2))
strx4 = Application.EncodeUrl(str1) 'from=zh-CHS&to=en '



'strx4 = Replace(Replace(Replace(encodeURIx(str1), "%", "%C2%"), "%C2%E", "%C3%A"), "%C2%20", " ")
'strx5 = "3aabbc1a31e864bb89725aa04c217a5c"
'pd = "i=" & strx4 & "&from=zh-CHS&to=en&smartresult=dict&client=fanyideskweb&salt=" & strx1 & "&sign=" & strx3 & "&ts=" & strx & _
'"&bv=" & strx5 & "&doctype=json&version=2.1&keyfrom=fanyi.web&action=FY_BY_REALTlME"

'strx4 = Replace(Replace(Replace(encodeURIx(str1), "%", "%C2%"), "%C2%E", "%C3%A"), "%C2%20", " ")

PostData = "i=" & strx4 & "&from=zh-CHS&to=en" & "&smartresult=dict&client=fanyideskweb&salt=" & strx1 & "&sign=" & strx3 & "&ts=" & strx & _
"&bv=3aabbc1a31e864bb89725aa04c217a5c&doctype=json&version=2.1&keyfrom=fanyi.web&action=FY_BY_REALTlME"


'inputtext=%E7%BE%8E%E5%9B%BD%E5%A4%A7%E9%80%89&type=ZH_CN2JA
'pd = UTF8_URLEncoding(pd)
'pd = "inputtext=" & Application.EncodeUrl(str1) & "&type=ZH_CN2JA"

' postdata = "{" & Chr(34) & "i" & Chr(34) & ": " & Chr(34) & strx4 & Chr(34) & "," & _
'            Chr(34) & "from" & Chr(34) & ": " & Chr(34) & "zh-CHS" & Chr(34) & "," & _
'            Chr(34) & "to" & Chr(34) & ": " & Chr(34) & "en" & Chr(34) & "," & _
'            Chr(34) & "smartresult" & Chr(34) & ": " & Chr(34) & "dict" & Chr(34) & "," & _
'            Chr(34) & "client" & Chr(34) & ": " & Chr(34) & "fanyideskweb" & Chr(34) & "," & _
'            Chr(34) & "salt" & Chr(34) & ": " & Chr(34) & strx1 & Chr(34) & "," & _
'            Chr(34) & "sign" & Chr(34) & ": " & Chr(34) & strx3 & Chr(34) & "," & _
'            Chr(34) & "ts" & Chr(34) & ": " & Chr(34) & strx & Chr(34) & "," & _
'            Chr(34) & "bv" & Chr(34) & ": " & Chr(34) & "3aabbc1a31e864bb89725aa04c217a5c" & Chr(34) & "," & _
'            Chr(34) & "doctype" & Chr(34) & ": " & Chr(34) & "json" & Chr(34) & "," & _
'            Chr(34) & "version" & Chr(34) & ": " & Chr(34) & "2.1" & Chr(34) & "," & _
'            Chr(34) & "keyfrom" & Chr(34) & ": " & Chr(34) & "fanyi.web" & Chr(34) & "," & _
'            Chr(34) & "action" & Chr(34) & ": " & Chr(34) & "FY_BY_REALTlME" & Chr(34) & "}"



    url = "http://fanyi.youdao.com/translate_o?smartresult=dict&smartresult=rule" '"https://fanyi.baidu.com/gettts?lan=en&text=minimize&spd=3&source=web",http://fanyi.youdao.com/translate_o?smartresult=dict&smartresult=rule&smartresult=ugc&sessionFrom=null" '"http://fanyi.youdao.com/translate?smartresult=dict&smartresult=rule" ',"http://fanyi.youdao.com/translate_o?smartresult=dict&smartresult=rule"

    Set x = CreateObject("WinHttp.WinHttpRequest.5.1")
    With x
        .Option(6) = False
        .Open "POST", url, False
        .setRequestHeader "Host", "fanyi.youdao.com"
        .setRequestHeader "Accept", "application/json, text/javascript"
'        .setRequestHeader "X-Requested-With", "XMLHttpRequest"
'        .setRequestHeader "Cache-Control", "no-cache"
        .setRequestHeader "Origin", "http://fanyi.youdao.com"
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.169 Safari/537.36"
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
        .setRequestHeader "Referer", "http://fanyi.youdao.com/"
        .setRequestHeader "Cookie", "OUTFOX_SEARCH_USER_ID=325673277@27.38.20.19" '"OUTFOX_SEARCH_USER_ID=-1135518799@27.38.20.197"
'        .setProxy 2, "localhost:8888"
        .Send (PostData)
        If .Status <> 200 Then Stop
        Debug.Print .responseText
    End With
    Set x = Nothing
    Set oDom = Nothing
    Set oWin = Nothing
 End Sub
 
' Sub main()
'    Set oDom = CreateObject("htmlfile")
'    Set oWin = oDom.parentWindow
'    oWin.execScript
'    jsonpCallback = oWin.eval("'jsonpCallback' + Math.floor(Math.random() * (100000 + 1))")
'    t = oWin.eval("new Date().getTime()")
'End Sub
Function encodeURIx(ByVal strText As String) As String
    With CreateObject("msscriptcontrol.scriptcontrol")
        .Language = "JavaScript"
        encodeURIx = .eval("encodeURI('" & Replace(strText, "'", "\'") & "');")
    End With
End Function

Function getUnixTime()  '获取Unix时间戳
    getUnixTime = DateDiff("s", "01/01/1970 00:00:00", Now())
End Function


Sub kdkld()
Debug.Print GetLongTime
End Sub

Function GetLongTime() As String
    With CreateObject("msscriptcontrol.scriptcontrol")
        .Language = "JavaScript"
        GetLongTime = .eval("new Date().getTime();")
    End With
End Function
