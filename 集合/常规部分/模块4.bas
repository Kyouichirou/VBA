Attribute VB_Name = "ģ��4"
'Option Explicit
Private Declare Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Sub dkkxxld()
'Set sapiX = CreateObject("SAPI.SpVoice")
'sapiX.Volume = 100 '����
'sapiX.Rate = 0 '�������� Խ��Խ��
'FlagsAsync = 1 'ͬ�����첽��1���첽
''���������ѡϵͳ�Ѱ�װ�����������Բ����У��ÿ�������˹���ѡ�õ�
''set colVoice=sapiX.getVoices() '��װ�ж��ٸ��������Ͽ�ѡ
''set sapiX.Voice=colVoice(0) 'ѡ��1������
'
'strText = "I asked God for a bike, but I know God doesn't work that way."
'strText = strText & "So I stole a bike and asked for forgiveness."
'strText = strText & "��ʼ��ֱ�����ϵ۴������г�����������ĥ�ϵ۰��¶��������·���� ��������͵��һ��Ȼ�����ϵۿ�ˡ��"
'
'sapiX.Speak strText, FlagsAsync 'ͬ�����첽
'MsgBox "���Դ��ڣ�ע��ͬ�����첽Ч��"
'sapiX.Skip 20
''sapiX.Pause '��ͣSPEAK���첽��Ч�����MSGBOX����ȥ�����ڳ���һ�ξͿ���
''MsgBox "PAUSE����������"
''sapiX.Volume = 30
'sapiX.Resume
'Randomize (Timer)
'For i = 0 To 99
'Debug.Print RandomNumx(10)
'Next
''Debug.Print Application.EncodeUrl("&")
Debug.Print DeleteUrlCacheEntry("https://aip.baidubce.com/")
End Sub


Private Function RandomNumx(ByVal randomnum As Integer) As Integer '�����
    Dim RndNumber, i As Byte
    
    Randomize (Timer) '��ʼ��rnd
100
    RandomNumx = Int(randomnum * Rnd) + 1
'    If listnum > 1 Then
'        For i = 1 To listnum
'            If RandomNumx = arrTemp(i) Then GoTo 100 '�����ظ��ľ�����ִ��
'        Next
'    End If
End Function

'i=%E5%B0%91%E5%A5%B3&from=zh-CHS&to=en&smartresult=dict&client=fanyideskweb&salt=15867649459408&sign=0a6bcd1042be82297221c2dd1478525c&ts=1586764945940&bv=9d1e6a4f9d4241fb7947f623cc9e4efa&doctype=json&version=2.1&keyfrom=fanyi.web&action=FY_BY_REALTlME
Sub ��ť2_Click()                                                    '��suwenkai��ʦ�Ĵ�������˲��䣬��л much �İ���
'Dim pd$, exp$, i%, j%
'Dim dic As New Collection
'Dim a As Variant
'"�������б���,ÿ����ͥ�ĲƸ������������ص���ʧ.������ͳ˵:" &
''exp = ""
'pd = "&type=AUTO2AUTO"
'pd = "&i=" & Replace(Replace(Replace(encodeURIx("��������Σ��"), "%", "%C2%"), "%C2%E", "%C3%A"), "%C2%20", " ") & pd
'pd = pd & "&doctype=json"
'Debug.Print pd
'fanyideskweb+���ô�+15858962462879+Nw(nmmbP%A-r6U3EUn]Aj


'strx = CStr(Val(GetLongTime) + 125)

Set oDom = CreateObject("htmlfile")
Set oWin = oDom.parentwindow
oWin.execScript
str1 = "��������ʧҵ���������������ܻ���"
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

Function getUnixTime()  '��ȡUnixʱ���
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
