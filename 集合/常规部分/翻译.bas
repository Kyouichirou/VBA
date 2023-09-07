Attribute VB_Name = "����"
Option Explicit
Private Declare Function timeGetTime Lib "winmm.dll" () As Long 'ʱ��api���Ծ�ȷ������

Function GetdicMeaning(ByVal keywordx As String, ByVal excode As Byte) As String '��ȡ����
    Dim i As Integer
    Dim strx As String
    
    If excode = 1 Then 'Ӣ��
        strx = SearchWordFromYoudao(keywordx)
    ElseIf excode = 2 Or excode = 3 Then '����/���ڿո��Ӣ��
        strx = SearchWordFromCiba(keywordx)
    End If

    i = Len(Trim(strx))
    If i = 0 Then
        GetdicMeaning = "δ��ȡ����Ϣ"
        Exit Function
    ElseIf i > 150 Then
        GetdicMeaning = "���ص���Ϣ����ʧ��"
        Exit Function
    End If
    If strx = "xnothingx" Then GetdicMeaning = "�����쳣!": Exit Function
    GetdicMeaning = strx
End Function

Private Function SearchWordFromCiba(tmpWord As String) As String '��ɽ
     Dim XH As Object
     Dim url As String
     Dim s() As String
     Dim str_tmp As String
     Dim str_base As String
     Dim hytmp As String, hy As Variant, i As Integer, j As Integer, k As Integer
     Dim time1 As Long, strx As String
     
     tmpWord = Replace(tmpWord, " ", "%20") '�ո����ӷ���  " % "
     strx = encodeURI(tmpWord) 'ת�봦��
     '------------------------------------------------------------------office2013��֧�� Application.EncodeURL
     If strx = "xnothingx" Then SearchWordFromCiba = strx: Exit Function
     url = "http://www.iciba.com/" & strx
     '������ҳ
     Set XH = CreateObject("Microsoft.XMLHTTP")
     If XH Is Nothing Then UserForm3.Label57.Caption = "��������ʧ��": Exit Function
     On Error Resume Next
     XH.Open "get", url, True '0
     XH.Send (Null)
     On Error Resume Next
     time1 = timeGetTime
     While XH.readyState <> 4
         DoEvents
         If timeGetTime - time1 > 5000 Then '����
            XH.Close
            Set XH = Nothing
            Exit Function
         End If
     Wend
     str_base = XH.responseText
     XH.Close
     Set XH = Nothing
     
     '�����ĺ���ֽ�
     hy = Replace(Split(Split(str_base, "s="""">")(1), "</ul>")(0), "prop"">", ">" & Chr(10))
     hy = Split(hy, "<span")
     j = LBound(hy)
     k = UBound(hy)
     For i = j + 1 To k
         hytmp = hytmp & Split(Split(hy(i), ">")(1), "<")(0)  'vbCrLf &
     Next i
     
     SearchWordFromCiba = Mid$(hytmp, 2)
End Function

Function SearchWordFromBing(tmpWord As String) As String '��Ӧ----����
'    http://cn.bing.com/dict/search?q=about+to&go=%E6%8F%90%E4%BA%A4&qs=bs&form=CM
    'http://cn.bing.com/dict/search?q=about+to&go=�ύ&qs=bs&form=CM
    Dim XH As Object
    Dim s() As String
    Dim str_tmp As String, url As String, hytmp As String
    Dim str_base As String, hy As Variant
    Dim i As Integer, j As Integer, k As Integer
    
    tmpWord = Replace(tmpWord, " ", "+") '���ֿո�����
    url = "http://cn.bing.com/dict/search?q=" & tmpWord
    Set XH = CreateObject("Msxml2.XMLHTTP") 'Microsoft.XMLHTTP")
    If XH Is Nothing Then UserForm3.Label57.Caption = "��������ʧ��": Exit Function
    On Error Resume Next
    XH.Open "GET", url, 0 'True
    XH.Send '(Null)
     While XH.readyState <> 4
         DoEvents
     Wend
     str_base = XH.responseText
     XH.Close
     Set XH = Nothing
     '�����ĺ���ֽ�
     hy = Split(hy, "<span class=""pos"">")
     j = LBound(hy)
     k = UBound(hy)
     For i = j + 1 To k
         hytmp = hytmp & DelHtml(Split(hy(i), "</span></span>")(0)) & vbCrLf
     Next i
     If UBound(hy) = 0 Then hytmp = ""
     SearchWordFromBing = Left$(hytmp, Len(hytmp) - 1) '
End Function

Function SearchWordFromYoudao(ByVal tmpWord As String) As String '�е�
    'http://dict.youdao.com/search?q=����&keyfrom=dict.index
    Dim XH As Object
    Dim s() As String, i As Integer, j As Integer, k As Integer
    Dim str_tmp As String, url As String
    Dim str_base As String
    Dim tmpTrans As String, time1 As Long
    
    Set XH = CreateObject("Microsoft.XMLHTTP")
    If XH Is Nothing Then UserForm3.Label57.Caption = "��������ʧ��": Exit Function
'    tmpWord = Replace(tmpWord, " ", "%20") '���ֿո�����-��Ч

    url = "http://dict.youdao.com/search?q=" & tmpWord
    On Error Resume Next
    XH.Open "GET", url, True     '������ҳ
    XH.Send
    On Error Resume Next
    time1 = timeGetTime
    While XH.readyState <> 4
        DoEvents
        If timeGetTime - time1 > 5000 Then '����
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

Function encodeURI(strText As String) As String 'js�ַ�ת�� /��Ҫ�����Ĳ��ֽ���ת��
    Dim obj As Object
    '---------------------ExcelҲ�����õ�ת������ Application.EncodeURL
    Set obj = CreateObject("msscriptcontrol.scriptcontrol")
    If obj Is Nothing Then MsgBox "�ַ�ת��Sub�쳣", vbCritical, "Warning!!!": encodeURI = "xnothingx": Exit Function
    With obj
        .Language = "JavaScript"
        encodeURI = .eval("encodeURIComponent('" & strText & "');")
    End With
    Set obj = Nothing
End Function

Function DelHtml(ByVal strh As String) As String '������ȡ�ַ���
    Dim a As String
    Dim regEx As Object
    'Dim mMatch As Match
    'Dim Matches As matchcollection
    
    a = strh
    a = Replace(a, Chr(13) & Chr(10), "")
'    A = Replace(A, Chr(32), "")
    a = Replace(a, Chr(9), "")
    a = Replace(a, "</p>", vbCrLf)   '���������ϻس�
    Set regEx = CreateObject("vbscript.regexp")    '����������ʽ
    With regEx
        .Global = True
        .Pattern = "\<[^<>]*?\>"   '��<>��������html����
        .MultiLine = True  '������Ч
        .IgnoreCase = True  '���Դ�Сд(��ҳ����ʱ��������Ƚ���Ҫ)
        a = .Replace(a, "")   '��html����ȫ���滻Ϊ��
    End With
    a = Trim(a)
    
    '������Ŵ���
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






