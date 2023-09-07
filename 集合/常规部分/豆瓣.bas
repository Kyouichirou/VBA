Attribute VB_Name = "����"
Option Explicit
'private const
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'ʱ�� -����ѵ��
Private Declare Function timeGetTime Lib "winmm.dll" () As Long 'ʱ�� -����ѵ��
Private Const BdUrl As String = "https://www.baidu.com"
Private Const ByUrl As String = "https://cn.bing.com/search?q=site:book.douban.com%20"
'-------������ҳ����,ʹ������ĺô�
'����ƥ�临�ӵĹ���-��Ϊ��ҳ��������Ϣ�ǳ�����,����ͨ��split,left�Ⱥ����������������鷳,�����ڴ�����Ϣ���ϴ������ʱ,�ٶȽ���

Function DoubanBook(ByVal strWord As String) '��ȡ������������'���ڶ�����鼮����Ϣ�����˼��ܴ��������Ӧ����ȡ����(360����Ҳ�ܹ���ץȡ���������)
    Dim strx As String, url As String
    Dim xtemp As Variant, arr() As String
    Dim t As Long
    
    With UserForm3
        If TestURL(BdUrl) Then '������������Ƿ�������״̬
            url = ByUrl & encodeURI(strWord) & "&count=1" 'ָ����Ӧȥ��������������Ϣ
            '------------------------------------------------------------------------------------count=1Ϊ�����������,��ʾ����һ���������
            'site:book.douban.com-��ʾָ����������������Ϣ,%20��ʾ "+"
            '-------------------------------------��Ϊ���������ݻ��ж�����Ž��,���Է��ص�һ���,׼ȷ�Ȳ�һ����
            With CreateObject("MSXML2.XMLHTTP")
                .Open "GET", url, True
                .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                .Send
                t = timeGetTime
                Do While .readyState <> 4 And timeGetTime - t < 5000 '�ȴ����ݵķ���
                    DoEvents
                Loop
                strx = .responseText
            End With
'                If InStr(strx, "https://book.douban.com/review/") > 0 Then GoTo 100 '��Ҫ��һ��ϸ�����صĽ��-���־ɵ�ҳ����õ���5���Ƶ�����,�°汾������Ϊ10����
                If InStr(strx, "�û�����") > 0 Then '��ʾ��ȡ����Ч���������
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
                    .Label56.Caption = "" '�����Ϣ
                    .Label57.Caption = "�����ɹ�"
                Else
100
                    .Label57.Caption = "δ�ҵ��鼮��Ϣ"
                    .Label56.Caption = "δ�ҵ��鼮��Ϣ"
                End If
        Else
            .Label57.Caption = "���������쳣"
        End If
    End With
End Function

Private Function DoubanData(ByVal xtext As String) As String()
    Dim myreg As Object, match As Object, Matches As Object
    Dim arr() As String
    Dim arreg(), i As Byte
    '------------------------------------------------------https://tool.oschina.net/regex/,������ʽ���߲���
    '----------------------------------------------------- https://www.runoob.com/regexp/regexp-metachar.html ����
    ReDim DoubanData(2)
    ReDim arr(2)
    ReDim arreg(2)
    arreg = Array("�û�����:+(.+?)(\<)", "[\>]+(.+?)[\(]����[\)]", "[https]+://+book.douban.com/+[a-z]*\/+[0-9]*\/")
    Set myreg = CreateObject("VBScript.RegExp") '������ʽ
    For i = 0 To 2
        With myreg
            .Pattern = arreg(i) '��ȡ��������
            .Global = True
            .IgnoreCase = True '�����ִ�Сд
            Set Matches = .Execute(xtext)
            For Each match In Matches
                arr(i) = match.Value: Exit For
            Next
            If i = 1 Then xtext = arr(i) '���δ�����
        End With
        Set match = Nothing
        Set Matches = Nothing
    Next
    arr(0) = Trim(Replace(Replace(arr(0), "�û�����:", ""), "<", "")) '��ȡ�û�����
    If Len(arr(1)) > 0 Then arr(1) = Trim(Right$(arr(1), Len(arr(1)) - InStrRev(arr(1), ">")))
    DoubanData = arr
    Set myreg = Nothing
End Function

Function ObtainDoubanPicture(ByVal url As String) As String() '��ȡ�����������/����,����
    Dim myreg As Object, match As Object, Matches As Object
    Dim IE As Object, i As Byte
    Dim strx As String, strx1 As String, strx2 As String
    Dim arreg(), arr(2) As String, arrTemp(1)
    Dim t As Long
    '---------------https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa752093(v=vs.85)?redirectedfrom=MSDN
    Set IE = CreateObject("InternetExplorer.Application") '����ie���������
    IE.Visible = False
    IE.Navigate url
    t = timeGetTime
    Do While IE.readyState <> 4 And timeGetTime - t < 7000 '�ȴ��ʱ�䲻����7s,��������ʱ�����, ie�򿪶���, ��ֹ����ҳ������γɵ���ѭ��
        DoEvents
    Loop
    If IE.Busy = True Then IE.Stop 'ֹͣ����������ҳ
    With IE.Document
        arrTemp(0) = .getElementById("mainpic").InnerHtml '����+����
        arrTemp(1) = .getElementById("info").InnerHtml '����+����
    End With
    ReDim arreg(1)
    arreg = Array("[a-zA-z]+://[^\s]*", "[\>]+(\s.+?)+[\<]\/a[\>]")
    Set myreg = CreateObject("VBScript.RegExp") '������ʽ
    IE.Quit
    Set IE = Nothing
    For i = 0 To 1
        With myreg
            .Pattern = arreg(i) '��ȡ����ͼƬ����
            .Global = True
            .IgnoreCase = True '�����ִ�Сд
            Set Matches = .Execute(arrTemp(i))
            For Each match In Matches
                arr(i) = match.Value: Exit For
            Next
        End With
        Set match = Nothing
        Set Matches = Nothing
    Next
    strx1 = Trim(arr(0))
    arr(0) = Left$(strx1, Len(strx1) - 2) 'ͼƬ����
    strx2 = arr(1)
    strx2 = Trim(Replace(Replace(strx2, "</a>", ""), ">", ""))
    strx2 = Trim(Replace(strx2, Chr(10), "", 1, 2))
    If InStr(strx2, "[") > 0 And InStr(strx2, "]") > 0 Then
        arr(2) = AuthorNT(strx2) '����
        strx2 = Trim(Right(strx2, Len(strx2) - Len(arr(2))))
    End If
    arr(1) = strx2 '����
    ReDim ObtainDoubanPicture(2)
    ObtainDoubanPicture = arr
    Set myreg = Nothing
End Function

Function DoubanTreat(ByVal gradex As String, ByVal authorx As String, ByVal bookx As String) As String() '������Ϣ����
    Dim strx As String, strx1 As String, strx2 As String, strx3 As String, arr() As String
    Dim strx4 As String
    
    On Error Resume Next
    ReDim DoubanTreat(1 To 5)
    ReDim arr(1 To 5)
    If InStr(gradex, "v:average") > 0 Then
        strx = Split(Split(gradex, "v:average")(1), "</")(0)
        arr(1) = Trim(Mid$(strx, 3, Len(strx) - 2)) '����
    End If
    If InStr(authorx, "</a>") > 0 Then
        strx1 = Split(authorx, "</a>")(0)
        strx4 = Trim(Replace(Trim(Right$(strx1, Len(strx1) - InStrRev(strx1, ">"))), Chr(34), "", 1, 2)) '����
        strx4 = Trim(Replace(strx4, Chr(34), "", 1, 2))
        strx4 = Trim(Replace(strx4, Chr(10), "", 1, 2)) 'chr(10)���з�,ע���vbcr�ȳ�����������
        If InStr(strx4, "[") > 0 And InStr(strx4, "]") > 0 Then '[��]��Ұ����
            arr(5) = AuthorNT(strx4) '���߹���
            strx4 = Trim(Right(strx4, Len(strx4) - Len(arr(5))))
        End If
        arr(2) = strx4 '����
    End If
    
    If InStr(bookx, "title=") > 0 Then
        strx2 = Split(Split(bookx, "title=")(1), Chr(34))(1)
        arr(3) = Trim(strx2)
    End If
    If InStr(bookx, "src") > 0 Then
        strx3 = Trim(Split(Split(bookx, "src")(1), " ")(0))
        arr(4) = Trim(Replace(Trim(Mid$(strx3, 2, Len(strx3) - 1)), Chr(34), "", 1, 2)) '����
    End If
    DoubanTreat = arr
End Function

Private Function AuthorNT(ByVal textx As String) As String '����-��ȡ���ߵĹ���-������ʽ
    Dim myreg As Object, match As Object, Matches As Object
    '------------------------------------------------------https://tool.oschina.net/regex/,������ʽ���߲���
    '----------------------------------------------------- https://www.runoob.com/regexp/regexp-metachar.html ����
    Set myreg = CreateObject("VBScript.RegExp") '������ʽ
    With myreg
        .Pattern = "(\[)+(.+?)+(\])" 'ƥ��[��]�������ݵ�������������"["����"]",Ҳ��ƥ��"[]",Ҳ��ƥ��"[��"����"��]" ,\[��ʾת��, ����Ŀ����"["�������
        .Global = True
        .IgnoreCase = True '�����ִ�Сд
        Set Matches = .Execute(textx)
        For Each match In Matches
            AuthorNT = match.Value: Exit For
        Next
    End With
    Set myreg = Nothing
    Set match = Nothing
    Set Matches = Nothing
End Function
