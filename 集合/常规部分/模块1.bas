Attribute VB_Name = "ģ��1"
'Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'ʱ�� -����ѵ��
Private Declare Function timeGetTime Lib "winmm.dll" () As Long 'ʱ�� -����ѵ��
Sub djkloo()
BookRead "http://www.yuedu88.com/search.php?q=%E7%99%BD%E5%A4%9C%E8%A1%8C"
End Sub
Private Function BookRead(ByVal url As String) As String() '��ȡ�����������/����,����
    Dim myreg As Object, match As Object, Matches As Object
    Dim IE As Object, i As Byte
    Dim strx As String
    Dim t As Long
    Dim arr() As String
    Dim fl As Object, FilePath As String

    Set IE = CreateObject("InternetExplorer.Application") '����ie���������
    IE.Visible = False
    IE.Navigate url
    t = timeGetTime
    Do While IE.readyState <> 4 And timeGetTime - t < 7000 '�ȴ��ʱ�䲻����7s,��������ʱ�����, ie�򿪶���, ��ֹ����ҳ������γɵ���ѭ��
        DoEvents
        Sleep 25
    Loop
    If IE.Busy = True Then IE.Stop 'ֹͣ����������ҳ
    With IE.Document
        strx = .getElementById("BookText").InnerHtml '����+����
    End With
    IE.Quit
    Set IE = Nothing
Dim fl As Object, FilePath As String
    FilePath = "C:\Users\adobe\Downloads\temp.txt"
    Set fl = fso.OpenTextFile(FilePath, ForWriting, True, TristateUseDefault)
    fl.WriteLine strx
    fl.Close
    Set fl = Nothing
'    Set fl = fso.OpenTextFile(filepath, ForReading)
'    Do
'    strx = fl.ReadLine
'    Set myreg = CreateObject("VBScript.RegExp") '������ʽ
'    With myreg
'        .Pattern = "[\>]+(.+?)+[\<]"
'        .Global = True
'        .IgnoreCase = True '�����ִ�Сд
'        Set matches = .execute(strx)
'        i = matches.Count
'        If i = 0 Then
'            ThisWorkbook.Application.Speech.Speak ("δ��������")
'        Else
''        ReDim arr(i)
'        For Each match In matches
''            arr(i) = match.value
''            i = i + 1
'
'ThisWorkbook.Application.Speech.Speak ("δ��������")
'        Next
'        End If
'    End With
''    Loop
'    Set fl = Nothing
'    Set match = Nothing
'    Set matches = Nothing
End Function













Sub md()
Dim fl As TextStream
Dim strx As String

Set fl = fso.OpenTextFile("C:\Users\adobe\Downloads\temp.txt", ForReading, True, TristateUseDefault)
'Do While fl.AtEndOfLine = True
strx = fl.ReadAll
Debug.Print strx
'Loop
fl.Close

Set fl = Nothing
End Sub
