Attribute VB_Name = "模块1"
'Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) '时间 -单词训练
Private Declare Function timeGetTime Lib "winmm.dll" () As Long '时间 -单词训练
Sub djkloo()
BookRead "http://www.yuedu88.com/search.php?q=%E7%99%BD%E5%A4%9C%E8%A1%8C"
End Sub
Private Function BookRead(ByVal url As String) As String() '获取豆瓣封面链接/国籍,作者
    Dim myreg As Object, match As Object, Matches As Object
    Dim IE As Object, i As Byte
    Dim strx As String
    Dim t As Long
    Dim arr() As String
    Dim fl As Object, FilePath As String

    Set IE = CreateObject("InternetExplorer.Application") '创建ie浏览器对象
    IE.Visible = False
    IE.Navigate url
    t = timeGetTime
    Do While IE.readyState <> 4 And timeGetTime - t < 7000 '等待最长时间不超过7s,必须增加时间控制, ie打开豆瓣, 防止出现页面假死形成的死循环
        DoEvents
        Sleep 25
    Loop
    If IE.Busy = True Then IE.Stop '停止继续加载网页
    With IE.Document
        strx = .getElementById("BookText").InnerHtml '作者+国籍
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
'    Set myreg = CreateObject("VBScript.RegExp") '正则表达式
'    With myreg
'        .Pattern = "[\>]+(.+?)+[\<]"
'        .Global = True
'        .IgnoreCase = True '不区分大小写
'        Set matches = .execute(strx)
'        i = matches.Count
'        If i = 0 Then
'            ThisWorkbook.Application.Speech.Speak ("未发现数据")
'        Else
''        ReDim arr(i)
'        For Each match In matches
''            arr(i) = match.value
''            i = i + 1
'
'ThisWorkbook.Application.Speech.Speak ("未发现数据")
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
