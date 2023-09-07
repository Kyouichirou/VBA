VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cRegex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'----------------����
'���������ȫ��,��ֻƥ�䵥��
Private oRegE As Object
'([hH][tT]{2}[pP]://|[hH][tT]{2}[pP][sS]://|[wW]{3}.|[wW][aA][pP].|[fF][tT][pP].|[fF][iI][lL][eE].)[-A-Za-z0-9+&@#/%?=~_|!:,.;]+[-A-Za-z0-9+&@#/%=~_|] ƥ����ַ

Private Sub Class_Initialize()
    Set oRegE = CreateObject("VBScript.RegExp")
End Sub

Private Sub Class_Terminate()
    Set oRegE = Nothing
End Sub
'--------------------------------------------��

'-----------ִ��
Friend Function ReplaceText(ByVal strText As String, ByVal sPattern As String, ByVal rpText As String, Optional ByVal IsGlobal As Boolean = True, Optional ByVal IgnoreC As Boolean = True) As String
    If isReady(strText, sPattern) = False Then Exit Function
    With oRegE
        .Pattern = sPattern
        .Global = IsGlobal
        .IgnoreCase = IgnoreC
        ReplaceText = .Replace(strText, rpText)
    End With
End Function

Friend Function sMatch(ByVal strText As String, ByVal sPattern As String, Optional ByVal IsGlobal As Boolean = True, Optional ByVal IgnoreC As Boolean = True) As String
    '--------------------------��ȡ����ƥ��ֵ
    Dim matches As Object
    Dim match  As Object
    
    If isReady(strText, sPattern) = False Then Exit Function
    With oRegE
        .Pattern = sPattern
        .Global = IsGlobal
        .IgnoreCase = IgnoreC
        Set matches = .Execute(strText)
        For Each match In matches
            sMatch = match.Value: Exit For
        Next
    End With
    Set matches = Nothing
End Function

Friend Function xMatch(ByVal strText As String, ByVal sPattern As String, Optional ByVal IsGlobal As Boolean = True, Optional ByVal IgnoreC As Boolean = True) As String()
    Dim i As Long
    Dim matches As Object
    Dim match As Object
    Dim arr() As String
    '-------------��ȡ���ƥ��ֵ
    If isReady(strText, sPattern) = False Then Exit Function
    With oRegE
        .Pattern = sPattern
        .Global = IsGlobal
        .IgnoreCase = IgnoreC
        Set matches = .Execute(strText)
    End With
    i = matches.Count - 1
    If i < 0 Then Exit Function
    ReDim xMatch(i)
    ReDim arr(i)
    i = 0
    For Each match In matches
        arr(i) = match.Value: i = i + 1
    Next
    xMatch = arr
    Set matches = Nothing
End Function

Friend Function xSubmatch(ByVal strText As String, ByVal sPattern As String, Optional ByVal IsGlobal As Boolean = True, Optional ByVal IgnoreC As Boolean = True) As String()
    Dim i As Long, p As Long, n As Long, m As Integer, j As Integer
    Dim matches As Object
    Dim match As Object
    Dim arr() As String
    '-------------��ȡ�������ƥ��ֵ, sPattern����Ҫƥ�������, ��������������
    'sPartten="songID":+[\d]{6,} ����("songID"):+([\d]{6,})
    If isReady(strText, sPattern) = False Then Exit Function
    With oRegE
        .Pattern = sPattern
        .Global = IsGlobal
        .IgnoreCase = IgnoreC
        Set matches = .Execute(strText)
    End With
    n = matches.Count - 1
    If n < 0 Then Exit Function
    ReDim arr(n, 0)
    ReDim xSubmatch(n, 0) 'һά���鲻���ٴ�redimΪ��ά
    j = 0
    For Each match In matches
        p = match.SubMatches.Count - 1
        If p > m Then m = p: ReDim Preserve arr(n, m): ReDim Preserve xSubmatch(n, m)
        For i = 0 To p
            arr(j, i) = match.SubMatches(i)
        Next
        j = j + 1
    Next
    xSubmatch = arr
    Set matches = Nothing
End Function

Friend Function cTest(ByVal strText As String, ByVal sPattern As String, Optional ByVal IsGlobal As Boolean = True, Optional ByVal IgnoreC As Boolean = True) As Boolean
    '----����Ƿ���ƥ��ֵ
    If isReady(strText, sPattern) = False Then Exit Function
    With oRegE
        .Pattern = sPattern
        .Global = IsGlobal
        .IgnoreCase = IgnoreC
        cTest = .Test(strText)
    End With
End Function

Private Function isReady(ByVal strText As String, ByVal sPattern As String) As Boolean '������������
    isReady = True
    If Len(Trim(strText)) = 0 Or Len(sPattern) = 0 Then isReady = False
End Function
