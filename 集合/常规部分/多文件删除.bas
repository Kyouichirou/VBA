Attribute VB_Name = "多文件删除"
Option Explicit

Sub MFilesDele()   '多个文件删除
    Dim yesno As Variant
    Dim Arrow() As Integer, arr() As Integer
    Dim slc As Variant, slr As Variant
    Dim i As Integer, k As Integer, p As Integer, tfile As String, excodex As Integer, j As Integer
    
    On Error GoTo 100
    With ThisWorkbook.Sheets("书库")    '只有当选择的数量大于1, 选定区域在c列时,才进行操作
        k = Selection.Cells.Count
        If k < 2 Then
            .Label1.Caption = "选择数量少于2"
            Exit Sub
        ElseIf k > 10 Then
            .Label1.Caption = "选择数量超出范围"
            Exit Sub
        End If
        For Each slc In Selection.Columns              '判断所选择的区域是否满足要求,限制要求超过两行,限制选择c列
            If slc.Column <> 2 Then
                .Label1.Caption = "选择区域超出范围,请选择c列"
                Exit Sub
            End If
        Next
        i = 1
        ReDim Arrow(1 To k)     '获取不连续行的行号
        For Each slr In Selection.rows
            j = slr.Row
            If j < 6 Then
                .Label1.Caption = "选择文件有误,请重新操作"                '防止误操作
                Exit Sub
            End If
            Arrow(i) = j ''将选定区域的行号放进数组中暂时保存
            i = i + 1
        Next
        '---------------------------------------------------------------------------选择区域判断
        ReDim arr(1 To k)
        arr = Down(Arrow)                                         '删除行要采用倒序删除,否则会出现乱序的问题,把数组的值按降序进行重新排列
        yesno = MsgBox("是否删除本地文件?_", vbYesNo) '是否删除文件
        If yesno = vbYes Then '文件存在且执行删除命令
            excodex = 1
        Else
            excodex = 0
        End If
        For p = 1 To k
            tfile = .Range("e" & arr(p))
            If Len(.Cells(arr(p), "ab")) > 0 Then p = 1
            If excodex = 1 Then
                FileDeleExc tfile, arr(p), p, 0, .Cells(arr(p), "d") '执行删除本地文件
            Else
                DeleMenu arr(p) '移除书库
            End If
        Next
    End With
100
End Sub

Function Down(xi() As Integer) As Integer()  '降序函数
    Dim i As Integer, j As Integer, a As Integer, d() As Integer, m As Byte, n As Byte
    
    m = LBound(xi)
    n = UBound(xi)
    ReDim d(m To n)
    ReDim Down(m To n)
    d = xi
    If m = n Then
        Down = d
        Exit Function '只有一个数据
    End If

    For i = m To n - 1
        For j = i + 1 To n
            If d(j) > d(i) Then
                a = d(j): d(j) = d(i): d(i) = a
            End If
        Next
    Next
    Down = d
End Function

Private Function Downx(xi() As Integer) As Integer() '备用-使用sortlist来实现排序
    Dim sortlist As Object
    Dim i As Integer, k As Byte, arrTemp() As Integer
    'https://docs.microsoft.com/en-us/dotnet/api/system.collections.arraylist?view=netframework-4.8
    i = UBound(xi())
    ReDim arrTemp(1 To i)
    ReDim Downx(1 To i)
    Set sortlist = CreateObject("System.Collections.ArrayList") '注意区别createobject("System.Collections.SortedList")'ArrayList
    If sortlist Is Nothing Then MsgBox "无法创建对象": Exit Function
    With sortlist
        For k = 1 To i
            .Add xi(k)
        Next
        .sort
        i = i - 1
        For k = 0 To i
            arrTemp(k + 1) = sortlist(k)
        Next
    End With
    Downx = arrTemp
    Set sortlist = Nothing
End Function
