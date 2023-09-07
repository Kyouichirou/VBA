Attribute VB_Name = "低频率"
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal Index As Long) As Long
Private Const HORZRES = 8
Private Const VERTRES = 10
Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90

Function CreateWorksheet(ByVal dfilepath As String) '创建用于存储信息的表格
    Dim wb As Workbook
    Dim dicpath As String
    Dim wbdc As Workbook
    
    Set wb = Workbooks.Add
    With wb.Worksheets
        .Add after:=wb.Worksheets(.Count), Count:=6 - .Count '创建6张表
    End With
    With wb
        .Worksheets(1).Name = "打开记录"                                                                '表中分别写入表头
        .Worksheets(1).Range("a1:f1") = Array("统一编码", "文件名", "主文件名", "标识编码", "时间", "星期")
        .Worksheets(2).Name = "摘要记录"
        .Worksheets(2).Range("a1:f1") = Array("统一编码", "文件名", "主文件名", "标识编码", "时间", "内容")
        .Worksheets(3).Name = "备忘录"
        .Worksheets(3).Range("a1:c1") = Array("日期", "时间", "内容")
        With .Worksheets(4)
            .Name = "删除备份"
            ThisWorkbook.Sheets("书库").Range("b5:ag5").Copy .Cells(1, 1)
            .Cells(1, 33) = "删除原因"
            .Cells(1, 34) = "删除备注"
        End With
        .Worksheets(5).Name = "词库"
        .Worksheets(5).Range("a1:m1") = Array("编号", "英文", "音标", "中文", "自定义", "释义", "分类", "查询次数", "重要程度", "添加时间", "来源", "参考信息源", "生词本")
        .Worksheets(6).Name = "单词"
    End With
    wb.SaveAs dfilepath '路径进行非ansi字符检查?
    dicpath = ThisWorkbook.Path & "\单词表.xlsx"
    If fso.fileexists(dicpath) = True Then '单词表创建
        Set wbdc = Workbooks.Open(dicpath)
        Workbooks("单词表.xlsx").Sheets("词汇表").Cells.Copy Workbooks("lbrecord.xlsx").Sheets("单词").Range("a1") '将单词表的内容复制到表中
        wbdc.Close True
        Set wbdc = Nothing
    End If
    wb.Close savechanges:=True
    Set wb = Nothing
End Function

Sub AtClock() '动态时间 '不建议在窗体中执行多个时间sub-备用
'If timest = 1 Then
''DoEvents
'UserForm3.TextBox9.text = Format(Now, "yyyy-mm-dd HH:MM:SS")
'Application.OnTime Now + TimeValue("00:00:01"), "Atclock" '自我调用可以避免采用do循环造成的cpu占用问题
'End If
End Sub

Sub Deletempfile() '移除不需要文件-备用
'Dim arr()
'Dim i As Integer, k As Integer
'With ThisWorkbook.Sheets("书库")
'If .Range("b6") = "" Then Exit Sub
'If .[b65536].End(xlUp).Row > 6 Then
'arr = .Range("c6:d" & .[b65536].End(xlUp).Row).Value
'Else
'ReDim arr(1 To 1, 1 To 2)
'arr(1, 1) = .Range("c6").Value
'arr(1, 2) = .Range("d6").Value
'End If
'k = .[b65536].End(xlUp).Row - 5
'For i = k To 1 Step -1
'If Not UCase(arr(i, 2)) Like "EPUB" And Not UCase(arr(i, 2)) Like "PDF" And Not UCase(arr(i, 2)) Like "MOBI" And Not UCase(arr(i, 2)) Like "DO*" And Not UCase(arr(i, 2)) Like "XL*" And Not UCase(arr(i, 2)) Like "PP*" And Not UCase(arr(i, 2)) Like "AC*" And Not UCase(arr(i, 2)) Like "TX*" Then
'.Rows(i + 5).Delete Shift:=xlShiftUp
'Else
'   If arr(i, 1) Like "~$*" Then .Rows(i + 5).Delete Shift:=xlShiftUp
'End If
'Next
'End With
End Sub

Sub CheckAllFile() '检查文件的存在                 '全部执行判断目录下的文件是否存在
    Dim arre() As String
    Dim Elow As Integer, i As Integer
    
    With ThisWorkbook.Sheets("书库") '这里涉及到数组赋值时，只有一个值得问题
        Elow = .[e65536].End(xlUp).Row
        If Elow < 6 Then
        .Label1.Caption = "无数据"
        Exit Sub      '数据为空
        End If
        If Elow > 100 Then UserForm6.Show 0
        Call PauseRm '禁用selection事件
        If Elow > 6 Then
            arre = .Range("e6:e" & Elow).Value
            For i = 1 To Elow - 5
                If fso.fileexists(arre(i, 1)) = False Then
                    Call Delefile(arre(i, 1), i + 5, 2)
                End If
            Next
        ElseIf Elow = 6 Then
            If fso.fileexists(.Range("e6").Value) = False Then Call Delefile(.Range("e6").Value, 6, 2)
        End If
        Unload UserForm6
        .Label1.Caption = "执行完毕"
    End With
    Call EnableRm '启用右键事件
    ThisWorkbook.Save
End Sub

Function PSexist() As Boolean '判断powershell 是否存在 '扩展一下
    If ShellxExist = 1 Then PSexist = True: Exit Function
    If Len(ThisWorkbook.Sheets("temp").Range("ab4").Value) > 0 Then
        PSexist = True
    Else
        PSexist = False
    End If
End Function

Function CreateFolder(ByVal Folderpath As String, ByVal cmCode As Byte) As Boolean '创建文件夹/文档目录
    Dim i As Byte, xi As Variant, k As Byte, yesno As Variant, wt As Integer, strx As String, strx1 As String, m As Byte, j As Byte, n As Byte
    Dim strx2 As String
    '以下内容是根据表中的数据进行创建文件夹的,如果修改表中的数据,以下的内容也需要修改
    CreateFolder = True
    xi = Split(Folderpath, "\")
    i = UBound(xi)
    If Len(xi(i)) > 0 Then Folderpath = Folderpath & "\" '非根目录
    If cmCode = 1 Then
        strx = "Library"
        strx1 = "*[a-zA-Z]*"
        m = 3
    Else
        strx = "藏书"
        strx1 = "*[一-龥]*"
        m = 2
    End If
    Folderpath = Folderpath & strx
    
    If fso.folderexists(Folderpath) = True Then
        yesno = MsgBox("文件夹已存在,是否重新创建(原文件将被删除!)", vbYesNo, "Warning")
        If yesno = vbNo Then
            CreateFolder = False
            Exit Function
        Else
            wt = 200
            If fso.GetFolder(Folderpath).Size > 1048576000 Then wt = 350
            fso.DeleteFolder (Folderpath)
            Sleep wt
        End If
    End If
    fso.CreateFolder (Folderpath)
    Folderpath = Folderpath & "\"
    strx2 = Folderpath
    With ThisWorkbook.Sheets("报表")
        k = .[a65536].End(xlUp).Row
        j = 3
        n = j
100
            strx = .Cells(n, m).Value
            If strx Like strx1 Then
                fso.CreateFolder (strx2 & strx)
                Folderpath = strx2 & strx & "\"
            End If
            If n = k Then Exit Function
        For j = n To k
            strx = .Cells(j + 1, m + 1).Value
            If strx Like strx1 Then
                fso.CreateFolder (Folderpath & strx)
            Else
                n = j + 1
                GoTo 100
            End If
        Next
    End With
End Function

Sub ScreenDetail()
    Dim hDC As Long
    Dim x As Long, y As Long
    Dim x1 As Long, Y1 As Long
    hDC = GetDC(0)
    x = GetDeviceCaps(hDC, HORZRES)
    y = GetDeviceCaps(hDC, VERTRES)
    MsgBox "当前系统的屏幕分辨率为" & x & "X" & y
    x1 = GetDeviceCaps(hDC, LOGPIXELSX)
    Y1 = GetDeviceCaps(hDC, LOGPIXELSY)
    MsgBox "当前显示器的PPI为" & x1 & "X" & Y1
    ReleaseDC 0, hDC
End Sub

Sub ScreenDetaila() 'wps支持
    Dim WMIObject As Object
    Dim WMIResult As Object
    Dim WMIItem As Object
    Set WMIObject = GetObject("winmgmts:\\.\root\WMI")
    Set WMIResult = WMIObject.ExecQuery("Select * From WmiMonitorBasicDisplayParams")
    Dim Diagonal As Double
    Dim Width As Double
    Dim Height As Double
    Dim Counter As Byte
    Counter = 1
    For Each WMIItem In WMIResult
        Width = WMIItem.MaxHorizontalImageSize / 2.54
        Height = WMIItem.MaxVerticalImageSize / 2.54
        Diagonal = Sqr((Height ^ 2) + (Width ^ 2))
        MsgBox "Your monitor # " & Counter & " is approximiately " & Round(Diagonal, 2) & " inches diagonal"
        Counter = Counter + 1
    Next
    Set WMIObject = Nothing
    Set WMIResult = Nothing
End Sub

