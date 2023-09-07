Attribute VB_Name = "文件操作"
Option Compare Text                       '不区分大小写
Option Explicit
Dim filedyn As Boolean '判断本地文件被删除是否发生
Public Reasona As String, Reasonb As String, Filehashx As String
Public DeleFilex As Byte, MDeleFilex As Byte, AddPlistx As Byte '曾删改

Function FileDeleExc(ByVal FilePath As String, ByVal addrow As Integer, ByVal Px As Byte, ByVal cmCode As Byte, Optional ByVal filex As String, _
Optional ByVal FileName As String) As Boolean '执行删除 '文件路径,文件扩展名,所在行号,是否有非ansi,执行命令的来源
    Dim strx As String, cmCodex As Byte, k As Byte
    
    FileDeleExc = True
    filedyn = False '使用前初始化
    If cmCode = 0 Then           '删除命令来源于表格
        k = FileTest(FilePath, filex, FileName)
        If k >= 3 Then MsgShow "文件处于打开状态", "Warning", 1500: FileDeleExc = False: Exit Function
    End If
    If k = 0 Or cmCode = 1 Then
        On Error GoTo 100
        If Px > 0 Then
            fso.DeleteFile (FilePath)
        Else
            DeleteFiles (FilePath)
        End If
        filedyn = True
    End If
    Call DeleMenu(addrow) '删除目录
    Exit Function
100
    If Err.Number = 70 Then
        MsgBox "文件处于打开的状态"
    Else
        MsgBox "异常"
    End If
    FileDeleExc = False
    Err.Clear
End Function

Sub DeleMenu(ByVal addrow As Integer) '删除目录 'optional表示参数可写或者不写,位置要处于最后面,否知后面部分的参数全部都要写optional
    Dim addloc As String, i As Byte
    Dim arrback(1 To 34) As String
                                      '删除文件和删除目录分开
    On Error GoTo 100
    With ThisWorkbook.Sheets("书库")
        For i = 1 To 32
            arrback(i) = .Cells(addrow, i + 1).Value '将要删除的信息存储到数组
        Next
        '抹掉摘要的信息/删除备份
        arrback(33) = Reasona
        arrback(34) = Reasonb
        If Len(Filehashx) > 0 Then arrback(10) = Filehashx
       Call DeleOverBack(arrback(1), arrback(2), arrback(3), arrback(4), arrback(5), arrback(6), arrback(7), arrback(8), arrback(9), arrback(10), arrback(11), _
            arrback(12), arrback(13), arrback(14), arrback(15), arrback(16), arrback(17), arrback(18), arrback(19), arrback(20), arrback(21), arrback(22), arrback(23), _
            arrback(24), arrback(25), arrback(26), arrback(27), arrback(28), arrback(29), arrback(30), arrback(31), arrback(32), arrback(33), arrback(34))

        addloc = arrback(5)
        .rows(addrow).Delete Shift:=xlShiftUp
        Call DeleFileOver(addloc) '删除后执行检查目录和文件的存在状况
    End With
    If UF3Show = 3 Then DeleFilex = 1
100
    Reasona = ""
    Filehashx = ""
    Reasonb = "" '用完之后重置
End Sub

Function DeleFileOver(ByVal tfolder As String, Optional ByVal cmCode As Byte)
'-------------------------------------------------------------------------采用整体移除目录的方法,即文件的任何关联文件夹的文件都不存在了,目录中的数据才会全部移除 '删除文件的后续处理
    Dim rngad As Range, rngadx As Range
    Dim tfolderp As String
    Dim c As Byte, alow As Integer, dlow As Integer, i As Integer, flow As Integer, csc As Byte, blow As Integer
    Dim filec As Integer
    
    With ThisWorkbook
        .Application.ScreenUpdating = False
        .Application.Calculation = xlCalculationManual
        With .Sheets("书库")
            tfolderp = tfolder & "\"     '需要优化
            flow = .[f65536].End(xlUp).Row
            If cmCode = 0 Then
            Set rngadx = .Range("e6:e" & flow).Find(tfolderp, lookat:=xlPart) '检查文件夹是否还有其他的子文件夹文件存在目录
            End If
            c = UBound(Split(tfolder, "\")) '文件层级
            With ThisWorkbook.Sheets("目录")
                blow = .[b65536].End(xlUp).Row
                csc = .Cells.SpecialCells(xlCellTypeLastCell).Column
                If cmCode = 1 Then GoTo 401
                If Not rngadx Is Nothing Then
                    If filedyn = True Then                          '如果文件删除这个动作被执行
                        Set rngad = .Cells(4, 3).Resize(blow, csc).Find(tfolderp, lookat:=xlWhole) '精确搜索
                        If Not rngad Is Nothing And filedyn = True Then
                            rngad.Offset(0, 2) = Now '更新所在文件夹的修改时间
                            filec = rngad.Offset(0, 4).Value '更新文件夹数量
                            If IsNumeric(filec) = True Then
                                If filec > 1 Then
                                    filec = filec - 1
                                    rngad.Offset(0, 4) = filec
                                Else
                                    rngad.Offset(0, 4) = 0
                                End If
                            End If
                        End If
                    End If
                    Exit Function
                End If
                Set rngadx = Nothing
                Set rngad = Nothing
            End With
            '文件夹下文件已不存在目录
            With ThisWorkbook.Sheets("主界面")
                dlow = .[e65536].End(xlUp).Row
                If dlow < 37 Then GoTo 401
                Set rngad = .Range("e37:d" & dlow).Find(tfolder, lookat:=xlWhole) ''清除主文件夹
                If Not rngad Is Nothing Then
                    i = rngad.Row
                    If UF3Show = 3 Or UF3Show = 1 Then MDeleFilex = 1
                    If i = dlow Then
                        .Range("e" & i & ":" & "j" & i).ClearContents '如果是最后一行,就直接处理
                    Else
                        rngad.Delete Shift:=xlUp '删除单元格(不是清除内容)
                        rngad.Offset(0, 4).Delete Shift:=xlUp '删除添加时间
                    End If
                End If
                Do
                    Set rngad = .Range("d37:d" & dlow).Find(tfolderp, lookat:=xlPart) '将关联的文件目录全部移除,清除子文件夹
                    If Not rngad Is Nothing Then rngad.Delete Shift:=xlUp: rngad.Offset(0, 4).Delete Shift:=xlUp '采用直接删除的方式,所以不要合并区域的单元格,否则命令会出现警告弹窗
                Loop Until rngad Is Nothing '采用循环
            End With
401
            Set rngad = Nothing
            If c = 1 Then
                GoTo 100         '1级文件夹直接跳过二次筛选
            ElseIf c > 1 Then
                tfolderp = Split(tfolder, "\")(0) & "\" & Split(tfolder, "\")(1) & "\" '一级目录
            End If
            Set rngad = .Range("f6:f" & flow).Find(tfolderp, lookat:=xlPart) '判断是否还存在此文件夹在书库的子文件夹的信息
        End With
100
        With .Sheets("目录")
            If rngad Is Nothing Then '清除所有这个文件夹的关联文件夹
                Set rngad = .Cells(4, 3).Resize(blow, csc).Find(tfolderp, lookat:=xlPart) '模糊搜索                        '清除一级文件夹以下所有的文件
                If Not rngad Is Nothing Then
                    If .AutoFilterMode = True Then .AutoFilterMode = False '筛选如果处于开启状态则关闭
                    .Range("a3:a" & blow).AutoFilter Field:=1, Criteria1:=.Cells(rngad.Row, 1).Value
                    .Range("a3").Offset(1).Resize(blow - 3).SpecialCells(xlCellTypeVisible).Delete Shift:=xlShiftUp '删除掉筛选出来的结果
                    .Range("a3").AutoFilter
                End If
            Else  '否则具体的文件夹以及子文件夹
                tfolderp = tfolder & "\"  '重新赋值
                Do
                    Set rngad = .Cells(4, 3).Resize(blow, csc).Find(tfolderp, lookat:=xlPart) '精确搜索 '定点清除特定的文件
                    If Not rngad Is Nothing Then .rows(rngad.Row).Delete Shift:=xlShiftUp
                Loop Until rngad Is Nothing
            End If
        End With
        .Application.ScreenUpdating = True
        .Application.Calculation = xlCalculationAutomatic
    End With
    Set rngad = Nothing
    Set rngadx = Nothing
End Function

Function OpenFileLocation(ByVal address As String)  '打开文件所在位置
    If Len(address) = 0 Then Exit Function
    If fso.folderexists(address) = False Then Exit Function
    Shell "explorer.exe " & address, vbNormalFocus
End Function

Function FileCopy(ByVal addressx As String, ByVal FileName As String, ByVal adrowx As Integer, Optional ByVal cmCode As Byte) As Boolean '文件复制
    Dim mynewpath As String
    Dim rngad As Range
    Dim Filesize As Long
    Dim strfolder As String, strx As String
    
    On Error GoTo 100
    FileCopy = False
    If Len(addressx) = 0 Then Exit Function '空值退出
    If cmCode = 0 Then
        If fso.fileexists(addressx) = False Then '文件不存在,则执行删除整理,将检查文件的存在的任务分解成每一次的操作
            Call DeleMenu(adrowx)
            Exit Function
        End If
    End If
    With Application.FileDialog(msoFileDialogFolderPicker) '文件夹选择窗口
        .Show
        If .SelectedItems.Count = 0 Then Exit Function '未选择文件夹则退出sub
        strfolder = .SelectedItems(1)
    End With
    If CheckFileFrom(strfolder, 2) = True Then MsgShow "文件位置受限", "Warning", 1500: Exit Function '检查添加到的文件夹是否为受限位置
    mynewpath = strfolder & "\"
    strx = Left(strfolder, 1) '盘符
    Filesize = fso.GetFile(addressx).Size
    If Filesize > fso.GetDrive(strx).AvailableSpace Then MsgBox "磁盘空间不足!", vbCritical, "Warning": Exit Function '判断磁盘是否有足够的空间
    
    If fso.fileexists(mynewpath & FileName) = True Then
        MsgBox "文件已存在"
        Exit Function
    Else
        If Filesize <= 52428800 Then '限制50M
            fso.CopyFile (addressx), mynewpath
        Else                         '超出范围,调用cmd命令去执行
            Shell ("cmd /c" & "copy " & addressx & Chr(32) & strfolder), vbHide
        End If
        With ThisWorkbook.Sheets("目录")
            Set rngad = .Cells(4, 3).Resize(.[b65536].End(xlUp).Row, .Cells.SpecialCells(xlCellTypeLastCell).Column).Find(mynewpath, lookat:=xlWhole) '精确搜索'如果复制的文件的位置在目录上就更新时间
            If Not rngad Is Nothing Then rngad.Offset(0, 2) = Now
        End With
        FileCopy = True
    End If
100
    Set rngad = Nothing
End Function

Function DeleOverBack(ByVal str1 As String, ByVal str2 As String, ByVal str3 As String, ByVal str4 As String, ByVal str5 As String, _
ByVal str6 As String, ByVal str7 As String, ByVal str8 As String, ByVal str9 As String, _
ByVal str10 As String, ByVal str11 As String, ByVal str12 As String, ByVal str13 As String, ByVal str14 As String, _
ByVal str15 As String, ByVal str16 As String, ByVal str17 As String, ByVal str18 As String, ByVal str19 As String, _
ByVal str20 As String, ByVal str21 As String, ByVal str22 As String, ByVal str23 As String, ByVal str24 As String, _
ByVal str25 As String, ByVal str26 As String, ByVal str27 As String, ByVal str28 As String, ByVal str29 As String, _
ByVal str30 As String, ByVal str31 As String, ByVal str32 As String, ByVal str33 As String, ByVal str34 As String) '删除之前,备份-修改摘要上的信息

    Dim TableName As String
    Dim strx1 As String
    
    If Len(str1) = 0 Then Exit Function
    If RecData = True Then
        TableName = "摘要记录"
        strx1 = "DL-" & Mid$(str1, 5, Len(str1) - 4) '修改摘要信息
        SQL = "select * from [" & TableName & "$] where 统一编码='" & str1 & "'"                                       '查询数据
        Set rs = New ADODB.Recordset
        rs.Open SQL, Conn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount <> 0 Then
            SQL = "UPDATE [" & TableName & "$] SET 统一编码='" & strx1 & "' WHERE 统一编码='" & str1 & "'"          '更新摘要上的信息
            Conn.Execute (SQL)
        End If
        rs.Close
        Set rs = Nothing
        TableName = "删除备份"  '删除备份
        SQL = "Insert into [" & TableName & "$] (统一编码, 文件名, 文件类型, 文件路径, 文件所在位置, 文件初始大小, 文件修改时间, 文件大小, 文件创建时间, 文件Hash, 文件类别, 最近打开时间, 累计打开次数, 主文件名, 作者, PDF清晰度, 文本质量, 内容评分, 推荐指数, 标签1, 标签2, 标识编号, 添加时间, 名称, 评分, 链接, 异常字符标记, 备注, 来源, 文件名异常字符, 文件位置异常字符, 文字类型, 删除原因, 删除备注) Values ('" & str1 & "','" & str2 & "','" & str3 & "','" & str4 & "','" & str5 & "'," & str6 & ",'" & str7 & "','" & str8 & "','" & str9 & "','" & str10 & "','" & str11 & "','" & str12 & "','" & str13 & "','" & str14 & "','" & str15 & "','" & str16 & "','" & str17 & "','" & str18 & "','" & str19 & "','" & str20 & "','" & str21 & "','" & str22 & "','" & str23 & "','" & str24 & "','" & str25 & "','" & str26 & "','" & str27 & "','" & str28 & "','" & str29 & "','" & str30 & "','" & str31 & "','" & str32 & "','" & str33 & "','" & str34 & "')"
        Conn.Execute (SQL)
    End If
End Function

Function AddPList(ByVal filecode As String, ByVal filen As String, ByVal cmfrom As Byte) As Boolean '添加到优先阅读
    Dim i As Integer, k As Byte
    Dim strx As String, strx1 As String

    With ThisWorkbook.Sheets("主界面")        '主界面记录下信息
'        If .Range("k27").NumberFormatLocal <> "yyyy/m/d h:mm;@" Then .Range("k27:k33").NumberFormatLocal = "yyyy/m/d h:mm;@" '格式修正
        strx1 = LCase(Right$(filen, Len(filen) - InStrRev(filen, ".")))
        If strx1 Like "xl*" Then
            If strx1 <> "xls" And strx1 <> "xlsx" Then AddPList = False: MsgBox "此类型文件不允许添加", vbCritical, "Warning": Exit Function
        End If
        For k = 27 To 33
            strx = .Range("i" & k).Value
            If filecode = strx Then AddPList = False: Exit Function '如果已经存在,则不写入
            If Len(strx) = 0 Then Exit For       '当空值时,退出循环
        Next
        ThisWorkbook.Application.ScreenUpdating = False
        If Len(.Range("i27").Value) = 0 Then             '确保添加进来的值可以一直放在第一行
           .Range("i27") = filecode
           .Range("d27") = filen
           .Range("k27") = Now
        Else
            For i = 33 To 28 Step -1   'i不能是byte类型?                               '最后一行的数据不断被重新写入
                .Range("d" & i) = .Range("d" & i - 1)
                .Range("i" & i) = .Range("i" & i - 1)
                .Range("k" & i) = .Range("k" & i - 1)
            Next
            .Range("i27") = filecode
            .Range("d27") = filen
            .Range("k27") = Now
        End If
    End With
    AddPList = True
    If UF3Show = 3 Then AddPlistx = 1 '窗体隐藏时
    ThisWorkbook.Application.ScreenUpdating = True
    If cmfrom = 0 Then
        ThisWorkbook.Sheets("书库").Label1.Caption = "操作成功"
    Else
        UserForm3.Label57.Caption = "操作成功"
    End If
End Function

Sub FileMove()  '文件移动
    Dim addrowx As Integer, i As Byte, Filesize As Long
    Dim addressx As String, tfolderx As String
    Dim rngad As Range, filext As String
    Dim strfolder As String, mynewpath As String, strx As String
    
    On Error GoTo 100
    addrowx = Selection.Row
    With ThisWorkbook.Sheets("书库")
        addressx = .Range("e" & addrowx).Value
        If addrowx < 6 Or Len(addressx) = 0 Then Exit Sub
        filext = .Range("d" & addrowx).Value
        i = FileTest(addressx, filext, .Cells(addrowx, "c").Value)
        Select Case i                                 ''检查文件的存在/是否处于打开的状态
            Case 1: strx = "未获取到有效值"
            Case 2: strx = "文件不存在": Call DeleMenu(addrowx) '删除表格目录
            Case 3: strx = "文件是txt文件"
            Case 4: strx = "文件处于打开的状态"
            Case 5: strx = "程序异常"
        End Select
        .Label1.Caption = strx
        If i <> 0 And i <> 3 And i <> 6 Then Exit Sub
    
        With Application.FileDialog(msoFileDialogFolderPicker) '文件夹选择窗口
            .Show
            If .SelectedItems.Count = 0 Then Exit Sub      '未选择文件夹则退出sub
            strfolder = .SelectedItems(1)
        End With
        
        tfolderx = .Range("f" & addrowx).Value
        If tfolderx = strfolder Then Exit Sub '相同的文件夹
        If CheckFileFrom(strfolder, 2) = True Then MsgShow "文件位置受限", "Warning", 1500: Exit Sub '检查添加到的文件夹是否为受限位置
        strx = Left(strfolder, 1)
        Filesize = fso.GetFile(addressx).Size '获取文件的实际大小
        If Filesize > fso.GetDrive(strx).AvailableSpace Then MsgBox "磁盘控件不足!", vbCritical, "Warning": Exit Sub '判断磁盘是否有足够的空间
        
        mynewpath = strfolder & "\" '注意,迁移文件的时候,目标文件夹要有"\"
        If fso.fileexists(mynewpath & filext) = True Then   '判断文件是否已存在
            MsgBox "文件已存在"
            Exit Sub
        '---------------------------------------------------------------------------------------------------------------------------------前期准备
        Else
            If Filesize <= 52428800 Then '限制50M
                fso.MoveFile (addressx), mynewpath
            Else                         '超出范围,调用cmd命令去执行
                Shell ("cmd /c" & "move " & addressx & Chr(32) & mynewpath), vbHide
            End If
        End If
        DisEvents
        If ErrCode(mynewpath, 1) > 1 Then '检查新的文件夹的路径是否存在异常字符'并获取文件路径的异常字符的位置
            .Range("af" & addrowx) = errcodepx
            .Range("ab" & addrowx) = "ERC"
            If .Range("ae" & addrowx) = "ENC" Then .Range("ac" & addrowx) = "EDC"
        Else
            .Range("af" & addrowx) = ""
            If .Range("ae" & addrowx) = "EDC" Then
                .Range("ac" & addrowx) = "ENC"
            Else
                .Range("ab" & addrowx) = ""
            End If
        End If
        
        DeleFileOver (tfolderx) '当文件被移动后,其作用如同删除
        
        With ThisWorkbook.Sheets("目录") '文件迁移后,原来的/现在的文件夹修改时间发生变化
            Set rngad = .Cells(4, 3).Resize(.[b65536].End(xlUp).Row, .Cells.SpecialCells(xlCellTypeLastCell).Column).Find(mynewpath, lookat:=xlWhole)
            If Not rngad Is Nothing Then rngad.Offset(0, 2) = Now: rngad.Offset(0, 4) = rngad.Offset(0, 4) + 1
        End With
        .Range("e" & addrowx) = mynewpath & .Range("c" & addrowx) '更新目录上的信息
        .Range("f" & addrowx) = strfolder
        EnEvents
        If i = 3 Then
            MsgBox "修改成功!" & Chr(13) & "txt文档可以在打开的状态下移动\重命名\删除"  'txt文档可以在打开的状态下移动和重命名等操作
        Else
            MsgBox "修改成功!"
        End If
    End With
    Set rngad = Nothing
Exit Sub
100
    If Err.Number = 70 Then '文件处于打开的状态
        MsgBox "文件处于打开的状态"
    Else
        MsgBox "出现异常错误"
    End If
    EnEvents
    Err.Clear
End Sub
