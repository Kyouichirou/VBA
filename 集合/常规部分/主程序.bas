Attribute VB_Name = "主程序"
Option Explicit
'文件的所在文件夹最好不要包含非ansi字符, 但是文件名包含非ansi字符处理还不算麻烦
Option Compare Text                       '不区分大小写
Dim arrfiles(1 To 10000) As String                 '定义一个数组后面用以存放path数据(数值可变)
Dim arrbase(1 To 10000) As String                   '存储文件名
Dim arrextension(1 To 10000) As String             '存储文件扩展名
Dim arrsize(1 To 10000) As String                   '存储文件的大小
Dim arrparent(1 To 10000) As String               '存储文件所在位置
Dim arrdate(1 To 10000) As String                  '存储文件创建日期
Dim arrsizeb(1 To 10000) As Long                  '文件的大小,单位比特
Dim arrmd5(1 To 10000) As String                    '文件md5
Dim arredit(1 To 10000) As String                     '文件修改时间
Dim arrcode(1 To 10000) As String                    '文件路径异常字符
Dim arrcm(1 To 10000) As String                       '备注
Dim arrfnansi(1 To 10000) As String                 '文件名非ansi标注
Dim arrfpansi(1 To 10000) As String                  '文件位置非ansi标注
Dim arrfilen() As Variant, arrfilesize() As Variant, arrfilemody() As Variant, arrfilemd5() As Variant, arrfilep() As Variant '表格数组,用于比较

Public fso As New FileSystemObject '------------------------------------------------------核心
Public ShellxExist As Byte '判断Powershell是否存在
Public AddFx As Byte '标记文件添加动作被执行

Dim flc As Integer   '添加文件的总数                              '% integer型数据标识符
Dim dl As Integer '删除文件的总数
Dim ls As Integer   '近似文件统计                          'integer类型数据的范围-32768 到 32767
Dim fc As Long '文件总数
Dim b As Byte, a As Integer '目录的行列号

Dim md5x As Integer
Dim umd5x As Integer
Dim idele As Integer '统计删除了多少行数据
Dim ix As Integer
Dim deledic As New Dictionary '存储删除掉文件所在的行号,集中删除
Dim rnglists As Range
Dim F As Byte '用于判断文件夹是否处于添加的状态
Dim xi As Byte, c As Byte 'x用于标记数据的存在, c标记文件夹的层级
Dim Elow As Integer '表格数据的最后行号

Function ListAllFiles(ByVal addcode As Byte, ByVal FilePath As String) As Boolean               'addcode表示数据更新的方式,0为默认添加,1为包含子文件夹,2为不包含子文件夹
    Dim fd As Folder, fdatt As Long
    Dim i As Integer, j As Integer
    Dim tl As Single, t As Single '记录时间
    Dim rnglist As Range '目录
    Dim strp As String, strptemp As String, strfolder As String
    Dim bc As Integer, cN As Integer, tracenum As Byte
    Dim blow As Integer, cright As Byte '目录行列数据的分布的最后一行
    Dim filec As Integer, ifilec As Integer, clm As Integer, bcx As Integer
    Dim xtemp As Variant, alow As Integer
    
    If addcode = 0 And FilePath = "NU" Then
        With Application.FileDialog(msoFileDialogFolderPicker) '文件夹选择窗口(文件夹以此只能选一个)
            .Show
            If .SelectedItems.Count = 0 Then ListAllFiles = False: Exit Function '未选择文件夹则退出sub
            strfolder = .SelectedItems(1)
        End With
        FilePath = strfolder '& "\" '需要处理的文件的路径
    End If
    ListAllFiles = False
'    If CheckFileX(strfolder) = False And addcode = 0 Then MsgShow "文件夹不存在目标文件", "Warning", 1500: Exit function '调用cmd对文件夹进行快速检查,判断文件夹内是否存在目标文件
    If CheckFileFrom(strfolder, 2) = True Then '检查文件夹的来源,假如是系统盘的位置,只允许document,download,desktop三个位置的文件夹添加
        MsgBox "系统盘的文件只允许添加来自:" & vbCr & "Desktop" & vbCr & "Downloads" & vbCr & "Documents"
        Exit Function
    End If
    If CheckPathAsWorkbook(strfolder) = True Then MsgBox "文件夹添加位置受限", vbInformation, "Tips": Exit Function '限制来自本工作簿下所属的文件
    Set fd = fso.GetFolder(FilePath)          '将fd指定路径对象
        If fd.IsRootFolder Then                    '防止整个个盘符添加
        MsgBox "整个盘符添加,处理时间过长", vbOKOnly, "Careful!!!"
        Set fd = Nothing
        Exit Function
    End If
    fdatt = fd.Attributes
    If fd.ParentFolder.Path = Environ("SYSTEMDRIVE") & "\" Or fdatt = 18 Or fdatt = 1046 Then
    '------------------------禁止添加隐藏文件夹,系统盘的一级文件夹'禁止添加系统盘的一级文件夹,防止出现部分文件夹无法访问的问题,非目标文件夹
        Set fd = Nothing
        MsgBox "禁止添加疑是系统文件夹/隐藏文件夹", vbCritical, "Warning!!!"
        Exit Function
    End If
    
    ifilec = fd.Files.Count
    If ifilec = 0 And fd.SubFolders.Count = 0 Then
        Set fd = Nothing
        MsgBox "添加的文件夹为空", vbOKOnly, "Careful!"
        Exit Function
    End If
'-------------------------------------------------------------------------------------------------检查文件夹的基本情况
    With ThisWorkbook
    
        DisEvents '禁止干扰
        ' 需要注意在执行复杂的进程的时候,需要考虑各种容易被触发的事件,或者时有冲突的进程或者重复被执行的进程,以减少不必要的事件浪费
'        UserForm6.Show 0 '注意这里交互为0的时候后续的代码可以继续执行,为1时,代码会被中断
        
        t = Timer '计算代码运行时间的初始值
        
        flc = 0 '工程级/模块级变量的初始化 '涉及到两种变量的生存周期,除非进行调试或者重新设置,否则两种变量可以长期存在
        dl = 0
        ls = 0
        fc = 0
        b = 2 '目录'位置
        a = 0
        F = 0
        xi = 0
        c = 0
        ix = 0
        AddFx = 0
        idele = 0
        md5x = 0
        umd5x = 0
        Elow = 0 '参数的初始化
        c = UBound(Split(FilePath, "\")) '文件夹位于磁盘的目录等级
        
        If PSexist = True Then ShellxExist = 1 '需要判断powershell版本的信息,不能低于4.0
        
        DataTreat '处理已有的表格内容
        
        With .Sheets("目录")
            blow = .[b65536].End(xlUp).Row
            cright = .Cells.SpecialCells(xlCellTypeLastCell).Column '表格的最右侧
            If Len(.Range("b4").Value) = 0 Then '尚未写入数据
                a = blow + 1
                xi = 1    '用于标记表格是否已经存有数据
                GoTo 107
            Else
                If c = 1 Then                                          '表格尚未写入数据/一级目录不存在关联目录'即可判定新添加进来的文件夹的信息不需要判定直接写入数据
                    strp = FilePath & "\" '增加"\"符号是为了防止c:\a,c:\ab这种情况的出现,在模糊搜索中无法准确找到目标
                Else
                    xtemp = Split(FilePath, "\")
                    strp = xtemp(0) & "\" & xtemp(1) & "\" '一级目录
                End If
                Set rnglist = .Cells(4, 3).Resize(blow, cright).Find(strp, lookat:=xlPart) '通过检查一级文件夹,判断是否存在关联文件夹
                If rnglist Is Nothing Then
                    xi = 1 '这个文件夹尚未有子文件夹被添加进来 ,Xi=1 表示不需要执行比较
                    a = blow + 1
                    GoTo 107
                End If
                '------------------------------------------------------------------------------------------------检查是否已经存在相关的文件夹添加进来
                strp = FilePath & "\"
                Set rnglist = .Cells(4, 3).Resize(blow, cright).Find(strp, lookat:=xlWhole) '完全匹配查找
                
                If Not rnglist Is Nothing Then '有值 '这部分处理主文件夹位于目录,添加进来的为子文件夹 ''存在数据的最大列
                    F = 1
                    filec = rnglist.Offset(0, 4).Value '目前文件夹的文件数量
                    If addcode = 1 Then     '更新的方式不同 ,1,2分别表示是否包含子文件夹
                        a = rnglist.Row
                        If fd.DateLastModified = rnglist.Offset(0, 2) Then
                            F = 3 '标记这一层的文件不处理
                            GoTo 109
                        Else
                            If Int(Abs(filec - ifilec) / filec) > 50 Then '文件夹发生变化范围的幅度足够大 'abs绝对值函数
                                F = 4
                                GoTo 107
                            End If
                        End If
                    ElseIf addcode = 2 Then
                        If fd.DateLastModified = rnglist.Offset(0, 2) Then
                            GoTo 1001 '文件夹的内容没有发生变化(只更新文件夹一层)'不包含子文件夹
                        Else
                            a = rnglist.Row
                            GoTo 107
                        End If
                    End If
                    MsgBox "此文件已添加" '-已经添加的文件夹禁止再以添加的方式添加
                    GoTo 1001
                Else
                    clm = cright - 5 '数据的最右侧
                    bc = b + c
                    Do
                        strp = fd.Path
                        Set fd = fd.ParentFolder
                        strptemp = fd.Path & "\"
                        Set rnglist = .Cells(4, 3).Resize(blow, cright).Find(strptemp, lookat:=xlWhole) '两者之间为父子文件夹
                        If Not rnglist Is Nothing Then
                            F = 2
                            a = rnglist.Row + 1
                            .Cells(a, 1).EntireRow.Insert
                            GoTo 104
                        End If
                        '----------------------------------------------------------------------------------用于确定文件夹在表格的位置,确保任意添加进来的父子文件夹之间能够在相邻的位置
                        strp = strp & "\"
                        For bcx = bc To clm '采用循环的命令强制搜索的列顺序依次进行搜索(默认的搜索顺序是不确定的)
                            Set rnglist = .Cells(4, bcx).Resize(blow, 1).Find(strp, after:=.Cells(4, bcx), lookat:=xlPart)
                            If Not rnglist Is Nothing Then
                                F = 2
                                If c = 1 Then                        '保证1级文件夹位于区域的第一行
                                    For cN = rnglist.Row To 3 Step -1
                                        If .Cells(cN, 1) <> .Cells(rnglist.Row, 1) Then
                                            a = cN + 1
                                            .Cells(a, 1).EntireRow.Insert
                                            GoTo 104
                                        End If
                                    Next
                                Else
                                    If CInt(.Cells(rnglist.Row, 2)) = 1 Then '除了一级文件
                                        a = rnglist.Row + 1
                                    Else
                                        a = rnglist.Row
                                    End If
                                    .Cells(a, 1).EntireRow.Insert '非关联的父子文件夹一律放在上一行
                                End If
                                GoTo 104
                            End If
                        Next
                        bc = bc - 1          '文件层级的变化
                    Loop Until fd.IsRootFolder
                    a = blow + 1 '如果没找到
            End If
104
            Set fd = fso.GetFolder(FilePath)          '重新将fd指定路径对象
107
        End If
            .Cells(a, 2) = c
            .Cells(a, 2).NumberFormatLocal = "0_);[红色](0)"
            .Cells(a, b + c) = fd.Path & "\"
            .Cells(a, b + c + 1) = fd.DateCreated
            .Cells(a, b + c + 1).NumberFormatLocal = "yyyy/m/d h:mm;@" '格式调整
            If F > 0 Then
                .Cells(a, 1) = .Cells(rnglist.Row, 1)
            ElseIf F = 0 Then
                alow = .[a65536].End(xlUp).Row
                .Range("a" & alow).AutoFill Destination:=.Range("a" & alow & ":" & "a" & a), type:=xlFillDefault
            End If
            .Cells(a, b + c + 2) = fd.DateLastModified
            .Cells(a, b + c + 2).NumberFormatLocal = "yyyy/m/d h:mm;@"
            .Cells(a, b + c + 3) = fd.SubFolders.Count
            .Cells(a, b + c + 3).NumberFormatLocal = "0_);[红色](0)" '设置单元格的格式,否则会被Excel转换为其他的格式
            .Cells(a, b + c + 4) = fd.Files.Count
            .Cells(a, b + c + 4).NumberFormatLocal = "0_);[红色](0)"
            .Cells(a, b + c + 5) = fd.Size
            .Cells(a, b + c + 5).NumberFormatLocal = "0_);[红色](0)"
        End With
'------------------------------------------------------------------------------------------------------------------------------------------文件夹在目录的位置的确定,更新数据的方式
109
        c = c + 1
        SearchFolders fd, addcode                                   '调用sf子sub检索子文件夹和获取文件夹内文件的信息
        
        If flc = 0 Then
            If addcode > 0 Then
                GoTo 1001
            Else
                GoTo 1002 '当没有返回值的时候
            End If
        End If
            
        Call WriteData(2) '写入数据
    
1002
        With .Sheets("主界面")                     '往主界面写入信息部分(需要修改)
            If Len(.Range("e37").Value) = 0 Then            '尚未写入数据
                .Range("e37") = FilePath
                .Range("i37") = Now
            Else
                j = .[e65536].End(xlUp).Row
                Set rnglist = .Range("e37:e" & j).Find(FilePath, lookat:=xlWhole)
                If Not rnglist Is Nothing Then
                    GoTo 1001
                Else
                    .Range("e" & j + 1) = FilePath
                    .Range("i" & j + 1) = Now
                End If '
            End If
        End With
    End With
    
1001
    
    Set fd = Nothing
    Set rnglist = Nothing
    Set rnglists = Nothing
    ListAllFiles = True
    EnEvents
    
'    Unload UserForm6
'    tl = Timer - t
'    If tl > 30 Then
'        MsgBox "处理完成,花费时间: " & Format(tl, "0.0000") & "s" & vbCr _
'        & "总共发现: " & fc & "个文件" & vbCr _
'        & "成功导入: " & flc & "个文件" & vbCr _
'        & "发现可疑重叠文件: " & ls & "个" & vbCr _
'        & "删除重叠文件" & dl & "个"
'    Else
'        MsgBox "处理完成! " & vbCr _
'        & "总共发现: " & fc & "个文件" & vbCr _
'        & "成功导入: " & flc & "个文件" & vbCr _
'        & "发现可疑重叠文件: " & ls & "个" & vbCr _
'        & "删除重叠文件" & dl & "个"
'    End If
End Function

Private Sub DataTreat(Optional addcode As Byte) '生成表格的数组
    Dim itemp As Integer
    With ThisWorkbook.Sheets("书库")               '将这部分将用于和添加进来的文件进行比较,相比于find,在速度上有更好的优势,需要考虑数据量非常大的时候对于内存的占用,当数据足够大的时候,和find相比,速度的优势开始不断缩小
        Elow = .[c65536].End(xlUp).Row + 1
        If Elow < 5 Then MsgBox "书库结构被破坏", vbCritical, "Warning!!!":  Exit Sub
        If Elow = 6 Then Exit Sub
        itemp = Elow - 6
        If itemp = 1 Then
            If fso.fileexists(.Cells(6, "e").Value) = False Then   '如果文件特别少的时候可以逐一检查原有的文件是否还存在于目录对应的位置
                .rows(6).Delete Shift:=xlShiftUp '如果内容被全部清空
                Elow = 6 '新的位置
                If addcode = 0 Then
                    ClearAll (0)
                Else
                    ClearAll (1) '更新的方式-清除目录内容
                End If
                Exit Sub
            End If
            ReDim arrfilen(1 To 1, 1 To 1)
            ReDim arrfilesize(1 To 1, 1 To 1) 'redim不能修改数组的数据类型
            ReDim arrfilemody(1 To 1, 1 To 1)
            ReDim arrfilemd5(1 To 1, 1 To 1)
            ReDim arrfilep(1 To 1, 1 To 1)
            arrfilen(1, 1) = .Cells(6, "c").Value
            arrfilesize(1, 1) = .Cells(6, "g").Value
            arrfilemody(1, 1) = .Cells(6, "h").Value
            arrfilemd5(1, 1) = .Cells(6, "k").Value
            arrfilep(1, 1) = .Cells(6, "e").Value
        ElseIf itemp > 1 Then
            itemp = Elow - 1
            arrfilen = .Range("c6:c" & itemp).Value
            arrfilesize = .Range("g6:g" & itemp).Value '不要使用for循环取值,太慢.application.transpose又上限,在处理大概34000+的数据时,出现问题
            arrfilemody = .Range("h6:h" & itemp).Value
            arrfilemd5 = .Range("k6:k" & itemp).Value
            arrfilep = .Range("e6:e" & itemp).Value
        End If
    End With
End Sub

Private Sub WriteData(ByVal xi As Byte) '将获取到的内容写入到表格,清理数组的内容
    Dim itemp2 As Integer, i As Integer, k As Integer
    Dim arrdele() As Integer, arrdelex() As Integer
    
    Elow = Elow - idele
    With ThisWorkbook.Sheets("书库")                       '默认的一维数组是一行,transpose函数可以将数据转成列形式排列,
    '----------------------------------------------注意Application.Transpose转置的数据的上限,在转置表格成数组的时候上限为34446左右(不同的设备或者office可能数据有所不同)
        .Range("k" & Elow).Resize(flc) = ThisWorkbook.Application.Transpose(arrmd5)  '文件md5
        .Range("ab" & Elow).Resize(flc) = ThisWorkbook.Application.Transpose(arrcode) '异常字符标记
        .Range("ac" & Elow).Resize(flc) = ThisWorkbook.Application.Transpose(arrcm)   '备注
        .Range("ae" & Elow).Resize(flc) = ThisWorkbook.Application.Transpose(arrfnansi)
        .Range("af" & Elow).Resize(flc) = ThisWorkbook.Application.Transpose(arrfpansi)
        .Range("c" & Elow).Resize(flc) = ThisWorkbook.Application.Transpose(arrbase) '文件名
        .Range("d" & Elow).Resize(flc) = ThisWorkbook.Application.Transpose(arrextension) '文件扩展名
        .Range("e" & Elow).Resize(flc) = ThisWorkbook.Application.Transpose(arrfiles) '文件路径
        .Range("f" & Elow).Resize(flc) = ThisWorkbook.Application.Transpose(arrparent) '文件所在位置
        .Range("g" & Elow).Resize(flc) = ThisWorkbook.Application.Transpose(arrsizeb) '文件初始大小
        .Range("h" & Elow).Resize(flc) = ThisWorkbook.Application.Transpose(arredit) '文件修改时间
        .Range("i" & Elow).Resize(flc) = ThisWorkbook.Application.Transpose(arrsize) '文件大小
        .Range("j" & Elow).Resize(flc) = ThisWorkbook.Application.Transpose(arrdate) '文件创建时间
        itemp2 = flc + Elow - 1
        .Range("x" & Elow & ":" & "x" & itemp2) = Now '添加目录的时间
        .Range("ad" & Elow & ":" & "ad" & itemp2) = xi '标注文件的来源(是通过添加文件夹的方式添加进来的)
        .Range("b" & Elow) = .Cells(5, 1).Value
        itemp2 = itemp2 + 1
        .Range("b" & Elow).AutoFill Destination:=.Range("b" & Elow & ":" & "b" & itemp2), type:=xlFillDefault '添加统一编码
        .Cells(5, 1).Value = .Range("b" & itemp2) '确保所有的值都具有唯一性
        .Range("b" & itemp2).ClearContents
        If UF3Show = 1 Or UF3Show = 3 Then AddFx = 1 '这个参数用于窗体数据更新使用
        Erase arrfiles '清空数组,静态数组的清除和动态数组的清除有轻微的差别,动态数组将被完全抹掉,静态数组依然可以保留范围,值被抹掉
        Erase arrbase
        Erase arrextension
        Erase arrsize
        Erase arrparent
        Erase arrdate
        Erase arrsizeb
        Erase arrmd5
        Erase arredit
        Erase arrcode
        Erase arrcm
        Erase arrfnansi
        Erase arrfpansi
        Erase arrfilen
        Erase arrfilesize
        Erase arrfilemody
        Erase arrfilep
        Erase arrfilemd5
        If idele > 0 Then
            i = deledic.Count - 1
            ReDim arrdele(i)
            ReDim arrdelex(i)
            For k = 0 To i
                arrdele(k) = deledic.Keys(k)
            Next
            arrdelex = Down(arrdele) '降序排序
            For k = 0 To i
                .rows(arrdelex(k)).Delete Shift:=xlShiftUp
                DeleFileOver .Cells(arrdelex(k), 6).Value
            Next
    '        deledic.RemoveAll
            Erase arrdele
            Erase arrdelex
            Set deledic = Nothing
        End If
    End With
End Sub

Private Sub SearchFolders(ByVal fd As Folder, ByVal addcodex As Integer)              'ByVal,为按值传递方式, 有别于byref(按引用方式传递)
    Dim flx As File, fdatt As Long
    Dim sfd As Folder
    Dim strTemp As String
    Dim bc As Integer, bcx As Integer, clm As Integer
    Dim blow As Integer, cright As Integer '目录行列数据的分布的最后一行
    Dim filec As Integer, ifilec As Integer, strp As String
    
    ix = ix + 1 '用于标记文件夹
    If F = 3 Then GoTo 1007 '当文件夹更新的时候,第一层文件及没有发生变化,直接调到查看子文件夹/或者文件夹下没有文件
    For Each flx In fd.Files                   '搜索文件
        fc = fc + 1 '统计文件数量
        FileIn flx
    Next flx
1007
    If addcodex = 2 Then Exit Sub '只操作文件夹一层不涉及子文件夹
    
    If fd.SubFolders.Count = 0 Then Exit Sub  '子文件夹数目为零则退出sub
    
    For Each sfd In fd.SubFolders             '搜索子文件夹
        ifilec = sfd.Files.Count
        If ifilec = 0 And sfd.SubFolders.Count = 0 Then GoTo 107 '没有文件,子文件夹为空
        fdatt = sfd.Attributes
        If fdatt = 18 Or fdatt = 1046 Then GoTo 107 '过滤掉隐藏文件夹
        F = 0                          'f为模块变量,使用后重置
        With ThisWorkbook.Sheets("目录")
            blow = .[b65536].End(xlUp).Row
            cright = .Cells.SpecialCells(xlCellTypeLastCell).Column '当数据需要重复引用时,减少重复取值
            If xi = 1 Then
                a = blow + 1 '不需要执行比较
                GoTo 109
            End If
            strTemp = sfd.Path & "\"
            bc = b + c
            clm = cright - 5
            Set rnglists = .Cells(4, 3).Resize(blow, cright).Find(strTemp, after:=.Cells(4, 3), searchorder:=xlByColumns, lookat:=xlWhole)
            If Not rnglists Is Nothing Then '如果目录已经存在
                F = 1
                a = rnglists.Row
                If sfd.DateLastModified <> rnglists.Offset(0, 2) Then '文件的修改时间已经发生变化
                    filec = Rng.Offset(0, 4).Value
                    If filec = 0 Then F = 0: GoTo 109 '原来的文件夹没有内容
                    If 100 * Int(Abs(ifilec - filec) / filec) > 50 Then F = 4 '文件夹的数量变化范围 int转化为integer型,abs,绝对值,当文件夹的内容发生大的变化时
                    GoTo 109
                Else
                    F = 3       '时间相同,跳过该文件
                    If sfd.SubFolders.Count = 0 Then
                        GoTo 107
                    Else
                        GoTo 106
                    End If
                End If
            Else
                Do
                    strp = sfd.Path & "\"
                    Set sfd = sfd.ParentFolder
                    Set rnglists = Nothing
                    For bcx = bc To clm
                        Set rnglists = .Cells(4, bcx).Resize(blow, 1).Find(strp, after:=.Cells(4, bcx), lookat:=xlPart)
                        '-----------------在模糊搜索的时候,搜索的位置并不一定严格从最左边指定的区域开始,只能通过循环的方式,强制搜索每一列
                        If Not rnglists Is Nothing Then
                            If CInt(.Cells(rnglists.Row, 2)) >= c Then
                                a = rnglists.Row
                            Else
                                a = rnglists.Row + 1
                            End If
                            .Cells(a, 1).EntireRow.Insert
                            GoTo 110
                        End If
                    Next
                    bc = bc - 1 '文件夹层级上升
                Loop Until sfd.IsRootFolder
                a = blow + 1
            End If
    '-------------------------------------------------------------------------------------------------------------------------------用于确定子文件夹在目录部分的位置
110
            Set sfd = fso.GetFolder(strTemp) '重新赋值
109
            .Cells(a, 1) = .Cells(a - 1, 1)
            .Cells(a, 2) = c
            .Cells(a, 2).NumberFormatLocal = "0_);[红色](0)"
            .Cells(a, b + c) = sfd.Path & "\"
            .Cells(a, b + c + 1) = sfd.DateCreated
            .Cells(a, b + c + 1).NumberFormatLocal = "yyyy/m/d h:mm;@"
            .Cells(a, b + c + 2) = sfd.DateLastModified
            .Cells(a, b + c + 2).NumberFormatLocal = "yyyy/m/d h:mm;@"
            .Cells(a, b + c + 3) = sfd.SubFolders.Count
            .Cells(a, b + c + 3).NumberFormatLocal = "0_);[红色](0)"
            .Cells(a, b + c + 4) = ifilec
            .Cells(a, b + c + 4).NumberFormatLocal = "0_);[红色](0)"
            .Cells(a, b + c + 5) = sfd.Size
            .Cells(a, b + c + 5).NumberFormatLocal = "0_);[红色](0)"
        End With
106
        a = a + 1
        If sfd.SubFolders.Count > 0 Then
        c = c + 1 '标记文件的层级
        End If
108
        SearchFolders sfd, addcodex
107
    Next
    c = c - 1 '每一次都会穷尽二层文件夹下的一个文件夹下的最后一层子文件夹,当执行完毕之后,重新计算开始的那一层,所以c要进行递减
End Sub

Sub AddFile() '--------------------------------------------添加文件
    Dim fdx As FileDialog, fl As File, fd As Folder
    Dim selectfile As Variant
    Dim rngx As Range, strfd As String
    Dim i As Byte, k As Byte
    
    Set fdx = ThisWorkbook.Application.FileDialog(msoFileDialogFilePicker)
    On Error GoTo 100
    With fdx
        .AllowMultiSelect = True '允许选择多个文件(注意不是文件夹,文件夹只能选一个)
        .Show
        i = .SelectedItems.Count
        If i = 0 Then Exit Sub
        If i > 10 Then MsgBox "限制一次添加10个文件", vbOKOnly, "Careful!": Exit Sub
        ix = 0    '模块级别参数重置
        AddFx = 0
        idele = 0
        dl = 0
        ls = 0
        md5x = 0
        umd5x = 0
        Elow = 0
        flc = 0
    
        DisEvents
        DataTreat
        '-------------------准备工作
        For Each selectfile In .SelectedItems
            If k = 0 Then
                If CheckFileFrom(selectfile, 1) = True Or CheckPathAsWorkbook(selectfile, 1) = True Then Exit Sub '限制文件的来源
            End If
            k = k + 1
            Set fl = fso.GetFile(selectfile)
            FileIn fl
        Next
    End With
    '--------------------------------------------------------------------------------------文件的添加
    If flc = 0 Then GoTo 100
    Call WriteData(1)
    strfd = fl.ParentFolder
    Set fd = fso.GetFolder(strfd)
    strfd = strfd & "\"
    With ThisWorkbook.Sheets("目录") '添加的文件,说明这是新的文件,那么需要更新所在文件夹的信息
        Set rngx = .Cells(4, 3).Resize(.[b65536].End(xlUp).Row, .Cells.SpecialCells(xlCellTypeLastCell).Column).Find(strfd, after:=.Cells(4, 3), searchorder:=xlByColumns, lookat:=xlWhole)
    End With
    With fd
        If Not rngx Is Nothing Then rngx.Offset(0, 2) = .DateLastModified: rngx.Offset(0, 4) = .Files: rngx.Offset(0, 5) = .Size
    End With
    '------------------------------------------------------------------------------------------------------------信息写入目录/更新
100
    Set fdx = Nothing
    Set fl = Nothing
    Set fd = Nothing
    Set rngx = Nothing
    EnEvents
End Sub

Private Function FileIn(ByVal fl As File)
    '--------------------------------------文件信息处理 'f为0的时候,表示这个全新的文件夹,即执行的方式为全部更新,f=1的时候表示已有文件夹,
    '--------------------------------------部分更新,f=2的时候,存在已有的关联文件夹,全部更新,f=3的时候,已有文件夹,修改时间和目录一直,直接跳过
    Dim filefd As String '文件所在文件夹
    Dim filen As String '文件名
    Dim filemd5 As String '文件md5
    Dim filex As String '文件扩展名
    Dim filez As Long '文件大小 可最大2G的文件
    Dim filep As String '文件路径
    Dim filem As String '文件修改时间
    Dim filect As String '文件创建时间
    Dim p As Integer           '是否存在特殊字符
    Dim k As Byte '判断文件类型
    Dim md5t As Byte
    Dim j As Integer
    
    '--------------------------------------------------------- '对文件的基本属性进行获取和处理-----'10-20-30-40,如果代码运行出错,用于锚定错误的出现的区域
    On Error GoTo 100
10
    With fl
        filep = .Path '路径
        filex = fso.GetExtensionName(filep)
        filex = LCase(filex) '扩展名'限制添加进目录的文件类型 ' l/ucase为转换大小函数 ,这里统一使用小写
        If filex Like "epub" Or filex Like "pdf" Or filex Like "mobi" Then
            k = 1
        ElseIf filex Like "do*" Or filex Like "xl*" Or filex Like "pp*" Or filex Like "tx*" Or filex Like "ac*" Then
            k = 2
        End If
        If k = 0 Then Exit Function
        filez = .Size '大小
        If Len(filex) = 0 Or filez = 0 Or fl.Attributes = 34 Then Exit Function '文件扩展名为空/隐藏文件的跳过 ,34表示hidden属性,注意不要直接使用hidden来表示属性,无法识别
        filen = .Name '文件名
        filefd = .ParentFolder '文件夹
        filect = .DateCreated '创建时间
        filem = .DateLastModified '修改时间
    End With
    
    '-------------------------------------------------------------------对文件的信息进行比较和处理
20
    p = ErrCode(filen, 0, filefd)  '进行非ansi字符检查
    If p = -1 Then GoTo 100 '未获取到有效路径(读取数据异常)
    
    If Elow = 6 Then '表格尚未写入数据/不需要进行和表格的数据比较
        If k = 1 Then
            filemd5 = GetFileHashMD5(filep, p) '计算md5
            If Len(filemd5) = 2 Then
                If md5x > 1 Then md5t = 1 '进行文件名比较
            End If
            md5x = md5x + 1
        Else
            umd5x = umd5x + 1
        End If
    Else '已存在数据
        If F = 1 Or F = 4 Then '已更新的方式来比较文件,先进行文件名比较,之后再进行md5比较
            If FileComp(filen, filep, filemd5, filez, filem, 3) = 5 Then Exit Function '文件名相同,跳过'如果文件夹已经存在于目录,那么先不计算md5,先进行文件名比较
        End If
        If k = 1 Then
            filemd5 = GetFileHashMD5(filep, p) '计算md5
            If Len(filemd5) = 2 Then            '如果返回无效md5
                If FileComp(filen, filep, filemd5, filez, filem, 2) = 3 Then ls = ls + 1: Exit Function '与表格的数据进行比较
            Else
                If FileComp(filen, filep, filemd5, filez, filem, 1) = 1 Then Call DeleRFile(filep, filex, p): Exit Function '和表格的数据一致,删除
            End If
            md5x = md5x + 1
        Else
            If FileComp(filen, filep, filemd5, filez, filem, 2) = 3 Then ls = ls + 1: Exit Function '与表格的数据进行比较
            umd5x = umd5x + 1
        End If
    End If
    
    '----------------------------- '在添加中的文件比较 -先和表格的数据比较,在进行添加中的文件比较
30
    If md5x > 1 Then
        For j = 1 To flc '可以用词典dic.exist来取代
            If filemd5 = arrmd5(j) Then Call DeleRFile(filep, filex, p): Exit Function
        Next
    End If
    If umd5x > 1 Or md5t = 1 Then 'md5t表示无法返回有效md5
        For j = 1 To flc
            If filen = arrbase(j) Then
                If filez = arrsizeb(j) Then
                    If filem = arredit(j) Then GoTo 100 '跳过
                End If
            End If
        Next
    End If
    '-----------------------------------------------------------------------------------------数据写入
40
    flc = flc + 1
    arrmd5(flc) = filemd5
    arrfiles(flc) = filep                  '数组存储文件的路径
    arrbase(flc) = filen                   '包含扩展名的文件名
    arrextension(flc) = filex                    '文件扩展名
    arrparent(flc) = filefd '文件上一级目录
    arrdate(flc) = filect '文件创建
    arrsizeb(flc) = filez         '文件大小
    arredit(flc) = filem '文件修改时间
    arrcode(flc) = IIf(p > 1, "ERC", "")              'https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/iif-function
    If Tagfnansi = True And Tagfpansi = True Then 'd
        arrfnansi(flc) = errcodenx
        arrfpansi(flc) = errcodepx
        arrcm(flc) = "EDC"
    ElseIf Tagfnansi = True And Tagfpansi = False Then 's
        arrfnansi(flc) = errcodenx
        arrfpansi(flc) = ""
        arrcm(flc) = "ENC" '出现的位置
    ElseIf Tagfnansi = False And Tagfpansi = True Then 's
        arrfnansi(flc) = ""
        arrfpansi(flc) = errcodepx
        arrcm(flc) = "EPC" '出现的位置
    ElseIf Tagfnansi = False And Tagfpansi = False Then
        arrfnansi(flc) = ""
        arrfpansi(flc) = ""
        arrcm(flc) = ""
    End If
    If filez < 1048576 Then
        arrsize(flc) = Format(filez / 1024, "0.00") & "KB"    '文件字节大于1048576显示"MB",否则显示"KB"
    Else
        arrsize(flc) = Format(filez / 1048576, "0.00") & "MB"
    End If
    Exit Function
100
    If Erl = 40 And flc > 1 Then flc = flc - 1
    Err.Clear
End Function

Private Function FileComp(ByVal filenx As String, ByVal filep As String, ByVal filemd5x As String, ByVal filezx As String, _
ByVal filemx As String, ByVal cmCode As Byte) As Byte '判断文件是否和表格已存在的目录重叠
    Dim m As Integer, n As Integer
    Dim itemp As Integer
    
    itemp = Elow - 6
    With ThisWorkbook.Sheets("书库")
        If cmCode = 1 Then
            For m = 1 To itemp
                If filemd5x = arrfilemd5(m, 1) Then 'md5对比'可以用词典dic.exist来取代 这里需要注意只有md5这组数据可以如此处理,因为具有唯一性
                    If fso.fileexists(arrfilep(m, 1)) = False Then '同时检查文件是否存在
                        n = m + 5
                        deledic(n) = ""
                        idele = idele + 1
                        FileComp = 0
                        Exit Function
                    Else
                        FileComp = 1
                    End If
                    Exit Function
                End If
            Next
            FileComp = 2 '不存在
        
        ElseIf cmCode = 2 Then
            For m = 1 To itemp
                If filenx = arrfilen(m, 1) Then
                    If filezx = arrfilesize(m, 1) Then
                        If filemx = arrfilemody(m, 1) Then
                            If fso.fileexists(arrfilep(m, 1)) = False Then '同时检查文件是否存在
                                n = m + 5
                                deledic(n) = "" '.Cells(n, 6).Value
                                idele = idele + 1
                                FileComp = 0
                                Exit Function
                            Else
                                FileComp = 3
                                Exit Function
                            End If
                        End If
                    End If
                End If
            Next
            FileComp = 4
        
        ElseIf cmCode = 3 Then
            For m = 1 To itemp
                If filep = arrfilep(m, 1) Then '当文件夹已经添加,就执行整个路径的比较,如果文件夹内的文件不发生大的变化就不检查文件是否存在
                    If F = 4 Then '文件夹内的文件已经发生了大的变化
                        If fso.fileexists(arrfilep(m, 1)) = False Then '同时检查文件是否存在
                            n = m + 5
                            deledic(n) = ""
                            idele = idele + 1
                            FileComp = 0
                            Exit Function
                        End If
                    End If
                    FileComp = 5
                    Exit Function
                End If
            Next
            FileComp = 6
        End If
    End With
End Function

Private Function DeleRFile(ByVal filepx As String, ByVal filext As String, ByVal Px As Byte)  '删除重复的文件
    If FileTest(filepx, filext) = 0 Then '文件处于关闭的状态
        If Px > 0 Then '存在非ansi字符
            fso.DeleteFile (filepx) '判断文件是否处于打开的状态
        Else
            DeleteFiles (filepx) '删除文件
        End If
        With ThisWorkbook.Sheets("目录")
            If ix = 1 Then '第一层文件夹
                .Cells(a, b + c + 2) = Now '删除操作后文件夹的时间发生变化
            Else
                .Cells(a - 1, b + c + 1) = Now '需要修改
            End If
        End With
        dl = dl + 1 '累计删除文件的数量
    End If
End Function

Private Function CheckPathAsWorkbook(ByVal strText, ByVal cmCode As Byte) As Boolean '判断文件夹/文件的来源和工作薄的位置关系
    Dim WBPath As String
    
    WBPath = ThisWorkbook.Path & "\"
    CheckPathAsWorkbook = False
    If cmCode = 1 Then     '表示文件
        strText = fso.GetFile(strText).ParentFolder & "\"
    ElseIf cmCode = 2 Then
        strText = strText & "\"     '表示文件夹
    End If
    If Len(strText) > Len(WBPath) Then
        If InStr(strText, WBPath) > 0 Then CheckPathAsWorkbook = True
    Else
        If InStr(WBPath, strText) > 0 Then CheckPathAsWorkbook = True
    End If
End Function
