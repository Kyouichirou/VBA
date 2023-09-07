Attribute VB_Name = "文本分析"
Option Explicit
'需要细化
' 剔除掉常用的词汇
' 剔除重叠(aa)
' 剔除连续的----
' 剔除超出一定长度范围,但是没有"-"链接的字符串
' 剔除掉超长的(学科类专业名词例外, 如各类生物, 化学,根据各种命名法形成的超长词汇)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, _
ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
'-----------------------------------------https://fishc.com.cn/thread-70452-1-1.html
Private Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, ByVal lpFileMappigAttributes As Long, _
ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, _
ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
'---------------------------------------https://docs.microsoft.com/en-us/windows/win32/api/memoryapi/nf-memoryapi-mapviewoffile
'---------------------------------------https://docs.microsoft.com/en-us/windows/win32/api/memoryapi/nf-memoryapi-mapviewoffileex
Private Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress As Any) As Long
Private Const PAGE_READWRITE = &H4
Private Const FILE_MAP_READ = &H4
'--------------------------------------http://binaryworld.net/Main/ApiDetail.aspx?ApiId=5817
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
'''----------------------------------------------------
Private Type SAFEARRAY1D
    cDims As Integer      '数组的维度
    fFeatures As Integer  '用来描述数组如何分配和如何被释放的标志
    cbElements As Long    '数组元素的大小
    clocks As Long        '一个计数器，用来跟踪该数组被锁定的次数
    pvData As Long        '指向数据缓冲的指针--------------------关键所在
    rgsabound(0) As SAFEARRAYBOUND '述数组每维的数组结构，该数组的大小是可变的rgsabound是一个有趣的成员，它的结构不太直观。
                                   '它是数据范围的数组。该数组的大小依safearray维数的不同而有所区别。
                                   'rgsabound成员是一个SAFEARRAYBOUND结构的数组--每个元素代表SAFEARRAY的一个维。

End Type
'--------------------https://docs.microsoft.com/en-us/previous-versions/windows/embedded/ms912047(v=winembedded.10)
Public Enum LocalID
    zh_hk = 3076
    zh_ma = 5124
    zh_cn = 2052
    zh_sg = 4100
    zh_tw = 1028
    japan = 1041
    korea = 1042
    en_us = 1033
    en_uk = 2053
End Enum
'----------------codepage和localid涉及到vba所支持的语言相关
'vba内置字符串以Unicode的形式存在, 2个字节
'vba内置的函数大多调用ansi版本的api来实现功能
'lenb的速度比len快, 因为len计算出来的数据还需要 /2
'---------------------------------------------------
Public Enum CodePage '代码页, 可用于URLCodePage设置(winhttprequest)
    GB2312 = 20936   '简体中文
    GBK = 936        '中文扩展
    Big5 = 950       '繁体中文/港台
    GB18030 = 54936  '扩展版,包含绝大部分汉字(简繁)
    Shift_Jis = 932  '日文
    Ks_c_5601 = 949  '韩文
    IBM437_us = 437  '英文(US)
    UTF8 = 65001     'UTF-8
End Enum

'----------------------------------------------------------adodb.stream
Private Enum mReadText 'ado.stream读取文本的方式
    adReadAll = -1
    adReadLine = -2
End Enum
'指示只读权限
'指示读/写权限。
'与其他 *ShareDeny* 值（adModeShareDenyNone、adModeShareDenyWrite 或 adModeShareDenyRead）一起使用，以将共享限制传播给当前 Record 的所有子记录。
'如果 Record 没有子记录，将没有影响。如果它仅和 adModeShareDenyNone 一起使用，将产生运行时错误。
'但是，与其他值结合后，它可以和 adModeShareDenyNone 一起使用。例如，可以使用“adModeRead Or adModeShareDenyNone Or adModeRecursive”。
'允许其他人以任何权限打开连接?不拒绝其他人的读访问或写访问
'禁止其他人以读权限打开连接
'禁止其他人以写权限打开连接
'禁止其他人打开连接
'默认值?指示尚未设置或不能确定权限
'指示只写权限
Private Enum adConnectMode
    adModeRead = 1
    adModeReadWrite = 3        '常用
    adModeRecursive = &H400000
    adModeShareDenyNone = 16
    adModeShareDenyRead = 4
    adModeShareDenyWrite = 8
    adModeShareExclusive = 12
    adModeUnknown = 0
    adModeWrite = 2
End Enum

Private Enum adStreamType '指定数据类型
    adTypeBinary = 1     '二进制
    adTypeText = 2       '文本
End Enum
'-------------------------------------------------------------------adodb.stream

'使用内存映射方式查找大型文件中包含的字符串
Function FindTextInFile(ByVal strFileName As String, ByVal strText As String) As Long
    Dim hFile As Long, hFileMap As Long
    Dim nFileSize As Long, lpszFileText As Long, pbFileText() As Byte
    Dim ppSA As Long, pSA As Long
    Dim tagNewSA As SAFEARRAY1D, tagOldSA As SAFEARRAY1D
  
    hFile = CreateFile(strFileName, _
            GENERIC_READ Or GENERIC_WRITE, _
            FILE_SHARE_READ Or FILE_SHARE_WRITE, _
            0, _
            OPEN_EXISTING, _
            FILE_ATTRIBUTE_NORMAL Or FILE_ATTRIBUTE_ARCHIVE Or FILE_ATTRIBUTE_READONLY Or _
            FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_SYSTEM, _
            0) '打开文件
    If hFile <> 0 Then
        nFileSize = GetFileSize(hFile, ByVal 0&) '获得文件大小
        hFileMap = CreateFileMapping(hFile, 0, PAGE_READWRITE, 0, 0, vbNullString) '创建文件映射对象
        lpszFileText = MapViewOfFile(hFileMap, FILE_MAP_READ, 0, 0, 0) '将映射对象映射到进程内部的地址空间
          
        ReDim pbFileText(0) '初始化数组
        ppSA = VarPtrArray(pbFileText) '获得指向SAFEARRAY的指针的指针
        CopyMemory pSA, ByVal ppSA, 4 '获得指向SAFEARRAY的指针
        CopyMemory tagOldSA, ByVal pSA, Len(tagOldSA) '保存原来的SAFEARRAY成员信息
        CopyMemory tagNewSA, tagOldSA, Len(tagNewSA) '复制SAFEARRAY成员信息
        tagNewSA.rgsabound(0).cElements = nFileSize '修改数组元素个数
        tagNewSA.pvData = lpszFileText '修改数组数据地址
        CopyMemory ByVal pSA, tagNewSA, Len(tagNewSA) '将映射后的数据地址绑定至数组
        Dim m As Long, n As Long, k As Long, i As Long, p As Long
        Dim ibyte() As Byte
        ibyte = strText 'StrConv(strText, vbFromUnicode)
        m = UBound(ibyte)
        k = UBound(pbFileText)
        p = m - 1
        Dim x As Long
        x = 0
        For i = 0 To k
            If pbFileText(i) = ibyte(0) Then
                If pbFileText(i + m) = ibyte(m) Then
                    For n = 1 To p
                        If ibyte(n) = pbFileText(i + n) Then
                            x = x + 1
                            If x = p Then FindTextInFile = i: GoTo 100
                        Else
                            x = 0
                            Exit For
                        End If
                    Next
                End If
            End If
        Next
100
        CopyMemory ByVal pSA, tagOldSA, Len(tagOldSA) '恢复数组的SAFEARRAY结构成员信息
        Erase pbFileText '删除数组
          
        UnmapViewOfFile lpszFileText '取消地址映射
        CloseHandle hFileMap '关闭文件映射对象的句柄
    End If
    CloseHandle hFile '关闭文件
End Function

Function CheckFileKeyWordx(ByVal FilePath As String, ByVal Keyword As String) As Boolean '检查word文件是否包含关键词
    Dim wd As Object
    Dim reg As Object, Matches As Object
    '----------------------或者使用word find来查找https://docs.microsoft.com/zh-cn/office/vba/api/word.find.found
    '.found属性true,即表明查找到匹配项
    CheckFileKeyWordx = False
    Set wd = CreateObject(FilePath)   '执行速度较慢,基础速度要1s, 主要卡在创建word文件对象这一步
'---If InStr(1, wd.Content.Text, keyword, vbBinaryCompare) > 0 Then CheckFileKeyWordx = True 'instr在大文本的处理上, 速度远低于正则
    Set reg = CreateObject("VBScript.RegExp")
    With reg
        .Pattern = Keyword
        .Global = True
        .IgnoreCase = True
        Set Matches = .Execute(wd.Content.Text)
        If Matches.Count > 0 Then CheckFileKeyWordx = True
'        CheckFileKeyWordx = .test(keyword)
    End With
    wd.Close savechanges:=False
    Set reg = Nothing
    Set Matches = Nothing
    Set wd = Nothing
End Function

Sub StopTimer() '计时器 /理论上较高精度
With New Stopwatch
    .Restart
    Debug.Print FindTextInFile("C:\Users\adobe\Desktop\x31.txt", "我喜欢")
    .Pause
    Debug.Print Format(.Elapsed, "0.000000"); " seconds elapsed"
End With
End Sub

Private Function Words_Static(ByVal FilePath As String) As String() 'word文档可以直接使用word的属性来获取单词
    Dim objwd As Object
    Dim objwords As Object
    Dim i As Long, k As Long
    Dim dic As Dictionary
    Dim strTemp As String, strText As String
    Dim arr() As String
    Dim myreg As Object
    
    Set objwd = CreateObject(FilePath)
    Set objwords = objwd.Content.Words 'https://docs.microsoft.com/zh-cn/office/vba/api/word.words
    i = objwords.Count
    If i = 0 Then Exit Function
    Set myreg = CreateObject("VBScript.RegExp")
    With myreg
        .Pattern = "[a-z]+['|-|’]?[a-z]{1,}"
        .Global = True '不区分大小写
        .IgnoreCase = True
    End With
    Set dic = New Dictionary
    dic.CompareMode = vbTextCompare
    For k = 1 To i
        strTemp = Trim(objwords(k).Text)
        If Len(strTemp) > 1 Then
            strText = myreg.Replace(strTemp, "")
            If Len(strText) = 0 Then
                If dic.Exists(strTemp) Then
                    dic(strTemp) = dic(strTemp) + 1 '记录出现的次数
                Else
                    dic.Add strTemp, 1  '如果尚未出现,就添加值/item=1
                End If
            End If
        End If
    Next
    i = 0
    i = dic.Count
    If i = 0 Then Set dic = Nothing: Set objwords = Nothing: Exit Function
    i = i - 1
    ReDim arr(i, 1)
    ReDim Words_Static(i, 1)
    For k = 0 To i
        arr(k, 0) = dic.Keys(k)
        arr(k, 1) = dic.Items(k)
    Next
    Words_Static = arr
    Erase arr
    objwd.Close
    Set objwd = Nothing
    Set objwords = Nothing
    Set dic = Nothing
End Function

Sub WordAnalysis(ByVal FilePath As String, Optional ByVal strLen As Integer = 35) '单词分析/支持word/text文件等文本文件, 最好使用txt文档, 越简单的文件处理越快
    Dim textx As String, dic As Object
    Dim myreg As Object, Matches As Object, match As Object
    Dim arrTemp As Variant, i As Long, j As Long, temp As String, c As String, arr() As Integer, p As Long
    Dim wb As Workbook
    Dim wordapp As Object, ado As Object
    Dim newdc As Object, strx As String, strx1 As String
    Dim strTemp As String
    
    If fso.fileexists(FilePath) = False Then Exit Sub
    ThisWorkbook.Application.ScreenUpdating = False
    strx = Right$(FilePath, Len(FilePath) - InStrRev(FilePath, "."))
    strx = LCase(strx)
    If strx = "txt" Or strx Like "doc*" Then
        If strx <> "txt" Then
            Set wordapp = CreateObject("Word.Application") '读取word的内容非常慢(主要是创建word的速度非常慢)
            Set newdc = wordapp.documents.Add(FilePath)
            '或者Dim obj As Object
            'Set obj = CreateObject("C:\Users\*.docx")'可以直接创建doc文件的对象
            wordapp.Visible = False
            textx = newdc.Content.Text
            newdc.Close
            wordapp.Quit
            Set newdc = Nothing
            Set wordapp = Nothing
        Else
            Set ado = New ADODB.Stream
            With ado
                .Mode = 3 '读写权限
                .type = 2 '读取文本
                .CharSet = "us-ascii"  '关键, 不然读取的内容会出现乱码, 这里分析的是英文(纯英文)单词所以选ascii字符集(不包括中英混合),
                '-----------------------https://docs.microsoft.com/en-us/previous-versions/exchange-server/exchange-10/ms526296(v=exchg.10)
                .Open
                .LoadFromFile (FilePath) '加载文件
                textx = .ReadText()
                .Close
                Set ado = Nothing
            End With
        End If
    Else
        Exit Sub
    End If
    If Len(textx) = 0 Then Exit Sub
    Set dic = New Dictionary
    dic.CompareMode = vbTextCompare '不区分大小
    Set myreg = CreateObject("VBScript.RegExp") '正则表达式
    With myreg
        .Pattern = "[a-z]+['|-|’]?[a-z]{1,}"   '匹配要求, 字母,可以有" - ",长度大于2的单词,如果包含'或者-或者 ' 三种连字符,就完整输出, 如 what's, (小于(缺省)35的单词)"
                                               ' ?表示前面的符号可以出现一次或者不出现 |表示运算符 "或" +匹配多次,{1,} 连接的长度大于1
        .Global = True '不区分大小写           '注意chr(39)和chrw(8217)的符号的区别
        .IgnoreCase = True
        Set Matches = .Execute(textx)
    End With
    For Each match In Matches  '统计频次
        strTemp = match.Value
        If dic.Exists(strTemp) Then dic(strTemp) = dic(strTemp) + 1 Else dic.Add strTemp, 1
    Next
    Set Matches = Nothing
    p = dic.Count - 1
    If p < 5 Then MsgBox "数量太少不具备分析价值", vbInformation, "Tips": Exit Sub
'    For i = 0 To p - 1 '--------------单词排序 ,在Excel中不进行排序直接使用Excel自带的排序(如果输出内容到其他的载体, 再启用)
'        For j = i + 1 To p
'            If arrtemp(i) > arrtemp(j) Then
'                temp = arrtemp(i)
'                arrtemp(i) = arrtemp(j)
'                arrtemp(j) = temp
'            End If
'        Next
'    Next
    ReDim arr(p)
    ReDim arrTemp(p)
    For i = 0 To p             '获取出现的次数
       arrTemp = dic.Keys(i)
       arr(i) = dic.Items(i)
    Next
    Set wb = Workbooks.Add
    p = p + 1
    With wb            '---------------------创建表格输出分析结果
        .Worksheets(1).Name = "分析结果"
        With .Worksheets(1)
            .Cells(1, 1) = "单词:"
            .Cells(1, 2) = "出现次数:"
            .Cells(2, 2).Resize(p) = Application.Transpose(arr)
            .Cells(2, 1).Resize(p) = Application.Transpose(arrTemp)
            '--------------------------------------------------------数据写入
            .Cells(1, 5) = "累计分析单词总数:"
            .Cells(1, 6) = dic.Count
            .Cells(2, 5) = "生成时间:"
            .Cells(2, 6) = Format(Now, "yyyy-mm-dd")
            '------------------------------------------------------------Excel内置的排序功能
            With .sort
                .SortFields.Clear
                .SortFields.Add key:=Range("B1"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                .SetRange Range("A2:B" & p + 1)
                .Header = xlGuess
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            If p > 25 Then p = 25
            p = p + 1
            For i = 2 To p
                .Cells(i, 1) = WorksheetFunction.Proper(.Cells(i, 1).Value) '将这部分的数据的首个字母转为大写,其余部分转为小写
            Next
            CreatChart wb, p '创建分析图表
            .Cells(1, 5).ColumnWidth = 18
            .Cells(1, 6).ColumnWidth = 18 '调整显示的格子的大小
            .Cells(1, 5).HorizontalAlignment = xlRight '调整排列方式
            .Cells(2, 5).HorizontalAlignment = xlRight
            .Cells(1, 3).Select
        End With
        strx1 = Left$(FilePath, InStrRev(FilePath, ".") - 1) '文件保存的文件
        strx1 = strx1 & Format(Now, "yyyymmddhhmmss")
        strx = strx1 & ".xlsx"
        If fso.fileexists(strx) = True Then fso.DeleteFile strx
        If Err.Number = 70 Then Err.Clear: strx = strx1 & CStr(RandNumx(1000)) & ".xlsx"
        .SaveAs strx                           '将文件保存到同位置上
    End With
    Erase arr
    Erase arrTemp
    Set dic = Nothing
    Set myreg = Nothing
    ThisWorkbook.Application.ScreenUpdating = True
End Sub

Private Sub CreatChart(ByVal wb As Workbook, ByVal numx As Byte) '制作图表
    Dim Shx As Shape
    Dim Cha As Chart
    Dim dTextx As Shape, rTextx As Shape, aTextx As Shape
    
    Set Shx = wb.Sheets(1).Shapes.AddChart2(201, xlColumnClustered, 250, 60, 720, 350) 'top=60,height=350,width=720
    Set Cha = Shx.Chart
    With Cha
        .SetSourceData Source:=wb.Sheets(1).Range("A1:B" & numx)
        '------------------------------------------------------数据源
        numx = numx - 1
        .ChartTitle.Text = "Word Top" & CStr(numx) '标题
        '--------------------------------------标题
        .ApplyLayout (9) '----------------图表的类型 '通过录制获取到的值是6(不是目标需要的图表类型)
        '----------图表的风格(注意不是类型)
        .PlotArea.Select
        .PlotArea.Width = 634
        .PlotArea.Left = 24
        .PlotArea.Top = 39
        .PlotArea.Height = 276 '-------作图区域大小调整
        '-----------------------------------------图表作图区域
        .FullSeriesCollection(1).ApplyDataLabels '柱状图显示数据
        '-------------------------------------------------------柱状图柱子上显示数据
        .Axes(xlValue).AxisTitle.Format.TextFrame2.TextRange.Font.Size = 9.5
        .Axes(xlValue, xlPrimary).AxisTitle.Top = 128
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "单词出现次数" 'y轴信息
        '------------------------------------------------------------------------Y轴调整
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = ChrW(9670) & "单词"
        .Axes(xlCategory, xlPrimary).AxisTitle.Left = 658
        .Axes(xlCategory, xlPrimary).AxisTitle.Top = 298
        .Axes(xlCategory).TickLabels.Font.Size = 11 '调整x轴下文字的大小(录制的宏是错的)
        '--------------------------------------------------------------------------------------------X轴调整
        Set dTextx = .Shapes.AddTextbox(msoTextOrientationHorizontal, 612, 0, 108, 16) '添加文本框
        Set rTextx = .Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 332, 255, 16) '添加文本框
        Set aTextx = .Shapes.AddTextbox(msoTextOrientationHorizontal, 612, 332, 108, 16) '添加文本框
        '------------额外外加文本框
    End With
    With dTextx '时间
        .TextFrame.Characters.Text = "Date: " & Format(Now, "yyyy-mm-dd") '文本框写入信息
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignRight '文本框内容右对齐
    End With
    With rTextx '来源
        .TextFrame.Characters.Text = "Resource: File Analysis" '文本框写入信息Wallstreet Journal
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft '文本框内容作对齐
    End With
    With aTextx '作者
        .TextFrame.Characters.Text = "Drawing By: HLA" '文本框写入信息
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignRight '文本框内容右对齐
    End With
    '-------------------------文本框调整
    wb.Sheets(1).ChartObjects(1).Placement = xlFreeFloating '图表的位置不会因为表格而发生变化
    With Shx.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
    End With
   '-----------------------------边框调整
    Set dTextx = Nothing
    Set aTextx = Nothing
    Set rTextx = Nothing
    Set Shx = Nothing
    Set Cha = Nothing
End Sub

Private Function Min(ByVal one As Integer, ByVal two As Integer, ByVal three As Integer) As Integer
    Min = one
    If (two < Min) Then Min = two
    If (three < Min) Then Min = three
End Function
 
Private Function CompString(ByVal str1 As String, ByVal str2 As String, ByVal n As Integer, m As Integer) As Integer
    Dim i, j As Integer, p As Integer, q As Integer
    Dim ch1, ch2 As String
    Dim arr() As Integer
    Dim temp As Byte
    
    If (n = 0) Then CompString = m
    If (m = 0) Then CompString = n
    ReDim arr(n + 1, m + 1)
    arr(0, 0) = 0
    For i = 1 To n
        arr(i, 0) = i
        ch1 = Mid(str1, i, 1)
        For j = 1 To m
            arr(0, j) = j
            ch2 = Mid(str2, j, 1)
            If (ch1 = ch2) Then
                temp = 0
            Else
                temp = 1
            End If
            p = i - 1
            q = j - 1
            arr(i, j) = Min(arr(p, j) + 1, arr(i, q) + 1, arr(p, q) + temp)
        Next
    Next
    CompString = arr(n, m)
End Function
 
Function Similar(ByVal str1 As String, ByVal str2 As String) As Double '顺序字符串相似度比较
    Dim ldint As Integer
    Dim i As Integer, k As Integer
    Dim strLen As Integer
    
    i = Len(str1)
    k = Len(str2)
    ldint = CompString(str1, str2, i, k)
    If (i >= k) Then
        strLen = i
    Else
        strLen = k
    End If
    Similar = 1 - ldint / strLen
End Function

Function CheckFileKeyWordB(ByVal FilePath As String, ByVal Keyword As String, Optional ByVal cmCode As Byte = 0) As Boolean '检查txt文件内是否存在指定的关键词
    Dim obj As Object
    Dim strx As String * 1024 '定长字符串
    
    CheckFileKeyWordB = False
    Set obj = fso.OpenTextFile(FilePath, ForReading) '注意文本的编码, ansi/Unicode/uft8的区别
    With obj
        Do While Not .AtEndOfStream                    'binary模式为0,text为1, data为-2, 需要区分大小时,使用text
            strx = .Read(1024)                   '这种方法偏慢,处理一个18M的文件大概需要1.5s/vbtext,vbinary 二进制的表速度更快,可以达到0.95s左右
            If InStr(1, strx, Keyword, cmCode) > 0 Then CheckFileKeyWordB = True: Exit Do
        Loop
        .Close
    End With
    Set obj = Nothing
End Function

Sub dkslla()
'With New Stopwatch
'    .Restart
'    Debug.Print CheckFileKeyWord("C:\Users\adobe\Desktop\x31.txt", "我喜欢", 0, 4)
'    .Pause
'    Debug.Print Format(.Elapsed, "0.000000"); " seconds elapsed"
'End With
Dim ad As New ADODB.Stream
Dim arr() As Byte
With ad
.Mode = 3
.type = 1
.Open
.LoadFromFile "C:\Users\adobe\Desktop\testcode\utf_nobom.txt"
arr = .Read(-1)
.Close
End With
UEFCheckUTF8NoBom arr
Set ad = Nothing
End Sub

Private Function UEFCheckUTF8NoBom(ByRef bufAll() As Byte)
    
    Dim i As Long
    Dim cOctets As Long         '可以容纳UTF-8编码字符的字节大小 4bytes
    Dim bAllAscii As Boolean    '如果全部为ASCII，说明不是UTF-8
    Dim fmt
    bAllAscii = True
    cOctets = 0
    
    For i = 0 To UBound(bufAll)
        If (bufAll(i) And &H80) <> 0 Then
            'ASCII用7位储存，最高位为0，如果这里相与非0，就不是ASCII
            '对于单字节的符号，字节的第一位设为0，后面7位为这个符号的unicode码。
            '因此对于英语字母，UTF-8编码和ASCII码是相同的
            bAllAscii = False
        End If
        
        '对于n字节的符号（n>1），第一个字节的前n位都设为1，第n+1位设为0，后面字节的前两位一律设为10
        'cOctets = 0 表示本字节是leading byte
        If cOctets = 0 Then
            If bufAll(i) >= &H80 Then
                '计数：是cOctets字节的符号
                Do While (bufAll(i) And &H80) <> 0
                    'bufAll(i)左移一位
                    bufAll(i) = ShLB_By1Bit(bufAll(i))
                    cOctets = cOctets + 1
                Loop
                
                'leading byte至少应为110x xxxx
                cOctets = cOctets - 1
                If cOctets = 0 Then
                    '返回默认编码
                    fmt = "UEF_ANSI"
                    Exit Function
                End If
            End If
        Else
            '非leading byte形式必须是 10xxxxxx
            If (bufAll(i) And &HC0) <> &H80 Then
                '返回默认编码
                fmt = "UEF_ANSI"
                Exit Function
            End If
            '准备下一个byte
            cOctets = cOctets - 1
        End If
    
    Next i
    
    '文本结束.  不应有任何多余的byte 有即为错误 返回默认编码
    If cOctets > 0 Then
        fmt = "UEF_ANSI"
        Exit Function
    End If
    
    '如果全是ascii.  需要注意的是使用相应的code pages做转换
    If bAllAscii = True Then
        fmt = "UEF_ANSI"
        Exit Function
    End If
    
    '修成正果 终于格式全部正确 返回UTF8 No BOM编码格式
    fmt = "UEF_UTF8NB"
    
End Function

Private Function ShLB_By1Bit(ByVal Byt As Byte) As Byte

'‘把BYTE類型變量左移1位的函數，參數Byt是待移位的字節，函數返回移位結果
'
'‘(Byt And &H7F)的作用是屏蔽最高位。 *2：左移一位

ShLB_By1Bit = (Byt And &H7F) * 2

End Function

Function CheckFileKeyWord(ByVal FilePath As String, ByVal Keyword As String, Optional ByVal cmCode As Byte = 0, _
Optional ByVal iCodex As Byte = 1, Optional ByVal isIgnoreCase As Boolean = True) As Boolean
    '----------------------------cmcode指定比较的类型,cmcodex指定读取的文件编码
    Const defBuffer As Long = 131072 '1024 * 128, 128K
    Const chBuffer As Long = 4096
    Dim oAstream As Object
    Dim Codex As String
    Dim sBuffer As String   '* 1024 '定长字符串有助于加快运算的速度
    Dim iPostion As Long, iBuffer As Long
    Dim iType As Byte
'    Dim cr As New cRegex
    '----------------http://wsh.style-mods.net/ref_stream/readtext.htm
    '中文,考虑编码的变量
    '英文,考虑编码,考虑大小写(二进制的方法会区分大小写) '不建议采用文本比较
    CheckFileKeyWord = False
    Select Case iCodex
        Case 1: Codex = "gb2312" '中文/ansi '最常用的编码
        Case 2: Codex = "us-ascii" '纯英文
        Case 3: Codex = "uft-8"
        Case 4: Codex = "unicode"
    End Select
    Set oAstream = New ADODB.Stream
    With oAstream
        .Mode = 3                   '读写权限
        If LenB(Codex) = 0 Then
            Dim tBytes() As Byte
            .type = adStreamType.adTypeBinary
            .Open
            .LoadFromFile FilePath
            iSize = .Size           '文件大小
            If iSize < 3 Then Exit Function
            tBytes = .Read(3)
            .Close
            If tBytes(0) = 239 Then
                If tBytes(1) = 187 And tBytes(2) = 191 Then Codex = "utf-8"
            ElseIf tBytes(0) = 255 Then
                If tBytes(1) = 254 Then Codex = "unicode"
            ElseIf tBytes(0) = 254 Then
                If tBytes(1) = 255 Then Codex = "utf-8"
            Else
            If iSize < chBuffer Then iBuffer = -1 Else iBuffer = chBuffer
            .Position = 0
            If iBuffer > 0 Then ReDim Bytes(iBuffer) Else ReDim tBytes(iSize)
            tBytes = .Read(iBuffer)                                           '截取部分内容用于判断具体的是否为utf-8 without BOM
            Check_Unicode (tBytes)
            End If
        End If
        .type = adStreamType.adTypeText '读取文本
        .CharSet = Codex            '关键, 不然读取的内容会出现乱码
        .Open                       'https://docs.microsoft.com/en-us/previous-versions/exchange-server/exchange-10/ms526296(v=exchg.10)
        .LoadFromFile (FilePath)    '加载文件
        
        sBuffer = Space$(iSize)     '核心关键的一步, 对整个处理速度起到核心作用, 大字符串的构造
        If iSize > defBuffer Then   '字符串的缓冲区大小为1024 * 64, 64K, 65536
            iBuffer = defBuffer
        Else
            iBuffer = mReadText.adReadAll
        End If
        iPostion = 1                '1.64m的文件大小获取到结果仅需要0.0065-0.0070之间
        .Position = 0               '指定流开始处到当前位置的偏移字节数。默认值为 0，表示流中的第一个字节。
        Do Until .EOS = True        'eos表示文件流的末尾 '测试显示过小或过大的值对速度都有明显的影响, 1024-131072之间的数据的效果是最为明显的
            Mid$(sBuffer, iPostion, iSize) = .ReadText(iBuffer)
            iPostion = iBuffer + i
                                    '这里需要注意, 不能整体读取, 如果文件过大将会导致速度大幅度下降,以1.64m文件读取为例, 速度下降超过20倍
        Loop
        .Close
    End With
    Set oAstream = Nothing
    If InStr(1, sBuffer, Keyword, cmCode) > 0 Then
        CheckFileKeyWord = True
    Else
        If isIgnoreCase = False Then
            If InStr(1, sBuffer, Keyword, cmCode) > 0 Then CheckFileKeyWord = True
        End If
    End If
    sBuffer = vbNullString
End Function



Sub dkkd()
Dim cr As New cRegex
With cr
cr.oReg_Initial
cr.oReg_Pattern = "ab"
cr.oReg_Text = "babb"
Debug.Print cr.cTest
End With
Set cr = Nothing
End Sub

Function CheckFileKeyWordC(ByVal FilePath As String, ByVal Keyword As String, Optional ByVal blockSize As Long = 131072) As Boolean '二进制判断是否包含关键词
    Dim arr() As Byte
    Dim arrx() As Byte
    Dim arrTemp() As Byte
    Dim ado As Object

    CheckFileKeyWordC = False
    Set ado = New ADODB.Stream
    ReDim arrx(blockSize - 1)
    With ado
        .Mode = 3
        .type = adTypeBinary
        .Open
        .Position = 0
        .LoadFromFile FilePath
        arrTemp = .Read(3)     '可以通过这个特性, 跳过BOM, 读取之后,将在3的位置之后开始读取
        '-------------------https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/mid-function
        If AscB(MidB(arrTemp, 1, 1)) = &HEF And AscB(MidB(arrTemp, 2, 1)) = &HBB And AscB(MidB(arrTemp, 3, 1)) = &HBF Then '判断文件的编码类型
            CheckFileKeyWordC = CheckFileKeyWord(FilePath, Keyword, , 3)
            .Close
            Set ado = Nothing
            Exit Function '"utf-8"  'uft8的处理暂时找不到合适的处理方法
        ElseIf AscB(MidB(arrTemp, 1, 1)) = &HFF And AscB(MidB(arrTemp, 2, 1)) = &HFE Then
            arr = Keyword '"unicode"
        Else
            arr = StrConv(Keyword, vbFromUnicode) '"ANSI/gb2312"
        End If
         .Position = 0 '注意这里读取数据之后重新恢复读取数据开始的位置
        Do Until .EOS = True
            arrx = .Read(blockSize)
            If InStr(1, arrx, arr, vbBinaryCompare) > 0 Then CheckFileKeyWordC = True: Exit Do
        Loop
        .Close
    End With
    Erase arr: Erase arrx
    Set ado = Nothing
End Function

Function CheckTextCode(ByVal FilePath As String) As String '检测文件的编码类型
'---------------https://blog.csdn.net/hongsong673150343/article/details/88584753
'  ANSI无格式定义
'  EFBBBF    UTF-8
'  FFFE      UTF-16/UCS-2, little endian
'  FEFF      UTF-16/UCS-2, big endian
'  FFFE 0000 UTF-32/UCS-4, little endian
'  0000 FEFF UTF-32/UCS-4, big endian
' 需要注意的是utf, 在uft8下, ascii编码的数据也是以1位字节存储的, 所以读取这些数据的时候无差别, Unicode统一采用2位, 字节, 还需要注意utf-8 无bom头的情况
'- ------------------------------------------- -
    Dim arr() As Byte
    Dim ado As Object
    Set ado = New ADODB.Stream
    With ado
        .type = 1
        .Mode = 3
        .Open
        .Position = 0
        .LoadFromFile FilePath
        arr = .Read(3)            'ascii码239,187,191就是BOM头的 EF BB BF
        If AscB(MidB(arr, 1, 1)) = &HEF And AscB(MidB(arr, 2, 1)) = &HBB And AscB(MidB(arr, 3, 1)) = &HBF Then
            CheckTextCode = "utf-8"
        ElseIf AscB(MidB(arr, 1, 1)) = &HFF And AscB(MidB(arr, 2, 1)) = &HFE Then '255 '254
            CheckTextCode = "unicode"
        Else
            CheckTextCode = "gb2312" 'ansi
        End If
        .Close
    End With
    Set ado = Nothing
End Function

Function StrChinese(ByVal strText As String) As Boolean '判断字符串是否包含中文, 可以区分中文汉字和日文汉字(不完全测试, 片假名等), 不受中文符号影响
    strText = StrConv(strText, vbNarrow)                'vbNarrow 将字符串中双字节字符转成单字节字符
    StrChinese = IIf(Len(strText) < LenB(StrConv(strText, vbFromUnicode)), True, False) '如果只是转换大小写, 不建议使用strconv函数, 可以单独用lcase, ucase函数
End Function

'正向否定预查(negative assert)，在任何不匹配pattern的字符串开始处匹配查找字符串。这是一个非获取匹配，
'也就是说，该匹配不需要获取供以后使用。例如"Windows(?!95|98|NT|2000)"能匹配"Windows3.1"中的"Windows"，
'但不能匹配"Windows2000"中的"Windows"。预查不消耗字符，也就是说，在一个匹配发生后，在最后一次匹配之后立即开始下一次匹配的搜索，而不是从包含预查的字符之后开始。
Sub Book_Analysis(ByVal FilePath As String)
    Dim arr() As Long
    Dim arrx
    Dim cre As New cRegex
    Dim strText As String
    Dim ado As New ADODB.Stream
    Dim i As Byte
'    Dim wb As Workbook
    
    DisEvents
'    arrx = Sheet8.Range("a1:a13").Value
'    ReDim arr(1 To 13)
    With ado
        .CharSet = "gb2312"
        .Mode = adModeReadWrite
        .type = adTypeText
        .Open
        .LoadFromFile FilePath
        strText = .ReadText
        .Close
    End With
    Set ado = Nothing
'    With cre
'        For i = 1 To 13
'            .oReg_Initial arrx(i, 1)
'            .sMatch strText
'            arr(i) = .sFirst_Index
'        Next
'    End With
'    strText = ""
     '"桐原(?!洋介|弥生子)亮司?" , 匹配桐原或者桐原亮司, 但是不匹配桐原洋介和桐原弥生子
     '(唐泽|西本)?(?!文代|礼子)(雪穗) , 匹配西本雪穗,唐泽雪穗或者雪穗,但是不匹配唐泽礼子, 不匹配西本文代
     '笹垣(润三)?, 匹配笹垣或者笹垣润三
     With cre
        .oReg_Initial "笹垣(润三)?"
        .xMatch strText
        arr = .aFirst_Index
     End With
    Set cre = Nothing
    Sheet8.Cells(1, 4).Resize(UBound(arr) + 1, 1) = Application.Transpose(arr)
'    Set wb = Workbooks.Add
'    With wb.Sheets(1)
'        .Cells(1, 1).Resize(44, 1) = arrx
'        .Cells(1, 2).Resize(44, 1) = wb.Application.Transpose(arr)
'    End With
'    Set wb = Nothing
    EnEvents
End Sub

Sub dkklldso()
Book_Analysis "C:\Users\adobe\Desktop\白夜行.txt"
End Sub


