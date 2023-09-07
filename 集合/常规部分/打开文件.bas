Attribute VB_Name = "打开文件"
Option Explicit
'---------------https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/declare-statement
Private Declare Function aShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function uShellExecute Lib "shell32.dll" Alias "ShellExecuteW" (ByVal hwnd As Long, _
ByVal lpOperation As String, ByVal lpFile As Long, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1
Public Conn As New ADODB.Connection
Public SQL As String
Public rs As ADODB.Recordset
'-----------------------------ado数据连接
Public Rng As Range  '注意使用的模块的层级与进程在模块和调用程序所处不同的模块
'--------------------搜索目录
Public Recentfile As String '最近阅读
Public Reditx As Byte '最近阅读
'---------------------窗体数据更新
Public QRfilepath As String '二维码地址
Public QRtextCN As String, QRtextEN As String, Barcodex As String '二维码,条形码
Public Turlx As String '浏览器地址
'-----------------------其他窗体/二维码/条形码
Public Statisticsx As Byte '用于统计代码
Public UF3Show As Byte 'uf3窗体启动标记
Public UF4Show As Byte 'uf4记录窗体
Public OpenFilex As Byte

Function OpenFile(ByVal filecode As String, ByVal filenx As String, ByVal filex As String, ByVal FilePath As String, _
ByVal spcode As Byte, ByVal Erransi As String, Optional ByVal cmCode As Byte) As Boolean '打开文件 'errcode用于标记文件路径是否存在空格的问题 'spode 0 一般请求,1表示来自控制板,2表示来自右键
    Dim i As Integer, k As Byte
    Dim fl As File, exepath As String, filepathx As String, recenttime As String
    
    On Error GoTo 100
    OpenFile = True '在执行较复杂的内容时返回执行的结果
    If filex = "xlsx" Then 'Excel无法打开同名文件
        If spcode = 1 Then           'spcode用于标识打开的来源是控制板还是直接在书库或者主界面打开
'            UserForm3.Hide
'            UserForm3.Show 0             '重新载入
        ' 考虑到userform在打开多个Excel关联文件较为麻烦的交互, 重新启用表格主界面作为备用的管理界面
            Unload UserForm3  '卸载掉窗体
            k = 1
            Call Rewds                    '当打开的文件为excel类的时候,重置窗口
        End If
        Workbooks.Open (FilePath) '这里需要注意窗口运行,代码中断的问题
    Else
        If Erransi = "ERC" Then   '判断文件的路径是否存在有特殊字符
            exepath = OpenBy(FilePath) '获取关联文件类型的程序     '这里可以改成cmd或者/powershell的方式打开文件
            If LenB(exepath) = 0 Then
                MsgBox "该文件类型不存在关联程序"
                Exit Function
            Else
                filepathx = """" & FilePath & """"             '可以防止文件路径存在空格的问题/shell命令没有碰到特殊字符的问题
                exepath = """" & exepath & """"
                Shell exepath & " " & filepathx, vbNormalFocus '注意程序路径后需要有空格,否则会出现找到程序的问题 '换一种打开方式
            End If
        Else
            ShellExecute 0, "open", FilePath, 0, 0, SW_SHOWNORMAL        '调用api的方式打开文件(不需要担心文件名空格的问题)
        End If
    End If
    If cmCode = 1 Then OpenFile = True: Exit Function '表示请求来源于控制板-比较
'    Call LockWorkSheet '确保表格可以处于编辑状态
'    ----------------------------------------选择打开的程序
    Rng.Offset(0, 11) = Now                    '打开文件的时间
    Rng.Offset(0, 12) = Rng.Offset(0, 12) + 1 '计算打开文件的次数
    
    Set fl = fso.GetFile(FilePath)
    With fl
        If .DateLastModified <> Rng.Offset(0, 6).Value Then '如果文件修改时间发生变化,文件的内容发生了变化
            Reditx = 1 '返回一个执行值用于更新窗体的属性
            Rng.Offset(0, 6) = .DateLastModified '文件的修改时间发生改变,即很大可能是文件的内容发生变化,即md5可能发生变化,大小发生变化
            Rng.Offset(0, 5) = .Size
            If .Size < 1048576 Then
                Rng.Offset(0, 7) = Format(.Size / 1024, "0.00") & "KB"
            Else
                Rng.Offset(0, 7) = Format(.Size / 1048576, "0.00") & "MB"
            End If
            filex = UCase(filex)
            If filex Like "EPUB" Or filex Like "MOBI" Or filex Like "PDF" Then Rng.Offset(0, 9) = GetFileHashMD5(FilePath)
        End If
        Set fl = Nothing
    End With
    '---------------检查文件的属性是否发生变化
    With ThisWorkbook.Sheets("主界面")         '主界面记录下信息
'        If .Range("w27").NumberFormatLocal <> "yyyy/m/d h:mm;@" Then Call DataSwitch '调整显示的格式
        If Len(.Range("u27").Value) > 0 Then
            If .Range("u27").Value = filecode Then .Range("w27") = Now: GoTo 1000 '如果目录已存在则不写入
        End If
        If Len(.Range("u27").Value) = 0 Then             '确保添加进来的值可以一直放在第一行
            .Range("p27") = Rng.Offset(0, 1)
            .Range("u27") = filecode
            .Range("w27") = Now
        Else
            For i = 33 To 28 Step -1
                .Range("p" & i) = .Range("p" & i - 1)
                .Range("u" & i) = .Range("u" & i - 1)
                .Range("w" & i) = .Range("w" & i - 1)
            Next
            .Range("p27") = Rng.Offset(0, 1)
            .Range("u27") = filecode
            recenttime = Now              '确保表格和窗口的时间完全一致'而不是直接使用now
            .Range("w27") = recenttime
            If spcode = 1 Then Recentfile = recenttime '用于记录最近打开文件的时间,判断是否需要更新窗口中的信息
        End If
    End With
    '----------------------------记录打开
1000
    RecordWrite filecode, Rng.Offset(0, 1).Value, Rng.Offset(0, 13).Value, Rng.Offset(0, 21).Value '打开文件记录
    If spcode = 2 Or k = 1 Then Set Rng = Nothing '控制板需要继续引用rng
    If UF3Show = 3 Then OpenFilex = 1 '-----------用于让窗体执行数据更新
    Exit Function '出错处理
100
    If Err.Number <> 0 Then
        Err.Clear
        Set Rng = Nothing
        OpenFile = False
    End If
End Function

Function SearchFile(ByVal filecode As String) '查找文件目录-全局变量
    With ThisWorkbook.Sheets("书库")
        If Len(filecode) = 0 Then Exit Function: Set Rng = Nothing
        Set Rng = .Range("b6:b" & .[b65536].End(xlUp).Row).Find(filecode, lookat:=xlWhole) '精确查找
    End With
End Function

Private Function RecordWrite(ByVal Unicode As String, ByVal filen As String, ByVal mfilen As String, ByVal idcode As String) '记录打开的操作
    Dim timea As String 'time As Date
    Dim FilePath As String
    
    If RecData = True Then
'        time = Format(Now, "yyyy年mm月dd日 hh:mm:ss")
        timea = Format(Recentfile, "ddd")
        SQL = "Insert into [打开记录$] (统一编码,文件名,主文件名,标识编码,时间,星期) Values ('" & Unicode & "', '" & filen & "', '" & mfilen & "', '" & idcode & "','" & Recentfile & "', '" & timea & "')" '#符号用于表示时间的变量 '#" & time & "#
        Conn.Execute (SQL)
    End If
End Function
