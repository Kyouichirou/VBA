VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "控制板"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   18195
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
'-------------------------------------------https://blog.csdn.net/softman11/article/details/6124345 '字符编码
'API 为窗口创建最小化按钮 ' _符号表示换行
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
  ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
  ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" ( _
  ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
  ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
  ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetActiveWindow Lib "user32.dll" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
  ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

'Constants
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const GWL_EXSTYLE = (-20)
Private Const HWND_TOP = 0
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_SHOWWINDOW = &H40
Private Const WS_EX_APPWINDOW = &H40000
Private Const GWL_STYLE = (-16)
Private Const WS_MINIMIZEBOX = &H20000
Private Const SWP_FRAMECHANGED = &H20
Private Const WM_SETICON = &H80
Private Const ICON_SMALL = 0&
Private Const ICON_BIG = 1&
'-------------------------------------------------------------------------------------添加最小化按钮
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const STILL_ALIVE = &H103
Private Const INFINITE = &HFFFF '(这个参数要注意, 等待的时间)
Private ExitCode As Long
Private hProcess As Long
Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long '判断数组是否完成初始化
'-----------------------控制shell执行
Private Const BdUrl As String = "https://www.baidu.com"

Dim NewM As Boolean '用于控制textbox菜单
Dim filepath5 As String, folderoutput As String 'md5/销毁-输出列表
Dim filepathc As String, folderpathc As String '解压文件
Dim filepathset As String, folderpathset As String '设置
'-------------------------------------------------------------------------------为防止非ansi字符的干扰,需要临时存储值到变量中(不是直接使用Textbox,或者label上的值)
Dim arrcompress() As String, flc As Byte, arrfilez() As Double '工具-解压文件
Dim imgx As Byte '用于控制图片控件
Dim pgx As Byte, pgx1 As Byte, pgx2 As Byte, pgx3 As Byte, pgx4 As Byte, pgx5 As Byte, pgx6 As Byte, pgx7 As Byte '多页面切换时控制控件的内容的生成
'----------------------------------------------------------------
Dim browser1 As Byte, browserkey As String '用于控制浏览器控件
Dim arraddfolder() As String
Dim docmx As Integer '书库的最后一行
Dim arrax() As Variant '编号, 文件名,文件扩展名,文件路径,文件位置
Dim arrbx() As Variant '打开次数
Dim arrsx() As Variant '推荐指数/评分
Dim arrux() As Variant '标签1/标签2
Dim spyx As Integer '搜索存储书库列表值
Dim storagex As String
Dim searchx As Byte '文本搜索
'----------------------------------------------搜索/用于存储需要搜索的内容
Dim voicex As Byte '判断单词是否查询发音
Dim vbsx As Byte, vbfilex As String '存储vbs路径 '将经常调用的参数存储到内存中去
'-------------------------------------------------------------测试搜索
Dim wm As Object '创建临时的Windows media用于播放背景音乐 ' WindowsMediaPlayer
'-----------------------------------------------------
Dim arrlx() As String 'treeview临时存储key
Dim arrch() As String 'treeview存储选中文件的子项目的key
Dim ich As Byte 'treeview查找nodes的所有子项
Dim s As Byte 'treeviewnodes数组的临时值
'--------------------------------------treeview相关
Dim arrTemp() As Integer, arrtemp2() As String, arrtemp1() As String, arrtemp3() As String '存储单词训练的试题和结果
Dim listnum As Byte '单词列表-随机数
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) '时间 -单词训练
Private Declare Function timeGetTime Lib "winmm.dll" () As Long '时间 -单词训练
Dim Flagpause As Boolean, FlagStop As Boolean, Flagnext As Boolean '控制执行的变量-单词训练
'--------------------------------------------------------单词训练
Dim clickx As Byte '控制checkbox.click事件的执行
'--------------------------------------------------------------------------------------最小化ico
Private Sub AddIcon() '增加显示图标
    Dim hwnd As Long
    Dim lngRet As Long
    Dim hIcon As Long
    'hIcon = Image1.Picture.Handle
    'hWnd = FindWindow(vbNullString, Me.Caption)
    lngRet = SendMessage(hwnd, WM_SETICON, ICON_SMALL, ByVal hIcon)
    'lngRet = SendMessage(hWnd, WM_SETICON, ICON_BIG, ByVal hIcon)
    lngRet = DrawMenuBar(hwnd)
End Sub

Private Sub AddMinimiseButton() '窗体增加最小化按钮
    Dim hwnd As Long
    
    hwnd = GetActiveWindow
    Call SetWindowLong(hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) Or WS_MINIMIZEBOX)
    Call SetWindowPos(hwnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Private Sub AppTasklist(myForm) '添加到任务栏
    Dim wStyle As Long
    Dim Result As Long
    Dim hwnd As Long
    
    hwnd = FindWindow(vbNullString, myForm.Caption)
    wStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
    wStyle = wStyle Or WS_EX_APPWINDOW
    Result = SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or _
      SWP_HIDEWINDOW)
    Result = SetWindowLong(hwnd, GWL_EXSTYLE, wStyle)
    Result = SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or _
      SWP_SHOWWINDOW)
End Sub
'---------------------------------------------------------------------------------------------------最小化ico

Private Sub CheckBox10_Click() '设置-禁止显示删除文件提示
    If clickx = 1 Then clickx = 0: Exit Sub
    If Me.CheckBox10.Value = True Then
        ThisWorkbook.Sheets("temp").Range("ab37").Value = 1
    Else
        ThisWorkbook.Sheets("temp").Range("ab37").Value = ""
    End If
End Sub

Private Sub CheckBox11_Click() '设置-自动计算md5
    If clickx = 1 Then clickx = 0: Exit Sub
    With Me
        If .CheckBox11.Value = True Then
            ThisWorkbook.Sheets("temp").Range("ab35") = 1
            .CheckBox12.Enabled = True
        Else
            ThisWorkbook.Sheets("temp").Range("ab35") = ""
            .CheckBox12.Enabled = False
            ThisWorkbook.Sheets("temp").Range("ab36") = ""
        End If
    End With
End Sub

Private Sub CheckBox12_Click() '设置-包括电子书
    If clickx = 1 Then clickx = 0: Exit Sub
    If Me.CheckBox12.Value = True Then
        ThisWorkbook.Sheets("temp").Range("ab36") = 1
    Else
        ThisWorkbook.Sheets("temp").Range("ab36") = ""
    End If
End Sub

Private Sub CheckBox13_Click() '文件保护
    Dim yesno As Variant, strx As String, strx1 As String, strx2 As String
    
    If Fileptc > 0 Then Fileptc = 0: Exit Sub '防止向checkbox赋值触发事件
    With Me
        strx = .Label29.Caption
        If Len(strx) = 0 Or .Label55.Visible = True Then .CheckBox13.Value = False: Exit Sub
        With ThisWorkbook.Sheets("temp")
            If Len(.Cells(41, "ab").Value) = 0 Then
               strx2 = ThisWorkbook.Path & "\protect"
                .Cells(41, "ab") = strx2
                If fso.folderexists(strx2) = False Then fso.CreateFolder strx2 '创建文件夹
            Else
                strx2 = .Cells(41, "ab").Value
                If fso.folderexists(strx2) = False Then fso.CreateFolder strx2
            End If
            strx1 = strx2 & "\" & Filenamei '.Label23.Caption
        End With
        
        If .CheckBox13.Value = True Then
            If Len(.Label76.Caption) > 52428800 Then
                yesno = MsgBox("文件太大,执行较慢,是否继续?", vbYesNo, "Warning")
                If yesno = vbNo Then Exit Sub
            End If
            
            If FileStatus(strx, 2) = 4 Then '检查文件是否存在
                Rng.Offset(0, 32) = 1
            Else
'                .Label55.Visible = ture
'                .TextBox1.Text = ""
                .Label57.Caption = "文件不存在"
                DeleFileOverx strx
            End If
        Else
            SearchFile strx
            If Rng Is Nothing Then
'                .Label55.Visible = ture
'                .TextBox1.Text = ""
                .Label57.Caption = "文件不存在"
                DeleFileOverx strx
            End If
            Rng.Offset(0, 32) = ""
            If fso.fileexists(strx1) = True Then fso.DeleteFile (strx1) '取消文件保护,复制的文件将被删除掉
        End If
    End With
    Set Rng = Nothing
End Sub

Private Sub CheckBox16_Click() '解压-删除源文件
    Dim yesno As Variant
    
    If Me.CheckBox16.Value = True Then
        yesno = MsgBox("注意: 此项勾选,不管文件是否解压成功,源文件都会被删除" & vbCr _
               & "(如:加密文件的密码输入错误)是否继续?", vbYesNo, "Warning!!!")
        If yesno = vbNo Then Me.CheckBox16.Value = False
    End If
End Sub

Private Sub CheckBox17_Click()
    If clickx = 1 Then clickx = 0: Exit Sub
    With Me.CheckBox17
        If .Value = True Then
            ThisWorkbook.Sheets("temp").Cells(43, "ab") = 1
        Else
            ThisWorkbook.Sheets("temp").Cells(43, "ab") = ""
        End If
    End With
End Sub

Private Sub CheckBox18_Click() '内置浏览器
    If clickx = 1 Then clickx = 0: Exit Sub
    With ThisWorkbook.Sheets("temp")
        If Len(.Cells(6, "ab")) = 0 Then Me.CheckBox18.Value = False: Exit Sub
        If Me.CheckBox18.Value = ture Then
            .Cells(45, "ab") = 1
        Else
            .Cells(43, "ab") = ""
        End If
    End With
End Sub

Private Sub CheckBox19_Click() '文件的销毁方式
    If clickx = 1 Then clickx = 0: Exit Sub
    With ThisWorkbook.Sheets("temp")
        If Me.CheckBox19.Value = True Then
            .Cells(50, "ab") = 1
        Else
            .Cells(50, "ab") = ""
        End If
    End With
End Sub

Private Sub CheckBox21_Click() '搜索主文件名
    If Me.CheckBox21.Value = True Then Me.CheckBox22.Value = False
End Sub

Private Sub CheckBox22_Click() '搜索文件名
    If Me.CheckBox22.Value = True Then Me.CheckBox21.Value = False
End Sub

Private Sub CheckBox23_Click() '同步添加封面
    If clickx = 1 Then clickx = 0: Exit Sub
    With ThisWorkbook.Sheets("temp")
        If Me.CheckBox23.Value = True Then
            .Cells(53, "ab") = 1
        Else
            .Cells(53, "ab") = ""
        End If
    End With
End Sub

Private Sub CheckBox25_Click() '不区分大小写
    With Me
    If .CheckBox25.Value = True Then
        .CheckBox26.Enabled = False
    Else
        .CheckBox26.Enabled = True
    End If
    End With
End Sub

Private Sub CheckBox26_Click() '极速模式
    With Me
    If .CheckBox26.Value = True Then
        .CheckBox25.Enabled = False
    Else
        .CheckBox25.Enabled = True
    End If
    End With
End Sub

Private Sub CheckBox27_Click() '调试模式
    If clickx = 1 Then clickx = 0: Exit Sub
    With ThisWorkbook.Sheets("temp")
        If Me.CheckBox27.Value = True Then
            .Cells(54, "ab") = 1
        Else
            .Cells(54, "ab") = ""
        End If
    End With
End Sub

Private Sub CommandButton106_Click() '文件夹-显示导出列表
    With Me
        If .CommandButton106.Caption = "显示导出列表" Then
            .CommandButton106.Caption = "隐藏导出列表"
            .ListBox6.Visible = True
        Else
            .ListBox6.Visible = False
            .CommandButton106.Caption = "显示导出列表"
        End If
    End With
End Sub

Private Sub CommandButton107_Click() '文件夹-导出列表-默认导出为text文档,导出目录为documents 'https://docs.microsoft.com/zh-TW/office/vba/Language/Reference/User-Interface-Help/createtextfile-method
    Dim fl As Object, strx As String
    Dim i As Integer, k As Integer
    
    With Me.ListBox6
        k = .ListCount
        If k = 0 Then Exit Sub
        strx = Environ("UserProfile") & "\Desktop\" & CStr(Format(Now, "yyyymmddhhmmss")) & ".txt" '按照时间创建txt文件
        If fso.fileexists(strx) = False Then
            Set fl = fso.CreateTextFile(strx, True)
            k = k - 1
            For i = 0 To k
                fl.WriteLine .List(i, 0) & "  " & .List(i, 1) '将列表的数据写入
            Next
            fl.Close
            Me.Label57.Caption = "操作成功"
        Else
            MsgBox "目录下已存在相同的文件", vbOKOnly, "Warning"
            Exit Sub
        End If
    End With
    Set fl = Nothing
End Sub

Private Sub CommandButton108_Click() '文件夹-导出文件
    Dim strfolder As String
    Dim i As Integer, k As Integer, strx3 As String, strx4 As String
    Dim strx As String, strx1 As String, fl As Object, strx2 As String, dicsize As Double, filez As Long, dics As String
    
    With Me.ListBox6
        k = .ListCount
        If k = 0 Then Exit Sub
        With ThisWorkbook.Application.FileDialog(msoFileDialogFolderPicker) '文件夹选择窗口(文件夹以此只能选一个)
            .Show
            If .SelectedItems.Count = 0 Then Exit Sub '未选择文件夹则退出sub
            strfolder = .SelectedItems(1)
        End With
        strfolder = strfolder & "\"
        k = k - 1
        Set fl = fso.CreateTextFile(strfolder & CStr(Format(Now, "yyyymmddhhmmss")) & ".txt", True, True) '创建txt文件
        dics = Left$(strfolder, 1)
        dicsize = fso.GetDrive(dics).AvailableSpace '磁盘的空间大小
        For i = 0 To k
            strx3 = .List(i, 0)
            SearchFile strx3    '搜索文件
            If Rng Is Nothing Then GoTo 101
            strx = Rng.Offset(0, 3) '文件路径
            strx4 = Rng.Offset(0, 1) '文件名
            strx1 = strfolder & strx4 '新的文件路径
            strx2 = strx4 '文件名
            If fso.fileexists(strx) = True And fso.fileexists(strx1) = False Then '文件存在且新文件夹下没有重复的文件
                filez = fso.GetFile(strx).Size                                    '这里可以细化-强制更新还是弹出提示窗口(overwrite or promotion)
                dicsize = dicsize - filez
                If dicsize < 209715200 Then MsgBox "磁盘空间不足!", vbCritical, "Warning": GoTo 100 '当磁盘的空间小于200M的时候
                strx = """" & strx & """"
                strfolder = """" & strfolder & """"  '------------'注意cmd命令的路径问题,当文件的路径存在空格
                Shell ("cmd /c" & "copy " & strx & Chr(32) & strfolder), vbHide     'fso.CopyFile (strx), strfolder
                fl.WriteLine strx3 & Space(2) & strx2 & Space(3) & "Success" & Chr(13)   '导出文件的同时,导出文件的列表
            Else
101
                fl.WriteLine strx3 & Space(2) & strx2 & Space(3) & "Fail" & Chr(13) 'space函数 https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/space-function
            End If
        Next
    End With
    Sleep 100
    Me.Label57.Caption = "操作成功"
100
    fl.Close
    Set fl = Nothing
End Sub

Private Sub CommandButton109_Click() '文件夹-清空列表
    Me.ListBox6.Clear
End Sub

Private Sub CommandButton110_Click() '文件添加到导出列表
    Dim i As Integer, k As Integer, p As Byte
    Dim strx As String, strx1 As String, strx2 As String
    
    With Me.ListView2
        k = .ListItems.Count
        If k = 0 Then Exit Sub
        For i = 1 To k
            If .ListItems(i).Selected = True Then
                strx = .SelectedItem.Text
                strx1 = .SelectedItem.SubItems(1)
                strx2 = .SelectedItem.SubItems(3)
                If CheckLb6(strx) = False Then
                    With Me.ListBox6
                        If .ListCount > 30 Then MsgBox "添加数量达到上限": Exit Sub
                        .AddItem
                        p = .ListCount - 1
                        .List(p, 0) = strx '编号
                        .List(p, 1) = strx1 '文件名
'                        .List(p, 2) = strx2 '文件路径
                    End With
                End If
            End If
        Next
    End With
End Sub

Private Sub CommandButton111_Click() '文件夹-移除列表
    Dim i As Integer, k As Integer
    
    With Me.ListBox6
        k = .ListCount
        If k = 0 Or .ListIndex < 1 Then Exit Sub '-1表示没选中
        k = k - 1
        For i = 0 To k
            If .Selected(i) = True Then .RemoveItem (i)
        Next
    End With
End Sub

Private Sub CommandButton112_Click() '文件夹-查看文件详情
    Dim i As Integer, k As Integer, strx As String
    
    With Me.ListView2
        k = .ListItems.Count
        If k = 0 Then Exit Sub
        For i = 1 To k
            '.SelectedItem.Text, 就算没有选中, 也会被执行, 默认为选中第一行, 注意checked和selected之间的区别
            If .ListItems(i).Selected = True Then
                strx = .ListItems(i).Text
                SearchFile strx
                If Rng Is Nothing Then Me.Label57.Caption = "文件丢失": Exit Sub
                ShowDetail (strx)
                Exit For
            End If
        Next
    End With
    Set Rng = Nothing
End Sub

Private Sub CommandButton113_Click() '功能区-最小
    If Workbooks.Count = 1 Then
        ThisWorkbook.Application.Visible = False
    Else
        ThisWorkbook.Windows(1).WindowState = xlMinimized
    End If
End Sub

Private Sub CommandButton114_Click() '工具设置
    With ThisWorkbook
        If .Application.Visible = False Then .Application.Visible = True
        .Windows(1).WindowState = xlMaximized
        With .Sheets("temp")
            .Activate
            .Range("aa1").Select
        End With
    End With
    Unload Me
End Sub

Private Sub CommandButton115_Click() '工具-重建
    Dim strx As String
    
    strx = ThisWorkbook.Path & "\lbrecord.xlsx"
    If fso.fileexists(strx) = True Then Me.Label57.Caption = "文件已存在": Exit Sub
    ThisWorkbook.Application.ScreenUpdating = False
    Call CreateWorksheet(strx)
    ThisWorkbook.Application.ScreenUpdating = True
    Me.Label57.Caption = "操作成功"
End Sub

Private Sub CommandButton116_Click() '工具-预查
    Dim strfolder As String
    Dim strx As String, strx1 As String
    Dim wsh As Object, fd As Folder
    Dim pid As Long, i As Integer
    
    With Application.FileDialog(msoFileDialogFolderPicker) '文件夹选择窗口(文件夹以此只能选一个)
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub '未选择文件夹则退出sub
        strfolder = .SelectedItems(1)
    End With
    Set fd = fso.GetFolder(strfolder)
    If fd.Files.Count = 0 And fd.SubFolders.Count = 0 Then Set fd = Nothing: Me.Label57.Caption = "空文件夹": Exit Sub
    strx = Split(strfolder, "\")(UBound(Split(strfolder, "\"))) '获取文件夹的名称
'    Set wsh = CreateObject("WScript.Shell")
    strx1 = Environ("UserProfile") & "\Desktop\" & strx & ".txt"
    
    strx1 = """" & strx1 & """"
    strfolder = """" & strfolder & """"
    '防止出现空格等干扰
'    wsh.Run "cmd /c tree " & strfolder & " /f >>" & strx1, 0 '调用cmd的tree命令 'https://wenku.baidu.com/view/66979c4fcf84b9d528ea7a96.html,cmd命令,0:hidden,3: max & activate
    
    pid = Shell("cmd /c tree " & strfolder & " /f >" & strx1, 0)
    
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0, pid)
'    Sleep 500
    Do
        GetExitCodeProcess hProcess, ExitCode
        DoEvents
        Sleep 50
        i = i + 1
    Loop While ExitCode = STILL_ALIVE And i < 75 '控制执行的时间, 等待时间不能超过75*50/1000 s
    If ExitCode = STILL_ALIVE Then Me.Label57.Caption = "正在生成文件树中,文件数量太多请稍后....": Exit Sub
    CloseHandle hProcess
    
    Shell "notepad.exe " & strx1, 3 '打开生成的文件
'    Set wsh = Nothing
    Set fd = Nothing
End Sub

Private Sub CommandButton117_Click() '工具-文件特定文件列表
    Dim strx As String, strx1 As String, strlen1 As Byte, strx2 As String, strx3 As String, strx4 As String, strx5 As String, strx6 As String
    Dim strx7 As String, strx8 As String
    Dim wsh As Object, i As Byte, xi As Variant, k As Byte, fd As Folder, fdz As Double, j As Single, wt As Integer, p As Integer, n As Single, pidx As Long

    With Me
        strx = Trim(folderoutput)
        strx1 = Trim(.TextBox27.Text)
        strlen1 = Len(strx1)
        If Len(strx) = 0 And strlen1 = 0 Then Exit Sub
        If InStr(strx, Chr(92)) = 0 Then Exit Sub          'chr(92),\,chr32,空格
        Set fd = fso.GetFolder(folderoutput)
'        fdz = fd.Size
        If fso.folderexists(folderoutput) = False Then Exit Sub
        If fd.Files.Count = 0 And fd.SubFolders.Count = 0 Then Exit Sub
        i = InStr(strx1, Chr(32))
        strx5 = fd.Name
        strx7 = Environ("UserProfile") & "\Desktop\" & strx5 & ".txt" '输出桌面
        If fso.fileexists(strx7) = True Then .Label57.Caption = "文件已存在": GoTo 100
        If strlen1 >= 6 And i = 0 Then .TextBox27.SetFocus: .Label57.Caption = "格式有误": Exit Sub '检查输入是否正确
        strx3 = "*."
        If strlen1 > 0 Then
            If i = 0 Then
                strx4 = strx3 & strx1
            Else
                xi = Split(strx1, Chr(32))
                i = UBound(xi)
                For k = 0 To i
                    strx4 = strx4 & strx3 & xi(k) & Chr(44) 'chr44,逗号
                Next
            End If
        Else
            strx4 = "*.*" '如果不输入扩展名,就默认任意扩展名
        End If
        strx4 = Chr(40) & strx4 & Chr(41) 'chr40/41括号
        strx6 = """" & strx & """"
        strx7 = """" & strx7 & """"
        strx8 = "for /r " & strx6 & " %a in " & strx4 & " do >>" & strx7 & " echo %~dpa%~nxa"  'for /r cmd命令
        Set wsh = CreateObject("WScript.Shell")
        wsh.Run ("cmd /c " & strx8), vbHide, True 'CreateObject("WScript.Shell"), 支持同步执行, true即表示同步执行, 但是需要考虑到执行的时间是否过长的问题
'        pidx = Shell("cmd /c " & strx8, vbHide)
        
'        If CheckProgramRun(pidx) = True Then .Label57.Caption = "目录生成中,稍后请手动打开": Exit Sub
    
'        If fdz < 1073741824 Then                    '等待文件完全生成的时间(不可取的手动测量的时间)
'            wt = 300
'        ElseIf fdz > 1073741824 And fdz < 10737418240# Then '1G
'            j = Int(fdz / 1073741824)
'            n = CSng(j)                '转换数据为单精度
'            p = (n + n / 10 - 0.1 - n / 50) * 300
'            wt = Int(Round(fdz / 1073741824 / j, 3) * p)
'        ElseIf fdz >= 10737418240# Then '10G
'            .Label57.Caption = "处理中..."
'            j = Int(fdz / 10737418240#)
'            n = CSng(j)
'            p = (n + n / 10 - 0.1 - n / 50) * 500 '根据不同设备的性能进行微调
'            wt = Int(Round(fdz / 10737418240# / j, 3) * p)
'        End If
'
'        Sleep wt '延时
100
        Shell "notepad.exe " & strx7, 3 '打开文件
        Set fd = Nothing
        Set wsh = Nothing
        folderoutput = ""
        .Label57.Caption = "处理完成"
    End With
End Sub

Private Sub CommandButton119_Click() '工具-CMD
    Shell ("cmd "), vbNormalFocus
End Sub

Private Sub CommandButton120_Click() '工具-随机密码
    Dim i As Byte, k As Byte

    With Me
        If .OptionButton7.Value = True Then
        i = 0
        ElseIf .OptionButton8.Value = True Then
        i = 1
        ElseIf .OptionButton9.Value = True Then
        i = 2
        End If
        If IsNumeric(.ComboBox13.Text) = True Then
            k = Abs(Int(.ComboBox13.Text))
        Else
            Exit Sub
        End If
        .TextBox28.Text = PasswordGR(i, k)
        If i = 0 Then .Label57.Caption = "纯数字密码安全性低"
    End With
End Sub

Private Sub CommandButton121_Click() '工具-关闭其他工作簿
    Dim i As Byte, k As Byte
    
    i = Workbooks.Count
    If i = 1 Then Exit Sub
    For k = 1 To i
        If Workbooks(k).Name <> ThisWorkbook.Name Then Workbooks(k).Close savechanges:=True
    Next
End Sub

Private Sub CommandButton122_Click() '备份文件
    If Len(Me.Label29.Caption) = 0 Then Exit Sub
    UserForm13.Show
End Sub

Private Sub CommandButton123_Click() '解压文件
    Dim strx As String, fd As Folder, fl As File, i As Byte, k As Byte, j As Byte, m As Byte, n As Byte
    Dim discz As Double, yesno As Variant
    Dim strx1 As String, strx2 As String, filez As Double, t As Long
    
    On Error GoTo 102
    With Me
        strx = Trim(.TextBox29.Text)
        m = Len(strx)
        If m = 0 Then Exit Sub
        strx1 = ThisWorkbook.Sheets("temp").Cells(42, "ab").Value
        If Len(strx1) = 0 Then MsgBox "尚未设置解压文件存放位置": Exit Sub
        If fso.folderexists(strx1) = False Then MsgBox "设置解压文件不存在": Exit Sub
        strx2 = Left$(strx1, 1)
        discz = fso.GetDrive(strx2).AvailableSpace '获取磁盘的大小
        If .CheckBox16.Value = True Then k = 1
        
        If InStr(strx, ".") > 0 Then '表示文件
            strx = filepathc
            If m <> Len(filepathc) Then
            If fso.fileexists(strx) = False Then .Label57.Caption = "文件不存在": Exit Sub
            If CheckFileFrom(strx, 1) = True Then MsgBox "文件来源受限": Exit Sub
            If TerminateEXE("bc.exe", 0) = 1 Then '检查
                yesno = MsgBox("是否关闭正在使用的Bandzip", vbYesNo, "Tips")
                If yesno = vbYes Then
                    TerminateEXE "bc.exe", 1
                    .Label57.Caption = "执行进程中,请勿使用Bandzip"
                Else
                    Exit Sub
                End If
            End If
            filez = fso.GetFile(strx).Size
            filez = filez * 5
            If filez > discz Then MsgBox "磁盘空间不足": Exit Sub
            ZipExtract strx, strx1
            If k = 1 Then
                t = timeGetTime
                Do
                    DoEvents
                    Sleep 200
                    i = TerminateEXE("bc.exe", 0)
                    If i = 0 Then Shell ("cmd /c" & "del /s " & strx), vbHide: .Label57.Caption = "操作完成": Exit Sub '当执行完成之后执行删除任务
                Loop Until i = 0 Or timeGetTime - t > 60000
                If i = 1 Then TerminateEXE "bc.exe", 1: .Label57.Caption = "操作出现异常": Exit Sub '如果无法在60秒内退出进程,就自动结束进程退出
            End If
            filepathc = ""
        Else                                '文件夹
            strx = folderpathc
            If fso.folderexists(strx) = False Then .Label57.Caption = "文件夹不存在": Exit Sub
            If CheckFileFrom(strx, 2) = True Then MsgBox "文件来源受限": Exit Sub
            ReDim arrcompress(1 To 100)
            ReDim arrfilez(1 To 100)
            flc = 0
            Set fd = fso.GetFolder(strx)
            If .CheckBox15.Value = False Then j = 1
            
            FoldersCompFile fd, j
            
            If flc = 0 Then Exit Sub
            .Label57.Caption = "任务执行中...请勿进行其他操作"
            For i = 1 To flc
                filez = arrfilez(i)
                filez = filez * 5
                discz = discz - filez
                If discz < 209715200 Then MsgBox "磁盘空间不足": GoTo 100
                ZipExtract arrcompress(i), strx1
            Next
            
            If k = 1 Then
            t = timeGetTime
            Do
                DoEvents
                Sleep 200
                i = TerminateEXE("bc.exe", 0)
                If i = 0 Then GoTo 101
            Loop Until i = 0 Or timeGetTime - t > 300000 '最长运行时间限制在5分钟
101
            If i = 1 Then TerminateEXE "bc.exe", 1: GoTo 102 '如果无法在60秒内退出进程,就自动结束进程退出
            For i = 1 To flc
                strx = arrcompress(i)
                strx = """" & strx & """" '防止存在空格等干扰因素
                Shell ("cmd /c" & "del /s " & strx), vbHide
            Next
            End If
100
            Erase arrcompress
            Erase arrfilez
            folderpathc = ""
        End If
        Sleep 100
        .Label57.Caption = "操作完成"
    End With
Exit Sub
102
Me.Label57.Caption = "操作出现异常"
Erase arrcompress
Erase arrfilez
Err.Clear
End Sub

Function FoldersCompFile(ByVal fd As Folder, Optional ByVal cmCode As Byte) '逐一获取压缩文件
    Dim sfd As Folder, fl As File, filex As String, strx As String
    
    For Each fl In fd.Files
        If flc > 100 Then Exit Function
        strx = fl.Name
        If InStr(strx, ".") = 0 Then GoTo 100
        filex = LCase(Right$(strx, Len(strx) - InStrRev(strx, "."))) '文件扩展名
        If filex Like "rar" Or filex Like "zip" Or filex Like "7z" Then '限定这三种压缩文件
            flc = flc + 1
            arrcompress(flc) = fl.Path
            arrfilez(flc) = fl.Size
        End If
100
    Next
    If fd.SubFolders.Count = 0 Or cmCode = 1 Then Exit Function
    For Each sfd In fd.SubFolders
        FoldersCompFile sfd, cmCode
    Next
End Function

Private Sub CommandButton124_Click() '转pdf
    Dim strx As String, i As Byte
    
    With Me
        strx = LCase(.Label24.Caption)
        If Len(strx) = 0 Then Exit Sub
        If strx Like "doc" Or strx Like "docx" Then
            If fso.fileexists(Filepathi) = True Then
                If Len(ThisWorkbook.Sheets("temp").Cells(43, "ab")) > 0 Then i = 1
                WordToPDF Filepathi, i
                .Label57.Caption = "操作成功"
            End If
        Else
            .Label57.Caption = "文件类型不匹配,限于Word"
        End If
    End With
End Sub

Private Sub CommandButton125_Click() '编辑-摘要-排版
    Dim strx As String, xi As Variant, i As Byte, k As Byte, strx1 As String
    
    With Me
        If .Label55.Visible = True Then Exit Sub
        strx = .TextBox2.Text
        If Len(strx) = 0 Then Exit Sub
        If InStr(strx, vbCrLf) > 0 Then         'vbCrLf等同于chr(10)换行符和chr(13)回车符
            xi = Split(strx, vbCrLf)
            i = UBound(xi)
            For k = 0 To i
                strx1 = strx1 & k + 1 & ". " & xi(k) & vbCrLf
            Next
            .TextBox2.Text = strx1
        Else
            .TextBox2.Text = "1. " & strx
        End If
    End With
End Sub

Private Sub CommandButton126_Click() '编辑-豆瓣链接-二维码
    Dim strx As String

    With Me
        QRtextEN = .Label106.Caption
        If Len(QRtextEN) = 0 Then Exit Sub
        UserForm18.Show
        Exit Sub
        '----------------下面部分备用
'        If Len(.Label106.Caption) = 0 Then Exit Sub
'        SearchFile .Label29.Caption
'        If rng Is Nothing Then .Label57.Caption = "文件不存在": Set rng = Nothing: Exit Sub
'        strx = rng.Offset(0, 33).Value
'        If Len(strx) = 0 Or fso.FileExists(strx) = False Then
'            If IsNetConnectOnline = False Then .Label57.Caption = "网络无法链接": Exit Sub
'            UserForm14.Show '显示在线图片
'        Else
'            QRfilepath = strx
'            UserForm16.Show '显示本地图片
'        End If
    End With
    Set Rng = Nothing
End Sub

Private Sub CommandButton127_Click() '工具-调试-清除IE缓存
    Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 1 "
    Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2 "
    Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8 "
    Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 16 "
    Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 32 "
    Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 255 "
    Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 4351 "
End Sub
'---------------------------------------------------浏览器
Private Sub CommandButton128_Click() '浏览器主页
    Dim strx As String
    
    strx = "https://www.baidu.com/"
    If browser1 = 1 Then
        With Me!web
            .Silent = True
            .Navigate (strx)
        End With
    ElseIf browser1 = 0 Then
        CreateWebBrowser (strx)
    End If
End Sub

Private Sub CommandButton143_Click() '外部链接打开
    Dim strx As String
    With Me!web
        strx = .LocationURL
    End With
    If Len(strx) = 0 Then Exit Sub
    If Len(ThisWorkbook.Sheets("temp").Cells(45, "ab").Value) > 0 Then
        Turlx = strx
        UserForm15.Show
        Exit Sub
    End If
    Webbrowser strx
End Sub

Private Sub CommandButton144_Click() '搜索豆瓣
    Dim strx As String, sengine As String, Urlx As String
    
    strx = Me.TextBox3.Text
    If Len(strx) = 0 Then Exit Sub
    sengine = "https://book.douban.com/subject_search?cat=1003&search_text="
    Urlx = sengine & Replace(strx, " ", "+")  'douban
    With Me!web
        .Silent = True
        .Navigate (Urlx)
    End With
End Sub

Private Sub CommandButton131_Click() '浏览器-金山
    Dim strx As String
    
    strx = "http://www.iciba.com/"
    If browser1 = 1 Then
        With Me!web
            .Silent = True
            .Navigate (strx)
        End With
    ElseIf browser1 = 0 Then
        CreateWebBrowser (strx)
    End If
End Sub

Private Sub CommandButton130_Click() '浏览器-豆瓣
    Dim strx As String
    
    strx = "https://book.douban.com/"
    If browser1 = 1 Then
        With Me!web
            .Silent = True
            .Navigate (strx)
        End With
    ElseIf browser1 = 0 Then
        CreateWebBrowser (strx)
    End If
End Sub

Function CreateWebBrowser(ByVal Urlx As String) '创建浏览器 '需要注意在multipage创建的webbrowser在页面切换后会消失,这里使用重新创建控件的方式复原而已 '页面上的窗体只是做个样子,实际上并不起作用
    On Error Resume Next
    Me.Controls.Add "shell.explorer.2", "Web", True
    If Len(Urlx) < 5 Then Urlx = "https://www.baidu.com"
    With Me!web
        .Top = 78
        .Left = 98
        .Height = 286
        .Width = 780
        .Navigate (Urlx)
        .Silent = True
    End With
    browser1 = 1
End Function

Private Sub CommandButton129_Click() '浏览器-获取豆瓣信息
    Dim strx As String, strx1 As String, yesno As Variant, strx2 As String, strx3 As String, strx4 As String, strx5 As String
    Dim arr() As String, strx6 As String, strx7 As String
    
    If Me.Label55.Visible = True Then Exit Sub
    strx = Me.Label29.Caption
    If Len(strx) = 0 Then Exit Sub
    If browser1 = 0 Then Exit Sub '如果浏览器控件尚未创建
    
    With Me!web
        strx1 = .LocationURL
    End With
    
    If InStr(strx1, "https://book.douban.com/subject/") = 0 Then Exit Sub '不在豆瓣的区域
    SearchFile strx1
    If Rng Is Nothing Then Me.Label57.Caption = "文件不存在": Exit Sub
    If Len(Rng.Offset(0, 25)) > 0 Then
        yesno = MsgBox("是否替换掉现有的链接", vbYesNo, "提示")
        If yesno = vbNo Then Set Rng = Nothing: Exit Sub
    End If
    Me!web.Stop
    With Me!web.Document
        strx2 = .getElementById("interest_sectl").InnerHtml '获得评分部分的源码
        strx3 = .getElementById("info").InnerHtml '作者,
        strx4 = .getElementById("mainpic").InnerHtml '书名+封面
    End With
    '------------------------------------------------https://www.w3school.com.cn/jsref/met_doc_getelementbyid.asp
    ReDim arr(1 To 5)
    arr = DoubanTreat(strx2, strx3, strx4)
    With Rng '将内容写入表格
        .Offset(0, 23) = arr(3) '名称
        .Offset(0, 24) = arr(1) '评分
        .Offset(0, 25) = strx           '链接
        .Offset(0, 14) = arr(2) '作者
        strx5 = arr(4) '封面名称'CheckRname
        xi = Split(strx5, "/")
        strx5 = xi(UBound(xi)) '文件名
        strx5 = Right$(strx5, Len(strx5) - InStrRev(strx5, ".") + 1)
        strx5 = strx1 & strx5
        strx6 = ThisWorkbook.Path & "\" & "bookcover"
        If fso.folderexists(strx6) = False Then fso.CreateFolder strx6
        strx5 = strx6 & "\" & strx5 '封面的存储路径
        strx7 = LCase(Right$(arr(4), 3))
        If strx7 = "jpg" Or strx7 = "png" Then '判断链接的内容是否满足要求
            If DownloadFilex(arr(4), strx5) = True Then .Offset(0, 34) = arr(4) '封面链接
            .Offset(0, 36) = strx5 '封面路径
        End If
        If Len(arr(5)) > 0 Then .Offset(0, 37) = arr(5) '作者国籍
    End With
    
    With Me
        .Label106.Caption = strx1
        .TextBox3.Text = arr(3)
        .TextBox4.Text = arr(2)
        .Label69.Caption = arr(1)
    End With
    
    Set Rng = Nothing
End Sub
'---------------------------------------------------浏览器

Private Sub CommandButton132_Click() '比较-文本比较
    Dim i As Byte

    With Me
    If .CommandButton132.Caption = "文本比较" Then
            With .TextBox30
                .Visible = True
                .Width = 254
                .Height = 230
                .Left = 10
                .Top = 12
            End With
            With .TextBox31
                .Visible = True
                .Width = 254
                .Height = 230
                .Left = 10
                .Top = 12
            End With
            For i = 204 To 228
                .Controls("label" & i).Visible = False
            Next
            For i = 98 To 101
                .Controls("commandbutton" & i).Visible = False
            Next
            .CommandButton104.Enabled = False
            .CommandButton132.Caption = "文件比较"
        Else
            With .TextBox30
                .Visible = False
            End With
            With .TextBox31
                .Visible = False
            End With
            For i = 204 To 228
                .Controls("label" & i).Visible = True
            Next
            .CommandButton104.Enabled = True
            .CommandButton132.Caption = "文本比较"
            For i = 98 To 101
                .Controls("commandbutton" & i).Visible = True
            Next
        End If
    End With
End Sub

Private Sub CommandButton133_Click() '单词-字典
    UserForm17.Show
End Sub

Private Sub CommandButton134_Click() '编辑-摘要-二维码
    QRtextCN = Me.TextBox2.Text
    If Len(QRtextCN) = 0 Then Exit Sub
    UserForm1.Show
End Sub

Private Sub CommandButton135_Click() '编辑-搜索-条形码
    Barcodex = Me.Label29.Caption
    If Len(Barcodex) = 0 Or Me.Label55.Visible = True Then Exit Sub
    UserForm19.Show
End Sub

Private Sub CommandButton136_Click() '编辑-文件名-二维码
    Dim strx As String
    
    If Me.Label55.Visible = True Then Exit Sub
    If Len(Filenamei) = 0 Then Exit Sub
    strx = Left$(Filenamei, InStrRev(Filenamei, ".") - 1) '不包含文件扩展名的文件名
    If Filenamei Like "*[一-龥]*" Then
        QRtextCN = strx
        UserForm1.Show
    Else
        QRtextEN = strx
        UserForm18.Show
    End If
End Sub

Private Sub CommandButton138_Click() '编辑添加到ftp
    If Len(Me.Label29.Caption) = 0 Then Exit Sub
    If Me.Label55.Visible = True Then Exit Sub
    UserForm20.Show
End Sub

Private Sub CommandButton139_Click() '编辑-文件销毁
    Dim i As Byte, strx As String, k As Byte
    With Me
        strx = .Label29.Caption
        If Len(strx) = 0 Then Exit Sub
        If .Label55.Visible = True Then Exit Sub
        If .Label236.Caption = "Y" Then i = 1
        If Len(ThisWorkbook.Sheets("temp").Cells(50, "ab").Value) > 0 Then k = 2 '深度清理
        If FileDestroy(Filepathi, k, i) = True Then
            SearchFile strx
            If Not Rng Is Nothing Then ThisWorkbook.Sheets("书库").rows(Rng.Row).Delete Shift:=xlShiftUp '删除表格 '被销毁的文件将不会被记录进备份中
            DeleFileOverx strx
            .Label57.Caption = "操作成功"
            Set Rng = Nothing
        End If
    End With
End Sub

Private Sub CommandButton140_Click() '工具-创建文件夹分类
    Dim strfolder As String
    Dim drx As Drive, i As Byte, k As Byte, yesno As Variant
    
    With ThisWorkbook.Application.FileDialog(msoFileDialogFolderPicker) '文件夹选择窗口(文件夹以此只能选一个)
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub '未选择文件夹则退出sub
        strfolder = .SelectedItems(1)
    End With
    For Each drx In fso.Drives
        If drx.DriveType = 2 Then '必须注意在多条件判断时,条件的执行后否造成错误,如,某值的判断必须不是空才能判断,如果是空就会出错,显然需要先确定这个值是否为空才进行下一步的判断
            i = i + 1
        End If
    Next
    If i > 1 Then '有两个以上的固定硬盘
        If UCase(Left(strfolder, 2)) = Environ("SYSTEMDRIVE") Then MsgBox "禁止选择系统盘", vbOKOnly, "Tips": Exit Sub
        If UBound(Split(strfolder, "\")) > 2 Then MsgBox "建议选择根目录或1级目录", vbOKOnly, "Tips": Exit Sub
    Else
        If CheckFileFrom(strfolder, 2) = True Then MsgBox "所选位置受限", vbOKOnly, "Tips": Exit Sub
    End If
    yesno = MsgBox("默认中文,选 ""否""将选择英文", vbYesNo, "Tips")
    If yesno = vbYes Then
        k = 0
    Else
        k = 1
    End If
    If CreateFolder(strfolder, k) = True Then
        Me.Label57.Caption = "创建成功"
    Else
        Me.Label57.Caption = "创建失败"
    End If
End Sub

Private Sub CommandButton141_Click() '非ansi检查
    Dim fdx As FileDialog, strfolder As String
    Dim selectfile As Variant

    If Me.CheckBox20.Value = True Then
    With ThisWorkbook.Application.FileDialog(msoFileDialogFolderPicker) '文件夹选择窗口(文件夹以此只能选一个)
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub '未选择文件夹则退出sub
        strfolder = .SelectedItems(1)
        If ErrCode(strfolder, 0) > 1 Then
            Me.Label238.Caption = errcodenx
        Else
            Me.Label238.Caption = "Clear"
        End If
        Exit Sub
    End With
    End If
    Set fdx = ThisWorkbook.Application.FileDialog(msoFileDialogFilePicker)
    With fdx
        .AllowMultiSelect = False '不允许选择多个文件(注意不是文件夹,文件夹只能选一个)
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub
        selectfile = .SelectedItems(1)
        If ErrCode(selectfile, 0) > 1 Then
            Me.Label238.Caption = errcodenx
        Else
            Me.Label238.Caption = "Clear"
        End If
    End With
    Set fdx = Nothing
End Sub

Private Sub CommandButton145_Click() 'windows管理工具
    Shell "cmd /c" & "C:\Windows\System32\control.exe /name Microsoft.AdministrativeTools", vbHide
End Sub

Private Sub CommandButton146_Click() '注册表
    Shell "cmd /c" & "regedit", vbHide
End Sub

Private Sub CommandButton147_Click() '浏览器
    UserForm15.Show
End Sub

Private Sub CommandButton148_Click() '文件复制到剪切板
    If Len(Filepathi) = 0 Then Exit Sub
    If fso.fileexists(Filepathi) = True Then
        CutOrCopyFiles Filepathi
        Me.Label57.Caption = "文件已复制到剪切板"
    Else
        Me.Label57.Caption = "文件不存在"
    End If
End Sub

Private Sub CommandButton149_Click() '清除数据
    UserForm2.Show
End Sub

Private Sub CommandButton150_Click() '生成测试文件
    TestFileGR
End Sub

Private Sub CommandButton152_Click() '编辑查看更多文件详情
    Dim strx As String
    
    With Me
        strx = .Label24.Caption
        If Len(strx) = 0 Then Exit Sub
        If strx Like "doc*" Or strx = "xlsx" Or strx Like "ppt*" Then
            FileDetail.Show
        Else
            .Label57.Caption = "不受支持格式"
        End If
    End With
End Sub

Private Sub Image1_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '双击打开大图片
    UserForm24.Show
End Sub

Private Sub Label23_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '文件名双击复制
    If Me.Label55.Visible = True Then Exit Sub
    If Me.Label74.Caption = "Y" Then SetClipboard Filenamei: Exit Sub
    CopyToClipboard Filenamei, "文件名已复制"  '注意这里不能直接复制标签上的内容, 假如文件名包含非ansi字符, 标签上的内容就是有问题的
End Sub

Private Function CopyToClipboard(ByVal strText As String, Optional ByVal strtips As String) '复制到粘贴板
    Dim textb As Object, strx As String
    
    With Me
        If Len(strText) = 0 Then Exit Function
        Set textb = .Controls.Add("Forms.TextBox.1", "Text1", False) '以创建临时textbox的方式实现复制内容;需要注意的是这种方法虽然可以避免很多复制的格式问题
        '-------------------------------------------------------------但是也有一个基本问题无法处理,非ansi的问题...又忘了这个老问题了
        With textb
            .Text = strText
            .SelStart = 0
            .SelLength = Len(.Text)
            .Copy
        End With
       If Len(strtips) > 0 Then .Label57.Caption = strtips
    End With
    Set textb = Nothing
End Function

Private Sub Label25_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.Label55.Visible = True Then Exit Sub
    If Me.Label74.Caption = "Y" Then SetClipboard Filepathi: Exit Sub
    CopyToClipboard Me.Label25.Caption, "文件路径已复制"
End Sub

Private Sub Label26_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.Label55.Visible = True Then Exit Sub
    OpenFileLocation Folderpathi
End Sub

Private Sub Label29_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.Label55.Visible = True Then Exit Sub
    CopyToClipboard Me.Label29.Caption, "编号已复制"
End Sub

Private Sub Label71_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.Label55.Visible = True Then Exit Sub
    CopyToClipboard Me.Label71.Caption, "Hash已经复制"
End Sub

Private Sub OptionButton10_Click() '字符串-加密-校检
    If Me.OptionButton10.Value = True Then StringCH (1)
End Sub

Sub StringCH(ByVal cmCode As Byte) ''字符串-处理
    Dim i As Byte
    
    Select Case cmCode
        Case 1: i = 1
        Case 2: i = 2
        Case 3: i = 3
        Case Else: Exit Sub
    End Select
    ThisWorkbook.Sheets("temp").Cells(38, "ab") = i
End Sub

Private Sub OptionButton11_Click() '字符串-处理
    If Me.OptionButton11.Value = True Then StringCH (2)
End Sub

Private Sub OptionButton12_Click() '字符串-处理
    If Me.OptionButton12.Value = True Then StringCH (3)
End Sub

Private Sub TextBox14_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single) '为textbox添加右键菜单
    On Error Resume Next
    If Button = 2 And Not NewM Then
        On Error Resume Next
        With ThisWorkbook.Application
            .CommandBars("NewMenu").Delete
            .CommandBars.Add "NewMenu", msoBarPopup, False, True
            With .CommandBars("NewMenu")
                .Controls.Add msoControlButton
                .Controls(1).Caption = "剪切"
                .Controls(1).FaceId = 21
                .Controls(1).OnAction = "Cutx"
                .Controls.Add msoControlButton
                .Controls(2).Caption = "复制"
                .Controls(2).FaceId = 19
                .Controls(2).OnAction = "Copyx"
                .Controls.Add msoControlButton
                .Controls(3).Caption = "粘贴"
                .Controls(3).FaceId = 22
                .Controls(3).OnAction = "Pastex"
                .ShowPopup
            End With
        End With
    End If
    NewM = Not NewM
End Sub

Private Sub TextBox14_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With Me.TextBox14 '可以利用textbox的这个属性作为实现复制信息到剪切板的间接途径
        If Len(.Text) = 0 Then Exit Sub
        .Text = .Value
        .SelStart = 0
        .SelLength = Len(.Text)
        .Copy
    End With
End Sub

Private Sub TextBox15_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With Me.TextBox15 '可以利用textbox的这个属性作为实现复制信息到剪切板的间接途径
        If Len(.Text) = 0 Then Exit Sub
        .Text = .Value
        .SelStart = 0
        .SelLength = Len(.Text)
        .Copy
    End With
End Sub

Private Sub TextBox16_Change()
    With Me.TextBox16 '可以利用textbox的这个属性作为实现复制信息到剪切板的间接途径
        If Len(.Text) = 0 Then Exit Sub
        .Text = .Value
        .SelStart = 0
        .SelLength = Len(.Text)
        .Copy
    End With
End Sub

Private Sub TextBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '搜索/编辑-文件摘要
    With Me.TextBox2 '可以利用textbox的这个属性作为实现复制信息到剪切板的间接途径
        If Len(.Text) = 0 Then Exit Sub
        .Text = .Value
        .SelStart = 0
        .SelLength = Len(.Text)
        .Copy
    End With
End Sub

Private Sub TextBox20_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '工具-md5计算
    Dim fdx As FileDialog
    Dim selectfile As Variant
    
    Set fdx = ThisWorkbook.Application.FileDialog(msoFileDialogFilePicker)
    With fdx
        .AllowMultiSelect = False '允许选择多个文件(注意不是文件夹,文件夹只能选一个)
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub
        filepath5 = .SelectedItems(1)
        If ErrCode(filepath5, 1) > 1 Then MsgShow "文件路径包含非ansi编码,请勿手动修改内容框的信息", "Tips", 1800
        Me.TextBox20.Text = filepath5
        If CheckFileFrom(filepath5, 1) = True Then Me.Label57.Caption = "文件来源受限": Exit Sub
        Me.TextBox21.SetFocus
    End With
    Set fdx = Nothing
End Sub

Private Sub TextBox26_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '工具-输出特定文件
    Dim strfolder As String
    
    With Application.FileDialog(msoFileDialogFolderPicker) '文件夹选择窗口(文件夹以此只能选一个)
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub '未选择文件夹则退出sub
        strfolder = .SelectedItems(1)
    End With
    folderoutput = strfolder
    If ErrCode(folderoutput, 1) > 1 Then MsgShow "文件路径包含非ansi编码,请勿手动修改内容框的信息", "Tips", 1800
    Me.TextBox26.Text = folderoutput
    Me.TextBox27.SetFocus
End Sub

Private Sub CommandButton67_Click() '筛选-清空筛选条件
    With Me
        .ComboBox1.Value = ""
        .ComboBox7.Value = ""
        .ComboBox8.Value = ""
    End With
End Sub

Private Sub ComboBox1_Click() '筛选
    With Me
        If .CheckBox3.Value = False Then
            .ComboBox7.Value = ""
            .ComboBox8.Value = ""
        End If
    End With
End Sub

Private Sub ComboBox7_Click() '筛选
    With Me
        If .CheckBox3.Value = False Then
            .ComboBox1.Value = ""
            .ComboBox8.Value = ""
        End If
    End With
End Sub

Private Sub ComboBox8_click() '筛选
    With Me
        If .CheckBox3.Value = False Then
            .ComboBox7.Value = ""
            .ComboBox1.Value = ""
        End If
    End With
End Sub

Private Sub CommandButton91_Click() '工具-文件处理
    Dim i As Byte
    If filepath5 = ThisWorkbook.fullname Then MsgBox "非法操作", vbCritical, "Warning": Exit Sub
    With Me
    If Len(Trim(.TextBox20.Text)) = 0 Then Exit Sub
    If .OptionButton6.Value = True Then
    i = 2
    ElseIf .OptionButton5.Value = True Then
    i = 1
    ElseIf .OptionButton5.Value = False And .OptionButton6.Value = False Then
    i = 0
    End If
    FileDestroy filepath5, i
    filepath5 = "" '使用完重置参数
    End With
End Sub

Private Function FileDestroy(ByVal strx As String, ByVal cmCode As Byte, Optional ByVal cmfrom As Byte) As Boolean '文件销毁
    Dim fl As File
    Dim flop As Object
    Dim i As Long, k As Byte, j As Double, yesno As Variant, errx As Integer, strx1 As String, p As Integer, m As Double, n As Byte, c As Integer
    
    On Error GoTo 100
    FileDestroy = False
    With Me
        strx = Trim(strx)
        If Len(strx) = 0 Then Exit Function
        yesno = MsgBox("文件将被彻底销毁,确定?_", vbYesNo, "Warning!!!")
        If yesno = vbNo Then Exit Function
        If fso.fileexists(strx) = False Then .Label57.Caption = "文件不存在": Exit Function
        If CheckFileFrom(strx, 1) = True Then .Label57.Caption = "该文件的来源受限": Exit Function '限制添加来自系统盘的文件
        
        Set fl = fso.GetFile(strx)
        j = fl.Size
        If j > 536870912 And cmCode = 2 Then
            .Label57.Caption = "文件过大,深度处理限制512M"
            Set fl = Nothing
            Exit Function
        End If
        'Excel类文件 '无法通过检查命令行来查看文件是否处于打开的状态
        If InStr(strx, ".") > 0 Then '防止没有扩展名的文件
            strx1 = LCase(Right$(strx, Len(strx) - InStrRev(strx, "."))) '文件扩展名
            If strx1 Like "xl*" Then
                c = Workbooks.Count
                strx1 = fl.Name
                For n = 1 To c
                    If strx1 = Workbooks(n).Name Then .Label57.Caption = "文件处于打开的状态": Set fl = Nothing: Exit Function
                Next
                c = -1
            End If
        End If
        If j >= 1048576 And cmCode = 2 Then
            If j < 10485760 Then
                p = 32
            ElseIf j >= 10485760 And j < 52428800 Then
                p = 128
            ElseIf j >= 52428800 And j < 104857600 Then '避免持续写入造成类似卡死的问题
                p = 512
            Else
                p = 1024
            End If
            j = j + 1048576               '深度处理-写入的数据完全覆盖之前的数据-写入的数据将比源原文件还大1024*1024
            .Label57.Caption = "处理中..."
        ElseIf cmCode = 1 Then
            j = 10240: p = 1
        ElseIf cmCode = 0 Then
            j = 1024: p = 1
        ElseIf cmCode = 2 Then
            j = 1048576: p = 16
        End If
        If cmfrom = 1 Then GoTo 101
        Set flop = fl.OpenAsTextStream(ForWriting, TristateMixed) '抹掉文件的信息
        m = Round((j / p), 0) 'round函数,0,不保留小数
        With flop
            For i = 1 To m
                k = RandNumx(1) '随机0/1
                strx1 = String(p, k) '生成p个相同的字符 'https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/string-function
                .Write strx1
            Next
            .Close
        End With
        .Label57.Caption = "处理完成"
    End With
    fso.DeleteFile (strx) '不经回收站删除文件
    Set fl = Nothing
    Set flop = Nothing
    FileDestroy = True
    Exit Function
100
    errx = Err.Number
    If errx = 70 Then
        If c <> -1 Then
101
            If WmiCheckFileOpen(strx) = False Then '如果是受密码保护的文件,常规方法无法向文件写入内容,调用powershell强制写入
                PowerSHForceW strx, j
                Err.Clear
                Set fl = Nothing
                fso.DeleteFile (strx)
                FileDestroy = True
                Me.Label57.Caption = "处理完成"
                Set flop = Nothing: Exit Function
            End If
            Me.Label57.Caption = "文件处于打开的状态"
        Else
            Me.Label57.Caption = "文件具有密码保护"
        End If
    Else
        Me.Label57.Caption = "异常,处理文件失败"
    End If
    Err.Clear
    Set fl = Nothing
    Set flop = Nothing
End Function

Private Sub CommandButton93_Click() '关闭设置
    Dim i As Byte, k As Byte
    
    With Me.MultiPage1
        k = .Pages.Count
        k = k - 1
        .Pages(k).Visible = False
        k = k - 1
        For i = 0 To k
        .Pages(i).Visible = True
        Next
        .Value = 0 '返回首页
    End With
End Sub

Private Sub CommandButton75_Click() '设置-about me
    UserForm7.Show 1
End Sub

Private Sub CommandButton94_Click() '编辑-搜索-刷新'重新获取文档的信息 '将信息写入列表和显示窗口
    Dim strx As String
    
    strx = Me.Label29.Caption
    If Len(strx) = 0 Or Me.Label55.Visible = True Then Exit Sub
    Call UpdateFileIn(strx)
End Sub

Function UpdateFileIn(ByVal filecode As String) '更新文件信息-和打开文件可以考虑合并
    Dim fl As File
    Dim fz As Long, filex As String
    With Me
    If FileStatus(filecode, 2) = 4 Then '文件存在于目录且文件存在于磁盘
        Set fl = fso.GetFile(Rng.Offset(0, 3).Value)
        If fl.DateLastModified <> Rng.Offset(0, 6).Value Then '文件的修改时间发生变化-通过检测文件的修改时间来判断文件的属性是否发生变化
            Rng.Offset(0, 6) = fl.DateLastModified '文件修改信息
            fz = fl.Size
            filex = UCase(Rng.Offset(0, 2).Value) '文件扩展名
            If fz < 1048576 Then
                Rng.Offset(0, 7) = Format(fz / 1024, "0.00") & "KB" '文件大小
            Else
                Rng.Offset(0, 7) = Format(fz / 1048576, "0.00") & "MB"    '更新的信息先写入表格中
            End If
            Rng.Offset(0, 5) = fz '文件初始大小
            If filex Like "EPUB" Or filex Like "MOBI" Or filex Like "PDF" Then Rng.Offset(0, 9) = GetFileHashMD5(fl.Path) '文件的hash值,当文件被修改后文件的hash值就发生改变
            Call FileChange '更新数据
        Else
            .Label57.Caption = "信息已是最新"
        End If
    Else
        .Label57.Caption = "文件已删除"
    End If
    .Label57.Caption = "更新完毕"
    End With
    Set Rng = Nothing
    Set fl = Nothing
End Function

Private Sub CommandButton102_Click() '添加比较a
    Dim str As String
    
    With Me
        str = .Label29.Caption
        If Len(str) = 0 Or .Label55.Visible = True Then Exit Sub
        If Len(.Label179.Caption) Or Len(.Label153.Caption) > 0 Then
             If str = .Label153.Caption Or str = .Label179.Caption Then .Label57.Caption = "已添加": Exit Sub
        End If
        Call FileCompA(str)
    End With
End Sub

Private Sub CommandButton95_Click() '文件夹-比较a
    Dim i As Integer, k As Integer, p As Byte
    
    With Me.ListView2
        k = .ListItems.Count
        If k = 0 Then Exit Sub
        For i = 0 To k
        If .ListItems(i).Checked = True Then p = 1: Set Rng = Nothing: Exit For
        Next
        If p = 0 Then Exit Sub
        Call FileCompA(.SelectedItem.Text)
    End With
End Sub

Function FileCompA(ByVal strx As String) '比较a
    With Me
        If strx = .Label179.Caption Then Exit Function
        SearchFile (strx)
        If Rng Is Nothing Then .Label57.Caption = "文件目录丢失": Exit Function
        .Label179.Caption = Rng.Offset(0, 0)
        .Label180.Caption = Rng.Offset(0, 1)
        .Label181.Caption = Rng.Offset(0, 5)
        .Label182.Caption = Rng.Offset(0, 8)
        .Label183.Caption = Rng.Offset(0, 6)
        .Label184.Caption = Rng.Offset(0, 12)
        .Label185.Caption = Rng.Offset(0, 14)
        .Label186.Caption = Rng.Offset(0, 17)
        .Label187.Caption = Rng.Offset(0, 18)
        .Label188.Caption = Rng.Offset(0, 19)
        .Label189.Caption = Rng.Offset(0, 20)
        .Label190.Caption = Rng.Offset(0, 15)
        .Label191.Caption = Rng.Offset(0, 16)
    End With
    Set Rng = Nothing
End Function

Private Sub CommandButton103_Click() '添加比较b
    Dim strx As String
    
    With Me
        strx = .Label29.Caption
        If Len(strx) = 0 Or .Label55.Visible = True Then Exit Sub
        If Len(.Label153.Caption) Or Len(.Label179.Caption) > 0 Then
            If strx = .Label153.Caption Or strx = .Label179.Caption Then .Label57.Caption = "已添加": Exit Sub
        End If
        Call FileCompB(strx)
    End With
End Sub

Private Sub CommandButton96_Click() '文件夹-添加比较b
    Dim i As Integer, k As Integer, p As Byte
    
    With Me.ListView2
        k = .ListItems.Count
        If k = 0 Then Exit Sub
        For i = 0 To k
        If .ListItems(i).Checked = True Then p = 1: Exit For
        Next
        If p = 0 Then Exit Sub
        Call FileCompB(.SelectedItem.Text)
    End With
End Sub

Function FileCompB(ByVal strx As String)
    With Me
        If strx = .Label153.Caption Then Exit Function
        SearchFile (strx)
        If Rng Is Nothing Then .Label57.Caption = "文件目录丢失": Set Rng = Nothing: Exit Function
        .Label153.Caption = Rng.Offset(0, 0)
        .Label154.Caption = Rng.Offset(0, 1)
        .Label155.Caption = Rng.Offset(0, 5)
        .Label156.Caption = Rng.Offset(0, 8)
        .Label157.Caption = Rng.Offset(0, 6)
        .Label158.Caption = Rng.Offset(0, 12)
        .Label159.Caption = Rng.Offset(0, 14)
        .Label160.Caption = Rng.Offset(0, 17)
        .Label161.Caption = Rng.Offset(0, 18)
        .Label162.Caption = Rng.Offset(0, 19)
        .Label163.Caption = Rng.Offset(0, 20)
        .Label164.Caption = Rng.Offset(0, 15)
        .Label165.Caption = Rng.Offset(0, 16)
    End With
    Set Rng = Nothing
End Function

Private Sub CommandButton104_Click() '两个文件进行md5比较
    Dim strx As String, strx1 As String, strx3 As String, strx4 As String
    Dim strx2 As String, strx5 As String, strx6 As String, strx7 As String
    
    With Me
        strx6 = .Label179.Caption
        strx7 = .Label153.Caption
        If Len(strx6) = 0 Or Len(strx7) = 0 Then Exit Sub
        If FileStatus(strx6, 2) = 4 Then
            strx = Rng.Offset(0, 3)
            strx2 = UCase(Rng.Offset(0, 2))
            Set Rng = Nothing
            If strx2 Like "EPUB" Or strx2 Like "MOBI" Or strx2 Like "PDF" Then .Label57.Caption = "此类文件不支持此功能": Set Rng = Nothing: Exit Sub
            If FileStatus(strx7, 2) = 4 Then
                strx1 = Rng.Offset(0, 3)
                strx5 = UCase(Rng.Offset(0, 2))
                If strx5 Like "EPUB" Or strx5 Like "MOBI" Or strx5 Like "PDF" Then .Label57.Caption = "此类文件不支持此功能": Set Rng = Nothing: Exit Sub
            Else
                Set Rng = Nothing
                .Label57.Caption = "文件丢失"
                Exit Sub
            End If
        Else
            Set Rng = Nothing
            .Label57.Caption = "文件丢失"
            Exit Sub
        End If
        If strx2 = strx5 Then
            strx3 = GetFileHashMD5(strx)
            strx4 = GetFileHashMD5(strx1)
            If Len(strx3) > 2 And Len(strx4) > 2 Then
                If strx3 = strx4 Then
                    .Label216.Caption = "Match"
                Else
                    .Label216.Caption = "MisMacth"
                End If
                .Label214.Caption = strx3
                .Label215.Caption = strx4
            Else
                .Label57.Caption = "未获取到有效值"
            End If
        Else
            .Label57.Caption = "类型不匹配"
        End If
    End With
    Set Rng = Nothing
End Sub

Private Sub CommandButton97_Click() '比较-比较
    Dim i As Byte, k As Byte, strLen As Byte, strlen1 As Byte
    Dim fz As Long, fz1 As Long
    Dim date1 As Date, date2 As Date, date3 As Date, date4 As Date
    Dim strx As String, strx1 As String, strx2 As String
    
    With Me
        If .CommandButton132.Caption = "文本比较" Then
            If Len(.Label179.Caption) > 0 And Len(.Label153.Caption) > 0 Then
                strx = .Label180.Caption
                strx1 = .Label154.Caption
                strx = Left(strx, Len(strx) - Len(Split(strx, Chr(46))(UBound(Split(strx, Chr(46))))) - 1)   '获取不包含扩展名的文件名
                strx1 = Left(strx, Len(strx1) - Len(Split(strx1, Chr(46))(UBound(Split(strx1, Chr(46))))) - 1)
                strLen = Len(strx)
                strlen1 = Len(strx1)
                If strLen >= strlen1 Then
                    For i = 1 To strLen
                        If InStr(strx, Mid(strx1, i, 1)) > 0 Then k = k + 1
                    Next
                Else
                    For i = 1 To strlen1
                        If InStr(strx1, Mid(strx, i, 1)) > 0 Then k = k + 1
                    Next
                End If
                .Label204.Caption = Format(k / i - 1, "0.0000") '文件相似度
                fz = CLng(.Label181.Caption)
                fz1 = CLng(.Label155.Caption)
                .Label205.Caption = Format((fz - fz1) / fz, "0.0000") '文件大小偏差
                date1 = .Label182.Caption
                date2 = .Label156.Caption
                date3 = .Label183.Caption
                date4 = .Label157.Caption
                .Label206.Caption = DateDiff("s", date1, date2) & "s" '创建时间偏差
                .Label207.Caption = DateDiff("s", date3, date4) & "s" '修改时间偏差
                If Len(.Label184.Caption) > 0 And Len(.Label158.Caption) > 0 Then '作者
                    If .Label184.Caption = .Label160.Caption Then
                        .Label208.Caption = "Y"
                    Else
                        .Label208.Caption = "N"
                    End If
                End If
                If Len(.Label185.Caption) > 0 And Len(.Label159.Caption) > 0 Then .Label209.Caption = .Label185.Caption & " / " & .Label161.Caption '打开次数
                If Len(.Label186.Caption) > 0 And Len(.Label160.Caption) > 0 Then .Label210.Caption = .Label186.Caption & " / " & .Label162.Caption '内容评分
                If Len(.Label187.Caption) > 0 And Len(.Label161.Caption) > 0 Then .Label211.Caption = .Label187.Caption & " / " & .Label163.Caption '推荐指数
                If Len(.Label190.Caption) > 0 And Len(.Label164.Caption) > 0 Then .Label212.Caption = .Label188.Caption & " / " & .Label164.Caption 'pdf
                If Len(.Label191.Caption) > 0 And Len(.Label165.Caption) > 0 Then .Label213.Caption = .Label188.Caption & " / " & .Label164.Caption '文本
            End If
        Else
            strx1 = .TextBox30.Text
            strx2 = .TextBox31.Text
            If Len(strx1) = 0 Or Len(strx2) = 0 Then Exit Sub
            .Label216.Visible = True
            If CRC32API(strx1) = CRC32API(strx2) Then
                .Label216.Caption = "Mactch"
            Else
                .Label216.Caption = "MisMactch"
            End If
        End If
    End With
End Sub

Private Sub CommandButton98_Click() '比较a-删除
    Dim p As Byte, i As Byte, strx As String
    
    With Me
        strx = .Label179.Caption
        If Len(strx) = 0 Then Exit Sub
        i = FileStatus(strx)
        If i >= 5 Then .Label57.Caption = "文件处于打开的状态": Set Rng = Nothing: Exit Sub
        If i = 0 Then
            With Rng
                If Len(.Offset(0, 26).Value) > 0 Then p = 1
                Call FileDeleExc(.Offset(0, 3).Value, .Offset(0, 2).Value, .Row, p, 1, 1)
            End With
        End If
        Call ClearComp(1)
    End With
    Set Rng = Nothing
End Sub

Private Sub CommandButton101_Click() '比较b-删除
 Dim p As Byte, i As Byte, strx As String
    
    With Me
        strx = .Label153.Caption
        If Len(strx) = 0 Then Exit Sub
        i = FileStatus(strx)
        If i >= 5 Then .Label57.Caption = "文件处于打开的状态": Set Rng = Nothing: Exit Sub
        If i = 0 Then
            With Rng
                If Len(.Offset(0, 26).Value) > 0 Then p = 1
                Call FileDeleExc(.Offset(0, 3).Value, .Offset(0, 2).Value, .Row, p, 1, 1)
            End With
        End If
        Call ClearComp(2)
    End With
    Set Rng = Nothing
End Sub

Private Sub CommandButton105_Click() '比较-清除
    With Me
        If .CommandButton132.Caption = "文本比较" Then
            If Len(.Label179.Caption) = 0 And Len(.Label153.Caption) = 0 Then Exit Sub
            Call ClearComp(3)
        Else
            .TextBox30.Text = ""
            .TextBox31.Text = ""
            .Label216.Caption = ""
        End If
    End With
End Sub
'------------------------------------------------------比较

Private Sub CommandButton99_Click() '比较-打开
    Dim i As Byte, strx As String, strx1 As String, strx2 As String
    
    With Me
        strx = .Label179.Caption
        strx1 = .Label180.Caption
        strx2 = LCase(Right$(strx1, Len(strx1) - InStrRev(strx1, ".")))
        If strx2 Like "xl*" Then
            If strx2 <> "xls" Or "xlsx" Then MsgBox "禁止打开此类文件", vbCritical, "Warning": Exit Sub
        End If
        i = FileStatus(strx)
        If i >= 5 Then .Label57.Caption = "文件处于打开的状态": Set Rng = Nothing: Exit Sub
        If i = 0 Then
        With Rng
            Call OpenFile(strx, .Offset(0, 1).Value, .Offset(0, 2).Value, .Offset(0, 3).Value, 1, .Offset(0, 26).Value, 1)
        End With
        End If
    End With
    Set Rng = Nothing
End Sub

Private Sub CommandButton100_Click() '比较-打开
    Dim i As Byte, strx As String, strx1 As String, strx2 As String
    
    With Me
        strx = .Label153.Caption
        strx1 = .Label154.Caption
        strx2 = LCase(Right$(strx1, Len(strx1) - InStrRev(strx1, ".")))
        If strx2 Like "xl*" Then
            If strx2 <> "xls" Or "xlsx" Then MsgBox "禁止打开此类文件", vbCritical, "Warning": Exit Sub
        End If
        i = FileStatus(strx)
        If i >= 5 Then .Label57.Caption = "文件处于打开的状态": Set Rng = Nothing: Exit Sub
        If i = 0 Then
        With Rng
            Call OpenFile(strx, .Offset(0, 1).Value, .Offset(0, 2).Value, .Offset(0, 3).Value, 1, .Offset(0, 26).Value, 1)
        End With
        End If
    End With
    Set Rng = Nothing
End Sub

Private Sub TextBox24_Change()
    With Me
        If Len(.TextBox24.Text) = 0 Then .Label109.Caption = ""
    End With
End Sub

Private Sub TextBox24_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer) '工具-质数检查-按下回车键,执行检查命令
    If KeyCode = 13 Then Call ChPN
End Sub

Private Sub CommandButton92_Click() '工具-质数检查
    Call ChPN
End Sub

Sub ChPN() '工具-检查质数
    Dim strx As String, strLen As Byte
    
    With Me
        strx = Trim(.TextBox24.Text)
        strLen = Len(strx)
        .Label109.Caption = ""
        If strLen > 5 Or IsNumeric(strx) = False Then Exit Sub
        If strx = 0 Then Exit Sub
        
        If CheckPN(strx) = True Then
            .Label109.ForeColor = &HC000&
            .Label109.Caption = "Y"
        Else
            .Label109.ForeColor = &HFF&
            .Label109.Caption = "N"
        End If
        .TextBox24.SetFocus
    End With
End Sub

'-------------------------------------------备忘录
Private Sub ListBox4_Click() '备忘录
    Dim i As Byte
    
    With Me
        .ListBox4.Visible = False
        .TextBox10.Visible = True
        i = .ListBox4.ListIndex
        .TextBox10.Text = .ListBox4.Column(1, i)  'listbox采用这种方法来显示值的位置
    End With
End Sub

Private Sub ComboBox9_Change() '备忘录-日期
    Dim timea As Date
    
    timea = Date
    With Me
        If .ComboBox9.Text <> CStr(timea) Then
            .CommandButton34.Enabled = False
            .TextBox10.Locked = True
        Else
            .CommandButton34.Enabled = True    '非当天的日志将无法修改
            .TextBox10.Locked = False
        End If
    End With
End Sub

Sub DateUpdate()                             '备忘录数据实时更新
    Dim dic As New Dictionary
    Dim TableName As String
    Dim i As Byte, k As Byte
    
    TableName = "备忘录"
    SQL = "select * from [" & TableName & "$]"
    Set rs = New ADODB.Recordset    '创建记录集对象
    rs.Open SQL, Conn, adOpenKeyset, adLockOptimistic
    
    'ReDim arr(1 To rs.RecordCount)
    k = rs.RecordCount
    If k = 0 Then Exit Sub
    For i = 1 To k
        dic(CStr(rs.Fields(0))) = ""
        rs.MoveNext '获取数据集的下一个值
    Next
    Me.ComboBox9.List = dic.Keys '只需要获取日期的更新
    rs.Close
    Set rs = Nothing
End Sub
'------------------------------------------------------------------------------备忘录

'---------------------------------------------------------------------------------------------------------------------------------------------------------功能区
Private Sub CommandButton17_Click() '功能区-添加文件夹
    AddFx = 0
    If ListAllFiles(0, "NU") = False Then Me.Label57.Caption = "文件夹受限": Exit Sub
    DataUpdate
End Sub

Private Sub CommandButton1_Click() '添加文件
    AddFx = 0
    If ListAllFiles(0, "NU") = False Then Me.Label57.Caption = "文件夹受限": Exit Sub
    DataUpdate
End Sub

Private Sub CommandButton51_Click() '音乐
    Dim strx As String, strx1 As String
    On Error GoTo 100
    'win10,如果windows media player不可用,看参考改用 bass, http://www.un4seen.com/,支持vb
    With ThisWorkbook
        strx = .Path & "\whitenoise.mp4"
        strx1 = .Sheets("temp").Range("ab8").Value
    End With
    With Me
        If fso.fileexists(strx) = False Or Len(strx1) = 0 Then '检查播放文件是否存在
            .Label57.Caption = "音频文件已丢失"
            Exit Sub
        End If
        If .Label66.Caption = "play" Then Exit Sub '已经处于播放状态
        If .Label66.Caption = "stop" Then
            wm.Controls.Play              '处于停止运行状态
            .Label66.Caption = "play"
            .Label57.Caption = "音乐播放中...."
        Exit Sub
        End If
        If Len(.Label66.Caption) = 0 Then
'            If fso.FileExists(Environ("ProgramW6432") & "\Windows Media Player\wmplayer.exe") = False Then '检查是否存在windows media player 'Environ("ProgramW6432") programfiles
'               .Label57.Caption = "Windows media player不存在，此功能不支持"
'               Exit Sub
'            End If
            Set wm = .Controls.Add("WMPlayer.OCX.7") '创建windows播放控件
            If wm Is Nothing Then
                .Label57.Caption = "wm控件创建失败"
                Set wm = Nothing
                Exit Sub
            End If
            wm.Visible = False '设置为隐藏
            wm.url = strx1
            .Label66.Caption = "play" 'label66用于暂时存储播放的状态值，用于控制按钮
            .Label57.Caption = "音乐播放中..."
        End If
    End With
Exit Sub
100
Me.Label57.Caption = "功能异常"
End Sub

Private Sub CommandButton52_Click() '音乐播放停止
    With Me
        If Len(.Label66.Caption) = 0 Or .Label66.Caption = "stop" Then Exit Sub
        wm.Controls.Stop
        .Label66.Caption = "stop"
        .Label57.Caption = "音乐播放已暂停"
    End With
End Sub

Private Sub CommandButton10_Click() '主页
    Call BackSheet("首页")
End Sub

Private Sub CommandButton11_Click() '书库
    Call BackSheet("书库")
End Sub

Sub BackSheet(ByVal shtn As String) '返回Excel表格中
    Unload Me
    With ThisWorkbook
        If .Application.Visible = False Then .Application.Visible = True
        .Windows(1).WindowState = xlMaximized
        .Sheets(shtn).Activate
    End With
End Sub

Private Sub CommandButton118_Click() '功能区-帮助
    Dim strx As String, strx1 As String, i As Integer

    With ThisWorkbook.Sheets("temp")
        strx = .Cells(39, "ab").Value
        strx1 = .Cells(39, "ac").Value
    End With
    With Me
        If Len(strx) = 0 Then
            .Label57.Caption = "文件已丢失"
        Else
            If Len(strx1) = 0 Then
                i = ErrCode(strx, 1)
                If i < 0 Then MsgBox "帮助文件路径检查异常", vbOKOnly, "Warning": Exit Sub
                If i > 1 Then ThisWorkbook.Sheets("temp").Cells(39, "ac").Value = "ERC": strx1 = "ERC"
            End If
            If fso.fileexists(strx) = False Then
                .Label57.Caption = "文件已丢失"
            Else
                Call OpenFile("N", "help.pdf", "pdf", strx, 1, strx1, 1)
            End If
        End If
    End With
End Sub

Private Sub CommandButton42_Click() '设置按钮-将设置页面一直放在最后
    Dim i As Byte, k As Byte
    
    With Me.MultiPage1
        k = .Pages.Count
        k = k - 2
        For i = 0 To k
            .Pages(i).Visible = False
        Next
        k = k + 1
        .Pages(k).Visible = True  '显示设置页面
        .Value = k
    End With
End Sub

Private Sub CommandButton7_Click() '功能区-Excel模式
    Dim strx As String
    Dim i As Byte, k As Byte, p As Byte
    Dim wd As Object, yesno As Variant
'    ThisWorkbook.Application.ScreenUpdating = False
'    Me.Hide
'    Me.Show 0
'    Call Rewds
'    CopyToClipboard ThisWorkbook.fullname '复制文件的路径到剪切板
'-------------------------------------------------------------------------
    yesno = MsgBox("本程序将退出,转到word,是否继续", vbYesNo, "Tips")
    If yesno = vbNo Then Exit Sub
    yesno = MsgBox("即将关闭保存现有的word文档,是否继续", vbYesNo, "Tips")
    If yesno = vbNo Then Exit Sub
    On Error Resume Next
    If CreateDB = False Then MsgBox "创建数据库失败", vbCritical, "Warning": Exit Sub '------------创建新的数据库
    strx = ThisWorkbook.Path
    strx = strx & "\LB.docm"
    If fso.fileexists(strx) = True Then
        If Err.Number > 0 Then Err.Clear
        Set wd = GetObject(, "word.application") '如果word不处于启动的状态
        If Err.Number > 0 And wd Is Nothing Then
            Err.Clear
            Set wd = CreateObject("word.application")
        End If
        With wd
            i = .documents.Count 'get如果word仅仅是开启,未打开任何的文档, createobject将不会计算现有的文档打开情况
            If i > 0 Then
                For k = 1 To i
                    If strx = .documents(k).fullname Then
                        p = p + 1
                    Else
                        .documents(k).Close savechanges:=True
                    End If
                Next
                If p > 0 Then GoTo 100
            End If
            .Visible = True
            .Activate
            .documents.Open (strx)
        End With
100
        Set wd = Nothing
    Else
        Me.Label57.Caption = "文件丢失"
    End If
    MeQuit
'    ThisWorkbook.Application.ScreenUpdating = True
End Sub

Private Sub MeQuit() '关闭程序
    Unload Me
    If UF4Show > 0 Then Unload UserForm4
    MsgShow "数据自动保存", "Tips", 1200
    With ThisWorkbook
        Call ResetMenu '关闭右键菜单
        With .Sheets("书库")
            .Label1.Caption = ""
            If .CommandButton21.Caption = "退出调试" Then
                .CommandButton21.Caption = "调试模式"
                .CommandButton21.ForeColor = &H80000012
                .CommandButton1.Enabled = True
                .CommandButton11.Enabled = True
            End If
        End With
        If Workbooks.Count > 1 Then
            EnEvents '--------------解除所有的干扰项,防止出现遗漏造成使用Excel的问题
            .Close savechanges:=True
        Else
            .Save
            .Application.EnableEvents = False '注意这里,由于启用workbook.close事件,如果不禁用事件,excel依然还会运行 '禁用事件后,再重新打开excel事件会自动恢复
            .Application.Quit
        End If
    End With
End Sub

Private Sub CommandButton39_Click() '关闭程序
    MeQuit
End Sub

Private Sub CommandButton19_Click() '精简模式
    Me.Hide
    If Workbooks.Count = 1 Then
        If ThisWorkbook.Application.Visible = True Then ThisWorkbook.Application.Visible = False '针对非excel文件作业的环境
        UserForm4.Hide
        UserForm4.Caption = "锁定"
        UserForm4.Show 1
    Else
        ThisWorkbook.Windows(1).WindowState = xlMinimized
        UserForm4.Show
    End If
End Sub
'----------------------------------------------------------------------------------------------功能区

Function CheckFileOpen(ByVal filecode As String) As Boolean '处理返回检查文件打开状态的信息
    Dim i As Byte, strx5 As String
    
    CheckFileOpen = False
    i = FileStatus(filecode) '检查文件的状态
    Select Case i
        Case 1: strx5 = "目录不存在"
        Case 3: strx5 = "文件不存在"
        Case 5: strx5 = "Excel不能打开同名文件"
        Case 6: strx5 = "文件处于打开的状态"
        Case 7: strx5 = "文件处于打开的状态"
        Case 8: strx5 = "异常"
    End Select
    If i = 1 Or i = 3 Then
        DeleFileOverx filecode
    End If
    If i <> 0 Then
        Me.Label57.Caption = strx5
        Set Rng = Nothing
        CheckFileOpen = True
    End If
End Function

Private Sub CommandButton13_Click() '文件执行-需要修改
    Dim i As Integer, p As Byte, xi As Byte, k As Byte, n As Single
    Dim strx3 As String, strx1 As String, strx2 As String, strx4 As String, filez As Long, wt As Integer, _
    Folderpath As String, dics As String, strx5 As String
    
    With Me
        If .Label55.Visible = True Then Exit Sub
        strx1 = .ComboBox6.Text
        strx3 = .Label29.Caption
        strx2 = Filepathi
        strx4 = Filenamei
        strx5 = .Label24.Caption '文件扩展名
        If Len(strx2) = 0 Or Len(strx1) = 0 Then Exit Sub '如果路径为空,文件已被删除,空值
        filez = .Label76.Caption
        '在打开文件的同时,创建txt文件,然后执行vbs脚本,每隔60s(假设),循环查看打开文件的commandline是否被关闭
        '如果捕捉到关闭 , 就将信息写入到txt文件当中去, 关闭vbs执行
        '那么可以间接实现文件打开和关闭的追踪
        If strx1 = "打开" Then
            If strx5 Like "xl*" Then
                If strx5 <> "xls" Or "xlsx" Then MsgBox "禁止打开此类文件", vbCritical, "Warning": Exit Sub
            End If
            If .Label74 = "Y" Then strx3 = "ERC"
            If .CheckBox13.Value = True Then
                xi = 1                    '------文件勾选保护
                Folderpath = ThisWorkbook.Sheets("temp").Cells(41, "ab").Value
                If Len(Folderpath) = 0 Or fso.folderexists(Folderpath) = False Then
                    Folderpath = ThisWorkbook.Path & "\protect"
                    fso.CreateFolder (Folderpath)
                    Folderpath = Folderpath & "\"
                    ThisWorkbook.Sheets("temp").Cells(41, "ab") = Folderpath
                End If
                dics = Left$(Folderpath, 1)
                Folderpath = Folderpath & "\"
                i = FileTest(Folderpath & strx4, strx5, strx4)
                If i >= 4 Then MsgBox "该文件已经存在保护文件,且处于打开的状态": Exit Sub
                If i = 2 Then '文件不存在
                    If filez > fso.GetDrive(dics).AvailableSpace Then MsgBox "磁盘空间不足!", vbCritical, "Warning": Exit Sub '判断磁盘是否有足够的空间
                    fso.CopyFile (strx2), Folderpath, True '覆盖
                End If
                strx2 = Folderpath & strx4 '新的文件路径
            Else
                If CheckFileOpen(strx3) = True Then Exit Sub
            End If
            
            If OpenFile(strx3, strx4, .Label24.Caption, strx2, 1, strx3, xi) = False Then .Label57.Caption = "文件打开失败": Exit Sub
            
            If xi = 1 Then Set Rng = Nothing: Exit Sub
            Call OpenFileOver(strx3) '窗体善后
        
        ElseIf strx1 = "删除" Then '删除涉及,文件是否存在与目录-是否存在于本地-是否处于打开的状态
            If CheckFileOpen(strx3) = True Then Exit Sub '检查文件是否处于打开的状态
            If .Label74.Caption = "Y" Then i = 1 '是否存在非ansi
            If Len(ThisWorkbook.Sheets("temp").Range("ab37").Value) = 0 Then UserForm12.Show '在删除时弹出删除原因窗口
            '当userform 以模式(modal)1显示的时候,后续代码将暂停执行
            If FileDeleExc(Rng.Offset(0, 3).Value, Rng.Row, i, 1) = True Then DeleFileOverx strx3 '执行删除命令成功-执行删除后善后事宜
            
        ElseIf strx1 = "打开位置" Then '判断文件是否存在于本地
            Call OpenFileLocation(Folderpathi)
            
        ElseIf strx1 = "导出文件" Then
            i = FileStatus(strx3, 2)
            Select Case i
                Case 1: strx5 = "目录不存在"
                Case 3: strx5 = "文件不存在"
            End Select
            If i = 1 Or i = 3 Then
'                .Label55.Visible = ture
'                .Label56.Caption = strx3
                DeleFileOverx strx3
                .Label57.Caption = strx5
                Set Rng = Nothing
                Exit Sub
            End If
            If FileCopy(strx2, strx4, Rng.Row, 1) = True Then
                Set Rng = Nothing
                .Label57.Caption = "导出文件中.."
                If filez < 104857600 Then    '100M
                    wt = 200
                Else
                    k = Int(filez / 104857600)
                    n = CSng(k)
                    wt = 200 * (n + n / 20 + n / 50)
                    If wt > 500 Then wt = 500
                End If
                Sleep wt
                .Label57.Caption = "操作成功"
            Else
                .Label57.Caption = "操作失败"
            End If

        ElseIf strx1 = "添加到导出列表" Then
            If CheckLb6(strx3) = False Then
                With .ListBox6
                    If .ListCount > 30 Then MsgBox "添加数量达到上限": Exit Sub
                    .AddItem
                     p = .ListCount - 1
                    .List(p, 0) = strx3 '编号
                    .List(p, 1) = Filenamei '文件名
'                    .List(p, 2) = Filepathi '文件路径
                End With
                .Label57.Caption = "操作成功"
            Else
                .Label57.Caption = "文件已添加"
            End If
        End If
    End With
    Set Rng = Nothing
End Sub

Function CheckLb6(ByVal filecode As String) As Boolean '判断导出列表上是否存在值
    Dim i As Byte, k As Byte
    '也可以创建一个模块级临时数组来暂时存储数据,用以比较
    CheckLb6 = False
    With Me.ListBox6
        k = .ListCount
        If k = 0 Then CheckLb6 = False: Exit Function
        k = k - 1
        For i = 0 To k
            If .List(i, 0) = filecode Then CheckLb6 = True: Exit Function
        Next
    End With
End Function

Function DeleFileOverx(ByVal filecodex As String) '删除文件后执行善后
    Dim i As Byte, itemf As ListItem
    Dim k As Byte, rnga As Range, p As Byte
    
    With Me.ListView1                  '清除主界面中的搜索结果
        If .ListItems.Count <> 0 Then
            Set itemf = .FindItem(filecodex, lvwText, , lvwPartial)
            If itemf Is Nothing Then
                GoTo 1001
            Else
                .ListItems.Remove (itemf.Index) '移除搜索结果
            End If
1001
            Set itemf = Nothing
        End If
    End With
    
    With Me '清除比较中的内容
        .Label56.Caption = filecodex
        If filecodex = .Label29.Caption Then
            .Label55.Visible = True
            DisablEdit '如果删除的内容是编辑界面显示的文件,那么就禁用编辑
        End If
        If Len(.Label179.Caption) > 0 Then
            If .Label179.Caption = filecodex Then p = 1       '清除比较项的内容
        End If
        If Len(.Label153.Caption) > 0 Then
            If .Label153.Caption = filecodex Then p = p + 2
        End If
        If p > 0 Then ClearComp (p)
    End With
'    Call CwUpdate '更新窗体的内容
    With ThisWorkbook.Sheets("主界面")       '主界面,优先阅读修改
        Set rnga = .Range("i27:i33").Find(filecodex, lookat:=xlWhole)
        If rnga Is Nothing Then Exit Function
        k = rnga.Row
        If k = 33 Then '如果是最后一行,那么直接将数据清除
            .Range("d33:l33").ClearContents
        Else
            If k = 27 Then prfile = .Range("i27").Value
            .Range("d" & k & ":" & "l" & k).ClearContents
            For i = k To 33
                If .Range("d" & i) = "" Then Exit For
                .Range("d" & i) = .Range("d" & i + 1)
                .Range("i" & i) = .Range("i" & i + 1)
                .Range("k" & i) = .Range("k" & i + 1)
            Next
        End If
        Me.ListBox2.RemoveItem (k - 27)
    End With
    Set rnga = Nothing
    DataUpdate '更新窗体的数据
End Function

Function ClearComp(ByVal cmCode As Byte) '清除比较内容
    Dim i As Byte, p As Byte
    p = cmCode
    With Me
        If p > 1 Then
            For i = 153 To 165
                .Controls("label" & i).Caption = ""
            Next
        End If
        If p = 1 Or p > 2 Then
            For i = 179 To 192
                .Controls("label" & i).Caption = ""
            Next
        End If
        If p = 3 Then
            For i = 204 To 216
                If i <> 215 Then .Controls("label" & i).Caption = ""
            Next
        End If
    End With
End Function

Private Sub CommandButton2_Click() '主界面-移除优先阅读
    Dim i As Byte, k As Byte, m As Byte
    Dim j As Integer, p As Byte
    
    j = Me.ListBox2.ListIndex
    p = Me.ListBox2.ListCount
    If p = 0 Or j = -1 Then Exit Sub '-1表示没有选择文件 '这里需要注意
    i = 27 + j
    With ThisWorkbook.Sheets("主界面")
        .Range("d" & i & ":" & "l" & i).ClearContents    '注意合并表格的影响
        Me.ListBox2.RemoveItem (j)
        If Len(.Range("d" & i + 1).Value) = 0 Then Exit Sub '如果下一行没有内容就退出
            m = 6 - j
            For k = 1 To m
                .Range("d" & i + k - 1) = .Range("d" & i + k)
                .Range("i" & i + k - 1) = .Range("i" & i + k)
                .Range("k" & i + k - 1) = .Range("k" & i + k)
            Next
            .Range("d" & i + k) = "" '最后一行清除掉
            .Range("i" & i + k) = ""
            .Range("k" & i + k) = ""
    End With
End Sub

Private Sub CommandButton23_Click() '主界面-清除最近阅读
    If Me.ListBox1.ListCount = 0 Then Exit Sub
    With ThisWorkbook.Sheets("主界面")
        .Range("p27:x33").ClearContents    '注意合并表格的影响
    End With
    Me.ListBox1.Clear                   '全部清除
End Sub

Private Sub CommandButton24_Click() '执行多条件筛选
    Dim strx1 As String, strx2 As String, strx3 As String
    Dim strx As String
    
    With Me
        If Len(.ComboBox1.Text) = 0 And Len(.ComboBox7.Text) = 0 And Len(.ComboBox8.Text) = 0 Then Exit Sub
        If .MultiPage1.Value <> 2 Then .MultiPage1.Value = 2 '回到文件夹页面 'listview必须在看见的状态下进行赋值否则会出现位置偏移的问题(在隐藏的状态下赋值只能使用list)
        strx1 = CStr(.ComboBox1.Value)
        strx2 = CStr(.ComboBox7.Value)
        strx3 = CStr(.ComboBox8.Value)
        strx = strx1 & strx2 & strx3
'        If strx1 & strx2 & strx3 = .Label81.Caption Then Exit Sub '同样的筛选条件
'        Call FileListv(strx1, strx2, strx3, 3, 18, 19, 1)
    If Len(storagex) > 0 Then
        If storagex = strx Then Exit Sub
    End If
    FileFilterTemp strx1, , strx2, , strx3
    End With
    storagex = strx
End Sub

Private Function FileFilterTemp(ByVal targetx As String, Optional ByVal cmCode As Byte = 3, Optional ByVal targetx1 As String, _
Optional ByVal cmcode1 As Byte = 1, Optional ByVal targetx2 As String, Optional ByVal cmcode2 As Byte = 2, Optional ByVal cmCodex As Byte = 0) As Byte '用于区分筛选/打开文件夹
                                        '-------------------------------------相比于筛选的方法运算速度更快
    Dim i As Integer, blow As Integer, c As Integer
    Dim x1 As Byte, x2 As Byte, x3 As Byte
    
    ArrayLoad '加载数据
    blow = docmx - 5
    If cmCodex = 1 Then targetx = targetx & "\" '筛选子文件夹
    With Me.ListView2.ListItems
        If cmCodex = 0 Then .Clear
        If cmCodex > 0 Then '打开文件夹
            For i = 1 To blow
                If cmCodex = 1 Then GoTo 98
                If arrax(i, cmCode) = targetx Then
                    If cmCodex = 2 Then GoTo 99
98
                    If InStr(arrax(i, cmCode), targetx) > 0 Then
99
                        With .Add
                            .Text = arrax(i, 1) '编号
                            .SubItems(1) = arrax(i, 2) '文件名
                            .SubItems(2) = arrax(i, 3) '扩展名
                            .SubItems(3) = arrax(i, 5) '位置
                            c = c + 1
                        End With
                    End If
                End If
            Next
        Else '筛选
            x1 = Len(targetx)
            x2 = Len(targetx1)
            x3 = Len(targetx2)
            For i = 1 To blow
                If x1 = 0 Then GoTo 100
                If arrax(i, cmCode) = targetx Then
100
                    If x2 = 0 Then GoTo 101
                    If arrsx(i, cmcode1) = targetx1 Then
101
                        If x3 = 0 Then GoTo 102
                        If arrsx(i, cmcode2) = targetx2 Then
102
                            With .Add
                                .Text = arrax(i, 1) '编号
                                .SubItems(1) = arrax(i, 2) '文件名
                                .SubItems(2) = arrax(i, 3) '扩展名
                                .SubItems(3) = arrax(i, 5) '位置
                                c = c + 1
                            End With
                        End If
                    End If
                End If
            Next
        End If
    End With
    If c = 0 And cmCodex = 0 Then
        Me.Label57.Caption = "未找到对应信息"
    ElseIf c > 0 And cmCodex > 0 Then
        FileFilterTemp = 1
    End If
End Function

Private Sub CommandButton27_Click() '添加到优先阅读-修改
    Dim strx As String
    With Me
        If .Label55.Visible = True Then Exit Sub
        strx = .Label29.Caption
        If Len(strx) = 0 Then Exit Sub
        If AddPList(strx, .Label23.Caption, 1) = True Then
            Call PrReadList
'            .Label57.Caption = "添加成功"
        Else
            .Label57.Caption = "已添加"
        End If
    End With
End Sub

Private Sub CommandButton28_Click() '编辑信息,防止误操作
  Call EnablEdit
End Sub

Private Sub CommandButton29_Click() '主界面-转到详细页 -这里的设置将根据搜索来进行调整,如果搜索修改成模糊搜索这里也需要变动
    Dim strx As String, k As Integer, xi As Byte
    Dim i As Integer
    With Me
        k = .ListView1.ListItems.Count
        If Len(Trim(.TextBox8.Text)) = 0 Or k = 0 Then Exit Sub
        With .ListView1
            For i = 1 To k
                If .ListItems(i).Selected = True Then xi = .SelectedItem.Index: Exit For
            Next
        End With
        If xi = 0 Then Exit Sub
        strx = .ListView1.SelectedItem.Text
        If strx = .Label29.Caption And .Label55.Visible = False Then .MultiPage1.Value = 1: Exit Sub '已经查询结果
        If strx = .Label56.Caption Then
            .Label57.Caption = "文件已被删除"
            .ListView1.ListItems.Remove (xi)     '可以进一步扩展搜索删除备份里的信息'可以进一步在删除文件之前获取文件的md5,用于判断后续的文件添加进来是否重复
            Exit Sub
        End If
        SearchFile strx      '搜索文件
        If Rng Is Nothing Then DeleFileOverx strx: Exit Sub
    End With
    Call ShowDetail(strx)
    Set Rng = Nothing
End Sub

Private Sub CheckBox8_Click() '编辑模式是否启用-相应的编辑按钮将全部启用,方便编辑信息,不需要手东启动按钮
    If clickx = 1 Then clickx = 0: Exit Sub
    With ThisWorkbook.Sheets("temp")
        If Me.CheckBox8.Value = True Then
            .Cells(31, "ab") = 1
        Else
            .Cells(31, "ab") = ""
        End If
    End With
End Sub

'-------------------------------------------------搜索/编辑-搜索
Private Sub CommandButton3_Click() '百度搜索
    SearchCheck 1
End Sub

Private Sub CommandButton5_Click() '豆瓣搜索
   SearchCheck 2
End Sub

Private Sub CommandButton6_Click() '维基百科
    Dim yesno As Variant
    yesno = MsgBox("此站点已经404, 是否继续打开", vbYesNo, "Warning")
    If yesno = vbNo Then Exit Sub
    SearchCheck 4
End Sub

Function MultiSearch(ByVal engine As Byte, Keyword As String) As String
    Dim SearchEngine As String
    Dim Urlx As String
    
    Select Case engine
    Case 1
        SearchEngine = "https://www.baidu.com/s?wd="  '百度搜索
        Keyword = Replace(Keyword, " ", "%20")
        Urlx = SearchEngine & Keyword
    Case 2
        SearchEngine = "https://www.douban.com/search?q="  'douban搜索
        Keyword = Replace(Keyword, " ", "+")
        Urlx = SearchEngine & Keyword
    Case 3
        SearchEngine = "https://www.bing.com/search?q="  'bing
        Keyword = Replace(Keyword, " ", "%20")
        Urlx = SearchEngine & Keyword
    Case 4
        SearchEngine = "https://en.wikipedia.org/w/index.php?search="   'wiki
        Keyword = Replace(Keyword, " ", "%20")
        Urlx = SearchEngine & Keyword
    Case 5
        SearchEngine = "" '预留
        Keyword = Replace(Keyword, " ", "%20")
        Urlx = SearchEngine & Keyword
    End Select
    MultiSearch = Urlx
End Function

Function SearchCheck(ByVal cmCode As Byte)  '检查搜索的内容
    Dim strx1 As String, strx As String
    Dim strLen As Byte
    Dim i As Byte, k As Byte, j As Byte, m As Byte
    
    With Me
        If .CheckBox21.Value = True Then i = 1 '如果勾选这两项
        If .CheckBox21.Value = True Then k = 1
        If i = 1 Or k = 1 Then
            m = 2
            If i = 1 Then               '优先搜索主文件名
                strx1 = Trim(.TextBox3.Text)
            ElseIf k = 1 Then
                strx1 = Trim$(Left$(Filenamei, InStrRev(Filenamei, ".") - 1)) '去掉扩展名
            End If
        Else
            strx1 = Trim(.TextBox1.Text)
            If strx1 Like "HLA*" Then Exit Function
        End If
        strLen = Len(strx1)
        If strLen = 0 Then Exit Function
        strx = MultiSearch(cmCode, strx1)
        If cmCode = 4 Then
            For j = 1 To strLen
                If Not Mid(strx1, j, 1) Like "[a-zA-Z]" Then .Label57.Caption = "搜索的内容包含非英文字符,位置:" & j: Exit Function '请求来自wikipedia/只允许英文
            Next
        End If
        If Len(ThisWorkbook.Sheets("temp").Cells(45, "ab").Value) > 0 Then m = 2 '使用内置浏览器
        Webbrowser strx, m
    End With
End Function
''----------------------------------------------------------------------------------------------搜索/编辑-搜索

'-----------------------------------------------------------主界面-备忘录
Private Sub CommandButton137_Click() '备忘录-添加信息
    Dim timea As Date, timeb As Date, str As String
    
    timea = Date
    timeb = Format(time, "hh:mm:ss")
    str = Me.TextBox10.Text
    If Me.TextBox10.Text <> "" And Me.ComboBox9.Text <> "" Then
        SQL = "Insert into [备忘录$] (日期,时间,内容) Values (#" & timea & "#,'" & CStr(timeb) & "', '" & str & "')"
        Conn.Execute (SQL)
        Call DateUpdate
    Else
'        Call Warning(6)
    End If
End Sub

Private Sub CommandButton32_Click() '备忘录-新建
    Dim timea As Date
    timea = Date      'date为日期函数
    With Me
        .TextBox10.Enabled = True
        .TextBox10.Visible = True
        .ListBox4.Visible = False
        If .ComboBox9.Text <> CStr(timea) Then .ComboBox9.Text = CStr(timea) 'cstr为转换文本函数
        .TextBox10.Text = ""    '保证新的文本框是空白的
    End With
End Sub

Private Sub CommandButton33_Click() '备忘录-查看
    Dim timea As Date, i As Byte, k As Byte
    Dim arr()
    Dim TableName As String
    
    TableName = "备忘录"
    timea = Me.ComboBox9.Text
    With Me.ListBox4
        If RecData = True Then
            .Clear
            SQL = "select * from [" & TableName & "$] where 日期=#" & timea & "#"
            Set rs = New ADODB.Recordset    '创建记录集对象
            rs.Open SQL, Conn, adOpenKeyset, adLockOptimistic
            If rs.BOF And rs.EOF Then   '用于判断有无找到数据
'                Call Warning(6)
                GoTo 100
            End If
            .Visible = True
            Me.TextBox10.Visible = False
            k = rs.RecordCount
            ReDim arr(1 To k, 1 To 2)
            For i = 1 To k
                arr(i, 1) = rs(1)
                arr(i, 2) = rs(2)
                rs.MoveNext
            Next
            
            For i = 1 To k
                .AddItem
                .List(.ListCount - 1, 0) = arr(i, 1) '分列显示
                .List(.ListCount - 1, 1) = arr(i, 2)
            Next
100
            rs.Close
            Set rs = Nothing
        Else
        Me.Label57.Caption = "异常"
        End If
    End With
End Sub

Private Sub CommandButton34_Click() '备忘录-添加信息
    Dim timea As Date, timeb As String, str As String
    
    timea = Date
    timeb = Format(time, "h:mm:ss")
    str = Trim(Me.TextBox10.Text)
    If Len(str) > 0 And RecData = True Then
        SQL = "Insert into [备忘录$] (日期,时间,内容) Values (#" & timea & "#,'" & timeb & "', '" & str & "')"
        Conn.Execute (SQL)
'        Call Warning(1)
        Call DateUpdate
    Else
'        Call Warning(6)
    End If
End Sub
'-----------------------------------------------------------主界面-备忘录

Private Sub CommandButton43_Click() '打开文件夹
    Dim arrt() As String, k As Byte, i As Byte, xi As Byte, strx1 As String, strx2 As String
    Dim p As Byte, m As Byte, chk As Byte
    Dim strx As String
    
    With Me.ListBox3
        k = .ListCount
        If k = 0 Then Exit Sub  'Or .ListIndex < 1
        k = k - 1
        If Me.CheckBox9.Value = True Then '勾选子文件夹
            m = 4: p = 1
        Else
            ReDim arrt(0 To k)
            m = 5: p = 2
        End If
        For i = 0 To k
            If .Selected(i) = True Then
                If InStr(.List(i, 1), "移除") = 0 Then '移除掉不存在的文件夹的内容
                    If m = 4 Then
                        strx = arraddfolder(i)           '.List(i, 0) '获取被选中的文件夹 'arraddfolder(i)
                        Exit For
                    Else
                        xi = xi + 1
                        arrt(i) = arraddfolder(i)
                    End If
                End If
            End If
        Next
    End With
    If xi = 0 And Len(strx) = 0 Then Exit Sub
    With Me
        .MultiPage1.Value = 2          ',listview '无法在隐藏状态下进行赋值,否则会出现位置偏移的问题
        If pgx2 = 1 Then .ListView2.ListItems.Clear
        If m = 4 Then
            chk = chk + FileFilterTemp(strx, m, , , , , p)
        Else
            For i = 0 To k
                If Len(arrt(i)) > 0 Then chk = chk + FileFilterTemp(arrt(i), m, , , , , p)
            Next
        End If
    End With
    If chk = 0 Then Me.Label57.Caption = "文件夹为空"
'        If .CheckBox9.Value = True Then
'            m = 4: p = 2
'        Else
'            m = 5: p = 1
'        End If
'        If m = 4 And xi > 1 Then .Label57.Caption = "不允许选择多个文件夹": Exit Sub
'        xi = xi - 1
'        .ListView2.ListItems.Clear '对页面进行清空
'        For i = 0 To xi
'            Call FileListv(arrt(i), strx1, strx2, m, 0, 0, 0, p)
'        Next
'    EnEvents
End Sub

Function FileListv(ByVal str As String, ByVal str1 As String, ByVal str2 As String, ByVal filtercode As Byte, _
ByVal filtercode1 As Byte, ByVal filtercode2 As Byte, ByVal methcode As Byte, Optional ByVal cmCode As Byte)        '筛选列表信息
    Dim rngf As Range
    Dim k As Integer, spl As Integer, j As Byte, p As Byte, m As Byte, flow As Integer, alow As Integer
    Dim rngx As Range, rngxs As Range
    Dim arrlist() As Variant, strx3 As String, strx As String, str3 As String, blow As Integer
    
    On Error GoTo ErrHandle
    With ThisWorkbook.Sheets("书库")       '当数据量特别大的时候,这种方法太慢
'        flow = .[f65536].End(xlUp).Row
        flow = docmx
        If Len(str) > 0 Then
            If methcode = 0 Then
                Set rngf = .Range("f6:f" & flow).Find(str, lookat:=xlWhole) '筛选的目标为空 '文件夹
                If rngf Is Nothing Then GoTo 101
                    j = 1
                Else
                    Set rngf = .Range("d6:d" & flow).Find(str, lookat:=xlWhole) '筛选的目标为空 '文件后缀
                    If rngf Is Nothing Then
                        GoTo 101
                    Else
                        j = 1
                    End If
            End If
        End If
        
        If methcode = 1 Then
            If Len(str1) > 0 Then
                Set rngf = .Range("s6:s" & .[s65536].End(xlUp).Row).Find(str1, lookat:=xlWhole) '内容评分
                If rngf Is Nothing Then
                    GoTo 101
                Else
                    p = 2
                End If
            End If
            If Len(str2) > 0 Then
            Set rngf = .Range("t6:t" & .[t65536].End(xlUp).Row).Find(str2, lookat:=xlWhole) '推荐指数
                If rngf Is Nothing Then
                    GoTo 101
                Else
                    m = 3
                End If
            End If
        End If
        '--------------------------------------------------------------------------------------------先进行文件查找,看看目录是否有对应的值
        If j = 0 And p = 0 And m = 0 Then GoTo 101
        DisEvents '------------------------------selection事件很容易被触发
        If .AutoFilterMode = True Then .AutoFilterMode = False '筛选如果处于开启状态则关闭
        Set rngx = .Range("b5:v" & flow)
        If cmCode = 0 Or cmCode = 1 Then '筛选发出的请求
            If j = 1 Then rngx.AutoFilter Field:=filtercode, Criteria1:=str
            If p = 2 Then rngx.AutoFilter Field:=filtercode1, Criteria1:=str1
            If m = 3 Then rngx.AutoFilter Field:=filtercode2, Criteria1:=str2
        Else
            strx3 = str & "\" '-----请求来自打开文件夹
            rngx.AutoFilter Field:=filtercode, Criteria1:="=" & strx3 & "*", Operator:=xlOr '筛选子文件夹 'Excel的筛选可以对多条件进行筛选
        End If
        '-----------当数据处于第一行, 筛选出来的结果也是第一行时, 会出现溢出错误即筛选后的.[b65536].End(xlUp).Row等于原来的行号, 如第6行就是筛选出来的结果
        blow = .[b65536].End(xlUp).Row
        If blow = 6 Then
            spl = 1
        Else
            spl = .Range("b6:b" & blow).SpecialCells(xlCellTypeVisible).Count
        End If
        Set rngxs = rngx.SpecialCells(xlCellTypeVisible)
        With ThisWorkbook.Sheets("temp")
            rngxs.Copy .Range("a1") '-------------------将筛选出来的值复制到临时的表格上(因为筛选出来的结果通常是不连续表格的数据,无法一次性赋值给数组)
            arrlist = ThisWorkbook.Application.Transpose(.Range("a2:d" & spl + 1).Value) '添加到数组中
        End With
        '---------------------------------------------获取筛选的结果
        If cmCode = 0 Then Me.ListView2.ListItems.Clear '使用前清除所有的内容' 如果是打开多个文件夹就不清空
        With Me.ListView2.ListItems
            For k = 1 To spl
                With .Add
                    .Text = arrlist(1, k)
                    .SubItems(1) = arrlist(2, k)
                    .SubItems(2) = arrlist(3, k)
                    .SubItems(3) = arrlist(4, k)
                End With
            Next
        End With
        .Range("f5:f" & flow).AutoFilter
101
    End With
    ThisWorkbook.Sheets("temp").Range("a1:z" & spl + 1).ClearContents '清除缓存的内容
    Set rngx = Nothing
    Set rngxs = Nothing
    Set rngf = Nothing
    Me.Label81.Caption = str & str1 & str2 '临时存储的值用于控制代码的执行
    If cmCode = 0 Then EnEvents
    Exit Function
ErrHandle:
    Me.Label57.Caption = Err.Number
    Err.Clear
    EnEvents
End Function

Private Sub CommandButton36_Click() '更新文件夹
    Dim arrt() As String, k As Byte, i As Byte, xi As Byte, p As Byte
    
    With Me.ListBox3
        k = .ListCount
        If k = 0 Or .ListIndex = -1 Then Exit Sub
        k = k - 1
        ReDim arrt(0 To k)
        For i = 0 To k
            If .Selected(i) = True Then
                If InStr(.List(i, 1), "移除") = 0 Then xi = xi + 1: arrt(i) = .List(i, 0) '获取被选中的文件夹
            End If
        Next
    End With
    If xi = 0 Then Exit Sub
    With Me
        .MultiPage1.Value = 2          ',listview '无法在隐藏状态下进行赋值,否则会出现位置偏移的问题
        If .CheckBox9.Value = True Then
            p = 1
        Else
            p = 2
        End If
        xi = xi - 1
        AddFx = 0
        For i = 0 To xi
            Call ListAllFiles(p, arrt(i))
        Next
    End With
    DataUpdate
End Sub

Private Sub CommandButton38_Click() '移除文件夹
    Dim yn As Variant
    Dim i As Integer, j As Integer, k As Byte, p As Byte, xi As Byte, n As Byte
    Dim rngf As Range
    Dim flow As Integer, strx As String, strx2 As String
    
    k = Me.ListBox3.ListCount
    j = Me.ListBox3.ListIndex
    If k = 0 Or j = -1 Then Exit Sub
    yn = MsgBox("此操作将同时移除书库的目录(不会删除本地文件)?_", vbYesNo) 'msgbox可选yes or no
    If yn = vbNo Then Exit Sub
    
    With ThisWorkbook
        With .Sheets("书库")
            flow = .[f65536].End(xlUp).Row
            p = k - 1
            For i = p To 0 Step -1 '删除一般采用倒删，避免删除后出现位置偏移的问题
                If Me.ListBox3.Selected(i) = True Then '如果被选定-支持多文件夹选择
                    n = n + 1
                    strx = Me.ListBox3.Column(0, i)
                    strx2 = strx & "\"
                    Set rngf = .Range("e6:e" & flow).Find(strx2, lookat:=xlPart) '筛选的目标为空
                    If rngf Is Nothing Then GoTo 100
                    If .AutoFilterMode = True Then .AutoFilterMode = False '筛选如果处于开启状态则关闭
                    strx2 = "=" & strx2 & "*" '------------------------------使用Excel的筛选,-包含
                    .Range("e5:e" & flow).AutoFilter Field:=1, Criteria1:=strx2, Operator:=xlAnd  '----模糊筛选包含,注意搜索的区域要选择e列, 即文件路径所在的列
                    .Range("e5").Offset(1).Resize(flow - 5).SpecialCells(xlCellTypeVisible).Delete Shift:=xlShiftUp '删除掉筛选出来时筛的结果
                    .Range("e5").AutoFilter
100
                    With ThisWorkbook.Sheets("主界面")
                        xi = 37 + i
                        .Range("e" & xi).Delete Shift:=xlUp
                        .Range("i" & xi).Delete Shift:=xlUp
                    End With
                    Call DeleFileOver(strx, 1) '移除目录
                    Me.ListBox3.RemoveItem (i)
                End If
            Next
        End With
    End With
    Set rngf = Nothing
    addfilec = flow - n
    Me.Label57.Caption = "操作成功"
    ThisWorkbook.Save
End Sub

Private Sub CommandButton4_Click() '查询书库/单词-执行搜索和数据获取
    Dim strx As String, strx1 As String
    Dim strLen As Byte, keyworda As String
    
    With Me
        strx = Trim(.TextBox1.Text)
        strLen = Len(strx)
        If strLen = 0 Then Exit Sub
        If Len(.Label29.Caption) > 0 Then
            If .Label29.Caption = strx1 Then Exit Sub '当数据已经查询之后
        End If
        If InStr(strx, "HLA") = 0 Then Exit Sub
        If InStr(strx, "&") > 0 Then strx1 = Trim(Split(strx, "&")(0))
        If strx1 = .Label56.Caption Then Exit Sub
        If strx Like "HLA-000*&*" Then keyworda = strx1
        SearchFile keyworda
        If Rng Is Nothing Then .Label57.Caption = "文件不存在": Exit Sub
        ShowDetail (keyworda)
    End With
    Set Rng = Nothing
End Sub

Private Sub CommandButton4_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '单词/书库切换-未完成
Exit Sub
    With Me
    If Len(.TextBox1.Text) = 0 Then
        .CommandButton4.Caption = "查询书库"
        .CommandButton44.Visible = False '发音
    End If
    End With
End Sub

Private Sub TextBox11_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '设置-双击选择
    Dim strfolder As String, str As String, strx As Byte
    Dim fdx As FileDialog
    Dim selectfile As String

    With Me
        str = .ComboBox11.Text
        Select Case str
            Case "浏览器": strx = 1
            Case "Axure": strx = 1
            Case "Mind": strx = 1
            Case "Note": strx = 1
            Case "PDF": strx = 1
            Case "截图": strx = 1
            Case "Spy++": strx = 1
            Case "备份": strx = 2
            Case "解压目录": strx = 2
            Case Else: strx = 0
        End Select
        
        If strx = 1 Then
            Set fdx = Application.FileDialog(msoFileDialogFilePicker) '文件选择窗口
            With fdx
                .AllowMultiSelect = False
                .Show
                .Filters.Clear '清除过滤规则
                .Filters.Add "Application", "*.exe" '过滤exe文件
                If .SelectedItems.Count = 0 Then Exit Sub
                selectfile = .SelectedItems(1)
            End With
            filepathset = selectfile
            .TextBox11.Text = selectfile
            Set fdx = Nothing
        ElseIf strx = 2 Then
            With Application.FileDialog(msoFileDialogFolderPicker) '文件夹选择窗口(文件夹以此只能选一个)
                .Show
                If .SelectedItems.Count = 0 Then Exit Sub '未选择文件夹则退出sub
                strfolder = .SelectedItems(1)
            End With
            folderpathset = strfolder
            .TextBox11.Text = strfolder
        ElseIf strx = 0 Then
            .Label57.Caption = "设置有误"
        End If
    End With
End Sub

Private Sub ComboBox11_Click() '设置-工具设置
    Dim str As String, strx As String

    str = Me.ComboBox11.Value
    With ThisWorkbook.Sheets("temp")
        Select Case str
            Case "浏览器": strx = .Range("ab10").Value
            Case "Axure": strx = .Range("ab11").Value
            Case "Mind": strx = .Range("ab12").Value
            Case "Note": strx = .Range("ab13").Value
            Case "PDF": strx = .Range("ab14").Value
            Case "截图": strx = .Range("a15").Value
            Case "备份": strx = .Range("ab16").Value
            Case "解压目录": strx = .Range("ab42")
            Case "Spy++": strx = .Range("ab51")
        End Select
        If Len(strx) > 0 Then
            Me.TextBox11.Text = strx
        Else
            Me.TextBox11.Text = "未设置"
        End If
        If str = "备份" Then
            Me.TextBox22.Enabled = True
            If Len(.Range("ab27").Value) > 0 Then
            If IsNumeric(.Range("ab27").Value) = True Then Me.TextBox22.Text = .Range("ab27").Value
            End If
        Else
            Me.TextBox22.Text = ""
            Me.TextBox22.Enabled = False
        End If
    End With
End Sub

Private Sub CommandButton41_Click() '设置-修改设置
    Dim str As String, strx As String, limitnum As Byte, strx2 As String

    On Error GoTo 100
    With Me
        strx = Trim(.TextBox11.Text)
        str = Trim(.ComboBox11.Value)
        If Len(strx) = 0 Or strx = "工具设置" Or strx = "未设置" Then Exit Sub  '无效内容
        If str = "备份" Or str = "解压目录" Then
            strx = folderpathset
            If fso.folderexists(strx) = False Then '备份文件夹设置
                .Label57.Caption = "此文件夹不存在"
                .TextBox11.SetFocus
                Exit Sub
            End If
            If strx = ThisWorkbook.Path Then
                .Label57.Caption = "不允许设置和程序同一个文件夹"
                .TextBox11.Text = ""
                .TextBox11.SetFocus
                Exit Sub
            End If
            If Right(strx, 1) = "\" Then
                .Label57.Caption = "文件夹设置后面不需要\符号"
                .TextBox11.SetFocus
                Exit Sub
            End If
            If CheckFileFrom(strx) = True Then
                .Label57.Caption = "系统盘受限位置"
                .TextBox11.Text = ""
                .TextBox11.SetFocus
                Exit Sub
            End If
            If str = "备份" Then
                If fso.GetFolder(strx).Files.Count > 0 Then
                    .Label57.Caption = "此文件夹已存在文件"
                    .TextBox11.Text = ""
                    .TextBox11.SetFocus
                End If
                If Len(.TextBox22.Text) > 0 Then
                    If IsNumeric(.TextBox22.Text) = True Then
                       limitnum = Int(.TextBox22.Text)
                       If limitnum < 6 Then
                            .Label57.Caption = "设置过低,设置值应大于5" '默认设置为10
                            GoTo 100
                        ElseIf limitnum > 30 Then
                            .Label57.Caption = "设置过低,设置值应小于30"
                            GoTo 100
                        End If
                        ThisWorkbook.Sheets("temp").Range("ab27") = limitnum
                   End If
                End If
            End If
            folderpathset = ""
        Else
            strx = filepathset
            If InStr(strx, "exe") = 0 Or InStr(strx, "\") = 0 Or fso.fileexists(strx) = False Then '程序设置
                .Label57.Caption = "程序不存在"
                .TextBox11.SetFocus
                Exit Sub
            End If
        End If
        filepathset = ""
         .Label57.Caption = "设置成功"
    End With
    With ThisWorkbook.Sheets("temp") '将设置写入表格
        Select Case str
            Case "浏览器": .Range("ab10") = strx
            Case "Axure": .Range("ab11") = strx
            Case "Mind": .Range("ab12") = strx
            Case "Note": .Range("ab13") = strx
            Case "PDF": .Range("ab14") = strx
            Case "截图": .Range("a15") = strx
            Case "备份": .Range("ab16") = strx
            Case "解压目录": .Range("ab42") = strx
            Case "Spy++": .Range("ab51") = strx
        End Select
    End With
    Exit Sub
100
    Me.Label57.Caption = "设置异常,代码: " & Err.Number
    Err.Clear
End Sub

Private Sub CheckBox1_Click() '设置-勾选发音-设置
    If clickx = 1 Then clickx = 0: Exit Sub
    With ThisWorkbook.Sheets("temp")
        If Me.CheckBox1.Value = True Then
            .Range("ab18") = 1
            voicex = 1
        Else
            .Range("ab18") = ""
            voicex = 0
        End If
    End With
End Sub

Private Sub CheckBox7_Click() '文件md5自动写入数据
    If clickx = 1 Then clickx = 0: Exit Sub
    With ThisWorkbook.Sheets("temp")
        If Me.CheckBox7.Value = True Then
            .Range("ab29") = 1
        Else
            .Range("ab29") = ""
        End If
    End With
End Sub

Private Sub CommandButton44_Click() '未完成-英文发音
Exit Sub
'    If Len(Me.TextBox1.text) = 0 Or Me.CheckBox1.Value = False Then Exit Sub '如果空或者不勾选则退出
'    If IsNumeric(Me.TextBox1.text) Or Me.TextBox1.text Like "*[一-龥]*" Then Exit Sub '纯数字或者包含中文
'    If Me.CommandButton4.Caption = "单词查询" Then Application.Speech.Speak (Me.TextBox1.text)
End Sub

Private Sub CommandButton45_Click() '单词-发音
    Dim strx As String
    
    strx = Me.TextBox14.Text
    If Len(strx) = 0 Then Exit Sub '如果空或者不勾选则退出
    If strx Like "*#*" Or strx Like "*[一-龥]*" Then Exit Sub '包含数字或者中文/不同的系统安装的语言发音模块不一样
    Call Speakvs(strx)
End Sub

Private Sub CommandButton46_Click() '添加单词-未完成

Exit Sub '尚未完成
If Len(Me.TextBox13.Text) > 0 And Len(Me.TextBox14.Text) > 0 Then

'Call Warning(1)
End If
End Sub
'------------------------------------------------------------------------------------------本地工具
Function SendTools(ByVal toolx As Byte) '选择需要执行的程序
    Dim xtool As String, exepath As String
    
    With ThisWorkbook.Sheets("temp")
        Select Case toolx
            Case 1: exepath = .Range("ab11").Value
            Case 2: exepath = .Range("ab12").Value
            Case 3: exepath = .Range("ab13").Value
            Case 6: exepath = .Range("ab14").Value
            Case 7: exepath = .Range("ab15").Value
            Case 8: exepath = .Range("ab51").Value
        End Select
    End With
    If Len(exepath) = 0 Or InStr(exepath, "exe") = 0 Or fso.fileexists(exepath) = False Then
        Me.Label57.Caption = "未设置程序"
        Exit Function
    End If
    xtool = exepath & Chr(32)
    Shell xtool, vbNormalFocus
End Function

Private Sub CommandButton142_Click()
    Call SendTools(8)
End Sub
Private Sub CommandButton81_Click() '工具-powershell ISE
    If Len(ThisWorkbook.Sheets("temp").Range("ab5").Value) = 0 Then
        Me.Label57.Caption = "不支持此功能"
        Exit Sub
    End If
    Shell ("PowerShell_ISE "), vbNormalFocus
End Sub

Private Sub CommandButton78_Click() '工具-powershell
    If Len(ThisWorkbook.Sheets("temp").Range("ab4").Value) = 0 Then
        Me.Label57.Caption = "不支持此功能"
        Exit Sub
    End If
    Shell ("powershell "), vbNormalFocus '打开powershell
End Sub

Private Sub CommandButton82_Click() '工具-vbe
    Unload Me
    With ThisWorkbook
        If .Application.Visible = False Then .Application.Visible = True
        If .Windows(1).Visible = False Then .Windows(1).Visible = True
        .Application.SendKeys ("%{F11}")
    End With
End Sub

Private Sub CommandButton83_Click() '工具-思维导图
    Call SendTools(2)
End Sub

Private Sub CommandButton84_Click() '工具-axure
    Call SendTools(1)
End Sub

Private Sub CommandButton31_Click() '工具-截图工具
    Call SendTools(7)
End Sub

Private Sub CommandButton85_Click() '工具-onenote
    Call SendTools(3)
End Sub

Private Sub CommandButton70_Click() '工具-pdf编辑
    Call SendTools(6)
End Sub

Private Sub CommandButton79_Click() '工具-备份
    Dim strx As String, limitnum As Integer, i As Integer, k As Integer
    Dim fd As Folder, fl As File
    Dim timea As Date, Oldestfile As String
    
    With ThisWorkbook.Sheets("temp")
        strx = .Range("ab26").Value
        
        If Len(strx) > 0 And fso.folderexists(strx) = True Then
            timea = Now '对比的初始值
            ThisWorkbook.Save   '保存文件
            fso.CopyFile (ThisWorkbook.fullname), strx & "\", overwritefiles:=True  '复制文件到新的文件夹
            fso.GetFile(strx & "\" & ThisWorkbook.Name).Name = CStr(Format(timea, "yyyymmddhmmss")) & ".xlsm" '将文件按照日期进行重命名
            .Range("ab19") = Now '备份的时间
        Else
            Me.Label57.Caption = "文件夹设置有误"
            Exit Sub
        End If
        
        If Len(.Range("ab27").Value) > 0 And IsNumeric(.Range("ab27").Value) = True Then '限制文件夹内的备份文件的数量
            limitnum = .Range("ab28").Value
            If limitnum < 6 Or limitnum > 30 Then limitnum = 10 '限制文件的下限5,上限30
        Else
            limitnum = 10 '默认为10
        End If
        
        Set fd = fso.GetFolder(strx)
        k = fd.Files.Count
        If k > limitnum Then '文件数量超过上限
            For Each fl In fd.Files '找出最旧的文件
                If fl.DateCreated < timea Then
                    timea = fl.DateCreated
                    Oldestfile = fl.Path
                End If
            Next
            fso.DeleteFile (Oldestfile) '删除掉最旧的文件
        End If
        Me.Label57.Caption = "备份成功"
    End With
    Set fd = Nothing
End Sub           '----------------------------------------------------------------------------工具-本地工具

Private Sub Label106_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '豆瓣链接
    Dim strx As String

    strx = Me.Label106.Caption
    If Len(strx) = 0 Or InStr(strx, "http") = 0 Then Exit Sub
    browserkey = strx
    Me.MultiPage1.Value = 6
    CopyToClipboard strx
    '---------------------打开豆瓣的网站稍慢(豆瓣的网站问题,不管是chrome/Firefox,IE,不管是否开启广告过滤,打开豆瓣都不会很快(豆瓣这个站点采用类似edge(旧款)加载模式,等到内容差不多加载完,才完整显示内容)
'    Me.Label57.Caption = "正在打开网站中..稍后"
    '--------------------豆瓣的链接全部有页面内的浏览器打开
'    Call WebBrowser(strx)
End Sub

Private Sub TextBox17_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '双击打开豆瓣的链接
    Dim strx As String

    strx = Me.TextBox17.Text
    If Len(strx) = 0 Or InStr(strx, "http") = 0 Then Exit Sub
    browserkey = strx
    Me.MultiPage1.Value = 6
    CopyToClipboard strx
'    Call WebBrowser(strx)
End Sub
'---------------------------------------------------------工具-网络工具

Function SendUrl(ByVal httpx As Byte, Optional ByVal cmCode As Byte) '打开的网站选择
    Dim Urlx As String
    
    Select Case httpx
        Case 1: Urlx = "http://www.iciba.com/"
        Case 2: Urlx = "https://note.youdao.com/web"
        Case 3: Urlx = "https://shouqu.me/my.html"
        Case 4: Urlx = "https://docs.microsoft.com/zh-cn/office/vba/api/overview/"
        Case 5: Urlx = "https://www.pstips.net/"
        Case 6: Urlx = "https://stackoverflow.com/"
        Case 7: Urlx = "http://club.excelhome.net/forum-2-1.html"
    End Select
    Call Webbrowser(Urlx, cmCode)
End Function

Private Sub CommandButton80_Click() '金山词典
    Call SendUrl(1)
End Sub

Private Sub CommandButton25_Click() '有道笔记
    Call SendUrl(2, 1)
End Sub

Private Sub CommandButton26_Click() '收趣书签
    Call SendUrl(3, 1)
End Sub

Private Sub CommandButton86_Click() 'VBA docs
    Call SendUrl(4)
End Sub

Private Sub CommandButton87_Click() 'ps tips
    Call SendUrl(5)
End Sub

Private Sub CommandButton89_Click() 'stack
    Call SendUrl(6)
End Sub

Private Sub CommandButton88_Click() 'excel
    Call SendUrl(7)
End Sub
'------------------------------------------------------------------------------工具-网络工具

Private Sub CommandButton90_Click() '工具-重启
    With ThisWorkbook
        If .Application.Visible = False Then .Application.Visible = True
        If .Windows(1).Visible = False Then .Windows(1).Visible = True
    End With
    End 'end命令的执行相当于vbe里面的重置按钮,所有的sub将结束,变量完全释放
End Sub

Private Sub ListView2_Click() '点击的时候,全部选中
    With Me.ListView2
        If .ListItems.Count = 0 Then Exit Sub
        .FullRowSelect = True
    End With
End Sub

Private Sub ListView2_DblClick() '文件-双击打开文件夹
    Dim k As Byte, n As Byte, p As Byte, i As Byte
    Dim strx As String, strx1 As String, strx2 As String
    
    With Me.ListView2
        If .ListItems.Count = 0 Then Exit Sub
        strx = .SelectedItem.Text
        strx1 = .SelectedItem.ListSubItems(1).Text
        strx2 = LCase(Right$(strx1, Len(strx1) - InStrRev(strx1, ".")))
        If strx2 Like "xl*" Then
            If strx2 <> "xls" Or "xlsx" Then MsgBox "禁止打开此类文件", vbCritical, "Warning": Exit Sub
        End If
        If CheckFileOpen(strx) = True Then Exit Sub '判断文件是否存在于目录或者本地磁盘/文件是否处于打开的状态
    End With
    With Rng
        If OpenFile(strx, .Offset(0, 1).Value, .Offset(0, 2).Value, .Offset(0, 3).Value, 1, .Offset(0, 27).Value) = False Then
            Me.Label57.Caption = "异常"
            Set Rng = Nothing
            Exit Sub
        End If
    End With
    Call OpenFileOver(strx)
    Set Rng = Nothing
End Sub

Private Sub TextBox14_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) '单词输入框, 限制输入数字
    With Me
        If KeyAscii > asc(0) And KeyAscii < asc(9) Then '如果输入的是数字就提示
            .Label57.Caption = "禁止输入数字"
            KeyAscii = 0
            .TextBox14.Text = ""
        Else
            .Label57.Caption = "" '当输入的其他内容的时候清空警告信息
        End If
    End With
End Sub

Private Sub CommandButton47_Click() '单词查询 '需要修正 '统一使用金山词霸
    Dim strx As String, strx1 As String, strx2 As String
    Dim strLen As Byte
    Dim i As Byte, k As Byte, j As Byte, bl As Byte, c As Byte
    
    On Error GoTo 100
    With Me
        strx = Trim(.TextBox14.Text) '去除两侧可能潜在的空格
        strLen = Len(strx)
        If strLen = 0 Or strLen > 16 Or strx Like "*#*" Then Exit Sub '输入的内容null or 太长/是包含数字就直接退出 -由于获取网络内容的速度不快,一般的不必要的请求都过滤掉
        bl = UBound(Split(strx, Chr(32))) '获取空格的数量
        If bl > 1 Then
            .TextBox13.Text = Replace(strx, Chr(32), "", 1, 1) '限制空格的数量,并去掉其中的一个空格
            .Label57.Caption = "输入的内容有误"
            Exit Sub
        End If
        
        If TestURL("https://www.baidu.com") Then '测试网络链接
            For i = 1 To strLen '判断输入的内容全部是否为中文或者英文 '中间空格的问题
                strx1 = Mid$(strx, i, 1)
                If strx1 Like "[a-zA-Z]" Then
                    k = k + 1
                ElseIf strx1 Like "[一-龥]" Then
                    j = j + 1
                End If
            Next
            If bl = 1 Then j = j + 1: k = k + 1 '如果存在空格,就+1 '允许英文存在一个空格
            If k = strLen Then
                If k < 2 Or k > 16 Then                 '英文输入,输入的内容过短,长都将过滤掉
                    .Label57.Caption = "输入的内容可能有误"
                    Exit Sub
                End If
                If bl = 1 Then '如果存在空格,就改掉用金山
                    c = 3
                Else
                    c = 1
                End If
            ElseIf j = strLen Then            '只有输入的内容正确执行-中文输入
                If j > 6 Then                                                        '限制输入内容的长度
                    .Label57.Caption = "输入的内容可能有误"
                    Exit Sub
                End If
                If bl = 1 Then strx = Replace(strx, Chr(32), "") '禁止中文存在空格
                c = 2
            Else
                .Label57.Caption = "输入的内容存在非中/非英的混合内容"
                Exit Sub
            End If
        Else
            .Label57.Caption = "网络连接异常"
            Exit Sub
        End If
        strx2 = GetdicMeaning(strx, c)
        If c = 3 Or c = 2 Then strx2 = Replace(strx2, "释义", "释义: ", 1, 1) '对返回的结果进行优化
        If Right$(strx2, 1) = Chr(59) Then strx2 = Left$(strx2, Len(strx2) - 1) 'chr(59)=";",右边最后一个符号去掉
        .TextBox13.Text = strx2
        If c = 1 Or c = 3 Then '勾选发音
            If voicex = 1 Then Speakvs (strx)
        End If
    End With
    Exit Sub
100
    If Err.Number <> 0 Then
    Me.Label57.Caption = "异常: " & Err.Number
    Err.Clear
    End If
End Sub

Private Sub CommandButton53_Click() '获取豆瓣评分
    Dim strx As String
    
    With Me
        If .CommandButton53.Caption = "豆瓣评分获取" Then
            strx = Trim(.TextBox3.Text)
            If Len(strx) < 2 Or .Label56.Caption = "未找到书籍信息" Then Exit Sub '避免非必要的执行.label56暂时存储执行的状态,可以用模块级变量替换掉
            Call DoubanBook(strx)
        Else
            SearchFile (.Label29.Caption) '编辑豆瓣信息
            If Rng Is Nothing Then .Label57.Caption = "文件丢失": Exit Sub
            .TextBox15.Text = Rng.Offset(0, 23).Value '名称
            .TextBox16.Text = Rng.Offset(0, 24).Value '评分
            .TextBox17.Text = Rng.Offset(0, 25).Value '链接
            .TextBox15.Visible = True
            .TextBox16.Visible = True
            .TextBox17.Visible = True
            CommandButton54.Visible = True
            Set Rng = Nothing
        End If
    End With
End Sub

Private Sub CommandButton54_Click() '添加豆瓣数据
    Dim strx As String, strx1 As String, strx2 As String, strx3 As String, strx4 As String, strx5 As String
    Dim arr() As String, strx7 As String, strx6 As String, strx8 As String
    With Me
        strx1 = .TextBox17.Text
        strx8 = .Label29.Caption
        If Len(strx1) = 0 Or Len(strx8) = 0 Then Exit Sub
        If FileStatus(strx8, 2) = 4 Then '检查文件是否存在目录/本地磁盘
            strx = .TextBox16.Text
            Rng.Offset(0, 23) = .TextBox15.Text '名称
            Rng.Offset(0, 24) = strx '
            Rng.Offset(0, 25) = strx1 '书籍链接
            .Label106.Caption = strx1
            .Label69.Caption = strx '评分
            '---------------------------------------------进一步获取豆瓣信息
            If Len(ThisWorkbook.Sheets("temp").Cells(53, "ab").Value) = 0 Then
                If TestURL(BdUrl) = False Then .Label57.Caption = "网络不可用": Exit Sub
                ReDim arr(2)
                arr = ObtainDoubanPicture(strx1) '豆瓣图片链接/作者,国籍
                strx2 = arr(0)
                strx3 = .Label29.Caption
                strx4 = ThisWorkbook.Sheets("temp").Cells(46, "ab").Value
                If Len(strx4) = 0 Then
                    strx5 = ThisWorkbook.Path & "\bookcover"
                    ThisWorkbook.Sheets("temp").Cells(46, "ab").Value = strx5
                    strx4 = strx5
                Else
                    If fso.folderexists(strx4) = False Then
                        strx5 = ThisWorkbook.Path & "\bookcover"
                        fso.CreateFolder (strx5)
                        ThisWorkbook.Sheets("temp").Cells(46, "ab").Value = strx5
                        strx4 = strx5
                    End If
                End If
                strx4 = strx4 & "\" '存储位置
                strx7 = LCase(Right(strx2, 3))
                If strx7 = "jpg" Or strx7 = "png" Then '符合要求
                    strx6 = strx4 & strx3 & "." & strx7
                    If DownloadFilex(strx2, strx6) = True Then
                        Rng.Offset(0, 34) = strx2  '链接
                        Rng.Offset(0, 36) = strx6 '本地文件路径
                        Imgurl = strx6
                        .Frame2.Width = 158
                        .TextBox2.Width = 143
                        .CommandButton134.Left = 610
                        .CommandButton125.Left = 654
                        With .Label239
                            .Visible = True
                            .Left = 728
                            .Top = 94
                            .Caption = "豆瓣封面"
                        End With
                        With .Image1
                            .Left = 708
                            .Top = 108
                            .Width = 84
                            .Height = 122
                            .Visible = True
                            .Picture = LoadPicture(strx6)
                            .PictureSizeMode = fmPictureSizeModeStretch '调整图片
                        End With
                        imgx = 1
                    Else
                        If imgx = 1 Then
                            Rng.Offset(0, 34) = strx2
                            If strx2 <> Rng.Offset(0, 36).Value Then Rng.Offset(0, 36) = ""
                            '-----------------如果图片文件没有下载成功,且原有的图片路径不一致, 而原来有图片那么就清除掉原有的路径
                            With Me
                                .Image1.Visible = False
                                .Label239.Visible = False
                                .Frame2.Width = 246
                                .TextBox2.Width = 231
                                .CommandButton134.Left = 698
                                .CommandButton125.Left = 742
                            End With
                            imgx = 0
                            Imgurl = ""
                        End If
                    End If
                End If
                Rng.Offset(0, 37) = arr(2) '国籍
                If Len(arr(1)) > 0 Then
                    Rng.Offset(0, 14) = arr(1) '作者
                    .TextBox4.Text = arr(2) & arr(1)
                End If
            End If
            Set Rng = Nothing
            .Label57.Caption = "添加成功"
        Else
            .Label57.Caption = "文件丢失"
        End If
    End With
End Sub

Private Sub CommandButton55_Click() '文件md5计算
    Dim p As Integer, strx As String
    
    With Me
        If Len(.Label71.Caption) > 0 Or Len(.Label25.Caption) = 0 Or Me.Label55.Visible = True Then Exit Sub
        If FileStatus(.Label29.Caption, 2) <> 4 Then Set Rng = Nothing: Exit Sub '判断文件是否存在
        If .Label74.Caption = "Y" Then '路径中是否存在非ansi编码字符
            p = 2
        ElseIf .Label74.Caption = "N" Then
            p = 1
        End If
        strx = GetFileHashMD5(.Label25.Caption, p)
        If Len(strx) = 2 Then .Label57.Caption = "未获得有效值": Exit Sub
        .Label71.Caption = strx
        If Len(ThisWorkbook.Sheets("temp").Range("ab29").Value) = 0 Then '如果勾选自动写入md5
            .CommandButton56.Enabled = True
            .Label57.Caption = "操作完成"
        Else
            Call WriteMd5(1)
        End If
'        Warning (1) '操作完成
    End With
    Set Rng = Nothing
End Sub

Function WriteMd5(ByVal xi As Byte) '写入md5
    With Me
        If xi = 0 Then SearchFile (.Label29.Caption)
        If Rng Is Nothing Then .Label57.Caption = "添加失败": Set Rng = Nothing: Exit Function
        Rng.Offset(0, 9) = .Label71.Caption
        Set Rng = Nothing
        Me.Label57.Caption = "添加成功"
'        Me.CommandButton55.Enabled = False
    End With
End Function

Private Sub CommandButton56_Click() '编辑-记录文件hash/不在需要检查,md计算部分已经处理
    Call WriteMd5(0)
End Sub

Private Sub CommandButton57_Click() '工具-简-繁转换
    Dim strx As String, strx1 As String, strLen As Byte, i As Byte, k As Byte
    
    With Me
        strx = Trim(.TextBox18.Text)
        strLen = Len(strx)
        If strLen = 0 Then Exit Sub
        
        If IsNumeric(strx) Then
            .TextBox18.Text = ""
            .Label57.Caption = "不允许输入纯数字"
            Exit Sub ''如果全部是数字就退出
        ElseIf strLen > 30 Then
            .Label57.Caption = "输入的字符串过长,限制30"
            Exit Sub
        End If
        
        For i = 1 To strLen
            strx1 = Mid(strx, i, 1)
            If strx1 Like "[一-龥]" Then k = k + 1 '判断是否存在中文
        Next
        
        If k = 0 Then
            .TextBox18.Text = ""
            .Label57.Caption = "输入的为无效信息"
            Exit Sub
        End If
        
        .TextBox19.Text = SC2TC(strx)
    End With
End Sub

Private Sub CommandButton58_Click() '工具-繁-简转换
    Dim strx As String, strx1 As String, strLen As Byte, i As Byte, k As Byte
    
    With Me
        strx = Trim(.TextBox18.Text)
        strLen = Len(strx)
        If strLen = 0 Then Exit Sub
        
        If IsNumeric(strx) Then
            .TextBox18.Text = ""
            .Label57.Caption = "不允许输入纯数字"
            Exit Sub ''如果全部是数字就退出
        ElseIf strLen > 30 Then
            .Label57.Caption = "输入的字符串过长,限制30"
            Exit Sub
        End If
        
        For i = 1 To strLen
            strx1 = Mid(strx, i, 1)
            If strx1 Like "[一-龥]" Then k = k + 1
        Next
        
        If k = 0 Then
            .TextBox18.Text = ""
            .Label57.Caption = "输入的为无效信息"
            Exit Sub
        End If
        
        .TextBox19.Text = TC2SC(.TextBox18.Text)
    End With
End Sub

Private Sub CommandButton59_Click() '工具-独立文件计算md5-需要修改
    Dim strx As String, strx1 As String

    With Me
        strx = .TextBox20.Text
        If Len(strx) > 4 Then
            If fso.fileexists(strx) = False Then Exit Sub '基本的文件路径长度C:\a '4
            .Label57.Caption = "处理中..."
            strx1 = GetFileHashMD5(strx)
            If Len(strx1) = 2 Then .Label57.Caption = "未获得有效值": Exit Sub
            .TextBox21.Text = UCase(strx1)
            .Label57.Caption = "处理完成"
        End If
    End With
End Sub

Private Sub CommandButton60_Click() '复制md5
    With Me.TextBox21 '可以利用textbox的这个属性作为实现复制信息到剪切板的间接途径
        If Len(.Text) = 0 Then Exit Sub
        .Text = .Value
        .SelStart = 0
        .SelLength = Len(.Text)
        .Copy
    End With
End Sub

Private Sub CommandButton61_Click() '转换字符串-md5-crc32-sha256
    Dim strx As String, strx1 As String, strx2 As String, i As Byte
    
    With Me
        strx = Trim(.TextBox18.Text)
        If Len(strx) = 0 Then Exit Sub
        strx1 = ThisWorkbook.Sheets("temp").Cells(38, "ab").Value
        If Len(strx1) = 0 Then
            i = 0
        Else
            If IsNumeric(strx1) = True Then
                i = Int(strx1)
            Else
                i = 0
            End If
        End If
        
        Select Case i
            Case 1: strx2 = GetMD5Hash_String(strx)
            Case 2: strx2 = CRC32API(strx1)
            Case 3: strx2 = SHA256Function(strx)
            Case Else
                strx2 = GetMD5Hash_String(strx)
        End Select
        .TextBox19.Text = strx2
    End With
End Sub

Private Sub CommandButton62_Click()
    With Me.TextBox19 '可以利用textbox的这个属性作为实现复制信息到剪切板的间接途径
        If Len(.Text) = 0 Then Exit Sub
        .Text = .Value
        .SelStart = 0
        .SelLength = Len(.Text)
        .Copy
    End With
End Sub

Private Sub CommandButton63_Click() '文件夹-折叠所有的节点
    Dim i As Integer, k As Integer
    With Me.TreeView1
        k = .Nodes.Count
        If k = 0 Then Exit Sub
        For i = 2 To k
            .Nodes(i).Expanded = False
        Next
    End With
    With Me
        If Len(.Label240.Caption) > 0 Then
            .Label240.Caption = ""
            .Label241.Caption = ""
        End If
    End With
End Sub

Private Sub CommandButton64_Click() '文件夹-展开所有的节点
    Dim i As Integer, k As Integer
    
    With Me.TreeView1
        k = .Nodes.Count
        If k = 0 Then Exit Sub
        For i = 1 To k
            .Nodes(i).Expanded = True
        Next
    End With
End Sub

Private Sub TextBox28_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With Me.TextBox28 '可以利用textbox的这个属性作为实现复制信息到剪切板的间接途径
        If Len(.Text) = 0 Then Exit Sub
        .Text = .Value
        .SelStart = 0
        .SelLength = Len(.Text)
        .Copy
    End With
    Me.Label57.Caption = "密码已复制"
End Sub

Private Sub TextBox29_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '工具-文件解压
    Dim fdx As FileDialog, strfolder As String
    Dim selectfile As Variant

    If Me.CheckBox14.Value = True Then
    With ThisWorkbook.Application.FileDialog(msoFileDialogFolderPicker) '文件夹选择窗口(文件夹以此只能选一个)
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub '未选择文件夹则退出sub
        strfolder = .SelectedItems(1)
        folderpathc = strfolder
        If ErrCode(folderpathc, 1) > 1 Then MsgShow "文件路径包含非ansi编码,请勿手动修改内容框的信息", "Tips", 1800
        Me.TextBox29 = folderpathc
        Me.TextBox29.SetFocus
        Exit Sub
    End With
    End If
    Set fdx = ThisWorkbook.Application.FileDialog(msoFileDialogFilePicker)
    With fdx
        .AllowMultiSelect = False '不允许选择多个文件(注意不是文件夹,文件夹只能选一个)
        .Show
        .Filters.Clear                                  '清除现有规则
        .Filters.Add "Zip", "*.7z; *.zip; *.rar", 1     '筛选文件
        .Filters.Add "All File", "*.*", 1
        If .SelectedItems.Count = 0 Then Exit Sub
        filepathc = .SelectedItems(1)
        If ErrCode(filepathc, 1) > 1 Then MsgShow "文件路径包含非ansi编码,请勿手动修改内容框的信息", "Tips", 1800
        Me.TextBox29 = filepathc
    End With
    Me.TextBox29.SetFocus
    Set fdx = Nothing
End Sub

Private Sub TextBox3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With Me.TextBox3 '可以利用textbox的这个属性作为实现复制信息到剪切板的间接途径
        If Len(.Text) = 0 Then Exit Sub
        .Text = .Value
        .SelStart = 0
        .SelLength = Len(.Text)
        .Copy
    End With
End Sub

Private Sub TreeView1_DblClick() '树状图节点双击展开
    With Me
        .TreeView1.Nodes(.TreeView1.SelectedItem.Index).Expanded = True
    End With
End Sub

Private Sub CommandButton65_Click() '更新文件夹
    Dim rnglistx As Range
    Dim filexpath As String
    Dim strx As String, strx1 As String
    Dim k As Integer, addcodex As Byte, j As Integer, i As Integer
    
    ReDim arrch(1 To 50)
    With Me
        If .TreeView1.Nodes.Count = 0 Then Exit Sub
        k = .TreeView1.SelectedItem.Index
        If k = 1 Then Exit Sub
        strx = .TreeView1.Nodes(k).Text           '-禁止无法更新的文件
        If InStr(strx, "无法更新") > 0 Then .Label57.Caption = "此文件受限": Exit Sub
        If .CheckBox2.Value = True Then '勾选
            addcodex = 1
        Else
            addcodex = 2
        End If
        filexpath = .TreeView1.Nodes(k).key
        If ListAllFiles(addcodex, filexpath) = False Then .Label57.Caption = "受限文件夹": Exit Sub
        If InStr(strx, "发生变化") > 0 Then
            .TreeView1.Nodes(k).Text = Left$(strx, Len(strx) - 6) '去掉发生变化的提示
        ElseIf InStr(strx, "已添加") = 0 Then
            .TreeView1.Nodes(k).Text = strx & "(已添加)" '变更标题
        End If
        If .TreeView1.Nodes(k).Children > 0 Then '包含有子文件夹
            If addcodex = 1 Then
                Call CheckTreeLists(.TreeView1, .TreeView1.Nodes(k)) '检查子项
                j = ich - 1
                For i = 1 To j                                '数组的最后一个值就是选定项
                    strx1 = .TreeView1.Nodes(arrch(i)).Text
                    If InStr(strx1, "发生变化") > 0 Then
                        .TreeView1.Nodes(arrch(i)).Text = Left$(strx1, Len(.TreeView1.Nodes(arrch(i)).Text) - 6) '去掉标题中的(发生变化)'.TreeView1.Nodes(arrch(arrch(i))).text
                    ElseIf InStr(strx, "已添加") = 0 Then
                        .TreeView1.Nodes(arrch(i)).Text = strx & "(已添加)" '变更标题,增加已添加
                    End If
                Next
            End If
        End If
    End With
    DataUpdate '整体数据更新
End Sub

Private Sub CommandButton66_Click() '文件夹-更新树状图
    With Me
        If Len(.Label80.Caption) = 1 Then
            .ListView2.ListItems.Clear
            With .TreeView1
                If Len(ThisWorkbook.Sheets("主界面").Range("e37").Value) = 0 Then '当数据被清空的时候
                    .Nodes.Clear
                Else
                    .Nodes.Clear
                    .Nodes.Add , , "Menus", "Menus" '根目录
                    .Nodes(1).Expanded = True
                    .Appearance = cc3D
                    .HotTracking = True
                    .Nodes(1).Bold = True
                    .LabelEdit = tvwManual '设置节点不可编辑
                    Call TreeLists(1)
                End If
            End With
            If Len(.Label240.Caption) > 0 Then
                .Label240.Caption = ""
                .Label241.Caption = ""
            End If
        End If
    End With
End Sub

Function TreeLists(ByVal exc As Byte) '树状展开
    Dim dic As New Dictionary '存储去重的主文件夹
    Dim arr() As String
    Dim arrx() As String
    Dim i As Integer, l As Integer, k As Integer, Elow As Integer
    Dim strx As String, strx1 As String
    
    strx = ThisWorkbook.Sheets("主界面").Range("e37").Value
    If Len(strx) = 0 Or InStr(strx, "\") = 0 Then Exit Function '不存在文件夹
    With Me
        strx1 = .Label80.Caption
        If exc = 0 And strx1 = "temp" Then '用于控制
            With .TreeView1
                .Appearance = cc3D
                .HotTracking = True
                .Nodes.Add , , "Menus", "Menus" '根目录
                .Nodes(1).Expanded = True '第一个节点展开的状态
                .Nodes(1).Bold = True
                .LabelEdit = tvwManual '设置节点不可编辑
            End With
        End If
        
        If strx1 = "temp" Or exc = 1 Then
            With ThisWorkbook.Sheets("主界面")
                Elow = .[e65536].End(xlUp).Row
                ReDim arr(1 To Elow - 36)
                For i = 37 To Elow
                    arr(i - 36) = Split(.Range("e" & i), "\")(0) & "\" & Split(.Range("e" & i), "\")(1) '最上层的目录
                Next
            End With
        
            For k = 1 To UBound(arr)
                If fso.folderexists(arr(k)) = False Then GoTo 10 '校检文件夹是否存在
                dic(arr(k)) = ""
10
            Next
            ReDim arrx(0 To UBound(dic.Keys))
            For l = 0 To UBound(dic.Keys)
                arrx(l) = dic.Keys(l)
            Next
            ListFolderx arrx
            If exc = 1 Then Exit Function '控制
            .Label80.Caption = 1 '标记
        End If
    End With
End Function

Function ListFolderx(ByRef arrt() As String) '树状展开 '数组必须是byref
    Dim fd As Folder
    Dim i As Integer
    Dim showname As String
    Dim rnglistx As Range
    Dim strp As String
    Dim tracenum As Byte, k As Byte, blow As Integer, j As Byte
    
    k = UBound(arrt())
    ReDim arrlx(1 To 50)
    With ThisWorkbook.Sheets("目录")
        blow = .[b65536].End(xlUp).Row
        j = .Cells.SpecialCells(xlCellTypeLastCell).Column
        For i = 0 To k
            s = 1                         '注意这里重新归1处理 s为模块级变量,用完进行重置
            tracenum = 0
            Set fd = fso.GetFolder(arrt(i))
            If fd.ParentFolder.Path = Environ("SYSTEMDRIVE") & "\" Then tracenum = 1 '位于系统盘所在的磁盘
            arrlx(s) = fd.Path
            showname = fd.Name
            If tracenum = 1 Then showname = showname & "(无法更新)"
            If tracenum <> 1 Then
                strp = fd.Path & "\"
                Set rnglistx = .Cells(4, 3).Resize(blow, j).Find(strp, lookat:=xlWhole)
                If Not rnglistx Is Nothing Then
                    If fd.DateLastModified <> rnglistx.Offset(0, 2) Then '文件夹的修改时间发生变化(意味着文件夹(不包含已有的子文件夹)的一层发生变化,修改/删除/新建文件夹/修改文件夹)
                        showname = showname & "(已添加)(发生变化)"
                    Else
                        showname = showname & "(已添加)"
                    End If
                End If
            End If
            
            With Me.TreeView1.Nodes
                .Add "Menus", 4, arrlx(s), showname
            End With
            ListFolderxs fd, tracenum
        Next
    End With
    Erase arrlx
    Set fd = Nothing
End Function

Private Function ListFolderxs(ByVal fd As Folder, ByVal contx As Byte) '文件夹-显示列表
    Dim sfd As Folder, i As Long
    Dim showname As String
    Dim rnglistx As Range
    Dim strp As String, strx As String
    
    On Error GoTo 110
10
    If fd.SubFolders.Count = 0 Then Exit Function '子文件夹数目为零则退出sub
    For Each sfd In fd.SubFolders
       strx = sfd.Path
       If ErrCode(strx, 1) > 1 Then GoTo 100 '限制包含非ansi字符的文件夹 , 如果需要需要处理非ansi字符的路径,需要建临时的数组用于保存路径
       If contx = 1 Then
          If strx <> Environ("UserProfile") Then GoTo 100    '只允许添加用户文件夹
       End If
       If contx = 2 Then
'           If sfd.Path <> Environ("UserProfile") & "\Downloads" And sfd.Path <> Environ("UserProfile") & "\Documents" And Environ("UserProfile") & "Desktop" Then GoTo 100 '只允许添加用户文件夹下的download和document,desktop三个文件夹
            If CheckFileFrom(strx, 2) = True Then GoTo 100
       End If
20
       i = sfd.Attributes
       If i = 18 Or i = 1046 Then GoTo 100 '这里需要注意系统文件夹的问题,拒接访问 ,这里18表示隐藏属性,有别于文件的34隐藏属性\1046,特殊文件类型
       '-------------------------或者可以进一步限制文件的属性为17
       showname = sfd.Name
       If contx = 1 Then showname = showname & "(无法更新)"
       If contx <> 1 Then
          With ThisWorkbook.Sheets("目录")
             strp = strx & "\"
             Set rnglistx = .Cells(4, 3).Resize(.[b65536].End(xlUp).Row, .Cells.SpecialCells(xlCellTypeLastCell).Column).Find(strp, lookat:=xlWhole)
                 If Not rnglistx Is Nothing Then
                     If sfd.DateLastModified <> rnglistx.Offset(0, 2) Then '文件夹的修改时间发生变化(意味着文件夹(不包含已有的子文件夹)的一层发生变化,修改/删除/新建文件夹/修改文件夹)
                         showname = showname & "(已添加)(发生变化)"
                         
                     Else
                         showname = showname & "(已添加)"
                     End If
                 End If
           End With
        End If
        With Me.TreeView1.Nodes
            If arrlx(s + 1) <> strx Then
                arrlx(s + 1) = strx
                .Add arrlx(s), 4, arrlx(s + 1), showname
            End If
        End With
30
        If sfd.SubFolders.Count > 0 Then s = s + 1
        If contx > 0 Then contx = 2 '执行循环下一阶段
        ListFolderxs sfd, contx
100
    Next
    s = s - 1 '重算
    Exit Function
110
If Err.Number = 70 Then Err.Clear: GoTo 100 '某些系统文件夹会出现拒绝访问权限70错误
End Function

Function CheckTreeLists(ByRef treevw As TreeView, ByRef nodThis As node) '文件夹-显示列表子节点,勾选
    Dim lngIndex As Integer

    If nodThis.Children > 0 Then
        lngIndex = nodThis.Child.Index
        Call CheckTreeLists(treevw, treevw.Nodes(lngIndex))
        While lngIndex <> nodThis.Child.LastSibling.Index
          lngIndex = treevw.Nodes(lngIndex).Next.Index
          Call CheckTreeLists(treevw, treevw.Nodes(lngIndex))
        Wend
    End If
    ich = ich + 1
    arrch(ich) = nodThis.Index
End Function

Private Sub TreeView1_Click() '树状图节点选定
    Dim i As Integer, k As Integer, m As Integer
    Dim fd As Folder, strx3 As String, strx4 As String, fdz As String, fdzx As Long
    Dim strx As String, strx1 As String, strx2 As String
    
    On Error Resume Next
    With Me.TreeView1
        k = .Nodes.Count
        If k = 0 Then Exit Sub
        m = .SelectedItem.Index
        strx = .Nodes(m).key
        strx3 = .Nodes(m).Text
        .SelectedItem.Bold = True
        For i = 2 To k
            If i = m Then GoTo 100
            .Nodes(i).Bold = False '非选中节点,字体的粗体取消
100
        Next
    End With
    If m = 1 Then Exit Sub
    If InStr(strx3, "无法更新") > 0 Then Exit Sub
    If Len(storagex) > 0 Then
        If strx = storagex Then Exit Sub
    Else
        storagex = strx '存储值,避免同样的内容反复被加载
    End If
    Set fd = fso.GetFolder(strx)
    strx4 = "Size: "
    '---------------------当访问某些特殊文件夹的时候会出现错误(无权限)
    fdz = CStr(fd.Size) ' 转换为文本   http://www.360doc.com/content/18/0613/16/36550511_762117327.shtml
    If Len(fdz) = 0 Then
        fdz = "未获取到数据"
    Else
        fdzx = Int(fdz) '转换为数字
        If fdzx < 1048576 Then
             fdz = Format(fdzx / 1024, "0.00") & "KB"    '文件字节大于1048576显示"MB",否则显示"KB"
        Else
            fdz = Format(fdzx / 1048576, "0.00") & "MB"
        End If
    End If
    With Me
        .ListView2.ListItems.Clear
        .Label240.Caption = strx4 & fdz
        .Label241.Caption = fd.DateLastModified
    End With
    Set fd = Nothing
    If FileFilterTemp(strx, 5, , , , , 2) = 0 Then Me.Label57.Caption = "此文件夹为空"
'   If strx = .Label81.Caption Then Exit Sub '已经写入数据 '控制
'   Call FileListv(strx, strx1, strx2, 5, 0, 0, 0)
End Sub

Private Sub CommandButton68_Click() '文件夹-显示列表
    Call TreeLists(0)
End Sub

Private Sub CommandButton69_Click() '随机推荐-文件夹
    Dim listnumt As Integer, listnumx As Integer, litm As Variant, randomx As Integer, i As Integer
    
    With ThisWorkbook.Sheets("书库")
        listnumt = .[b65536].End(xlUp).Row - 5
        If listnumt < 10 Then
            Me.Label57.Caption = "内容太少不推荐"
            Exit Sub '数量太少不显示
        End If
        Me.ListView2.ListItems.Clear
        If listnumt < 50 Then '显示推荐的数量
            listnumx = 5
        ElseIf listnumt > 49 And listnumt < 150 Then
            listnumx = 10
        ElseIf listnumt > 149 Then
            listnumx = 15
        End If
        ReDim arrTemp(1 To listnumx)
        For listnum = 1 To listnumx             '参数部分共享单词部分的参数arrtemp,listnum为模块级别变量
            randomx = RandomNumx(listnumt)
            arrTemp(listnum) = randomx
            Set litm = Me.ListView2.ListItems.Add()
            i = randomx + 5
            litm.Text = .Range("b" & i)
            litm.SubItems(1) = .Range("c" & i)
            litm.SubItems(2) = .Range("d" & i)
            litm.SubItems(3) = .Range("e" & i)
        Next
    End With
End Sub

Private Sub CommandButton76_Click() '工具-算法1
    Call Matchx1
    Me.CommandButton76.Enabled = False
    Me.CommandButton77.Enabled = True
End Sub

Private Sub CommandButton77_Click() '工具-算法2
    Call Matchx2
    MsgBox "Speed Defines The Winner", vbCritical, "Tips"
    Me.CommandButton76.Enabled = True
    Me.CommandButton77.Enabled = False
End Sub

'---------------------------------------------------------------------------单词训练
Private Sub Frame11_Click() '单词训练框架
    Me.TextBox23.SetFocus '点击就聚焦
End Sub

Private Sub CommandButton71_Click() '单词训练-开始
    Timeset = 2 '让另一个时间sub处于停止的状态
    With Me
        If RecData = False Then .Label57.Caption = "连接异常": Exit Sub '检查本地文件连接是否正常
        If Len(.ComboBox10.Value) = 0 Or IsNumeric(.ComboBox10.Value) = False Then Exit Sub
        If .ComboBox10.Value < 5 Or .ComboBox10.Value > 30 Then Exit Sub '防止意外手动输入造成的错误
        If .Label66.Caption = "play" Then '音乐控件在运行的状态
            wm.Controls.Stop
            .Label66.Caption = "stop"
        End If
        .Frame1.Enabled = False
        .Frame3.Enabled = False
        .Frame10.Enabled = False
        .CommandButton71.Enabled = False '执行后禁用
        .TextBox23.SetFocus
        .ComboBox10.Enabled = False
    End With
    Call Excesub
End Sub

Sub Excesub() '单词训练主程序
Dim sTest As String
Dim i As Byte, Alastrow As Integer, randomx As Integer, chlistnum As Byte, dicl As Byte
Dim time1 As Long, timex As Integer '记录时间
Dim spx As Integer, sp As Byte, spx1 As Byte, pausetime As Long, tipslen As Byte, tipsx As Byte, errc As Byte
Dim lastpath As String, TableName As String
Dim strx As String

Alastrow = 3041 '单词表的数量
listnum = 0

With Me
    chlistnum = Int(.ComboBox10.Value) '取整数,防止手动输入错误
    .Label86.Caption = "题数:" & chlistnum '可修改
    .Label88.Caption = "状态:进行中" '显示状态
    .Label94.Caption = "" '清空正确率
    .ListView3.ListItems.Clear '使用之前清空列表数据
    
    ReDim arrTemp(1 To chlistnum)
    ReDim arrtemp1(1 To chlistnum) '重置数组 '重置数组也能将原有的数组的内容清空
    ReDim arrtemp2(1 To chlistnum)
    ReDim arrtemp3(1 To chlistnum)
    TableName = "单词"
    Set rs = New ADODB.Recordset    '创建记录集对象    '或者把所有的单词提取出来放到数组里
    If rs Is Nothing Then MsgBox "无法创建对象", vbCritical, "Waring"
    For listnum = 1 To chlistnum
        spx = 1 '显示时间 '注意要重置参数
        spx1 = 1 '3s播放
        errc = 0
        pausetime = 0
        FlagStop = False
        Flagpause = False
        Flagnext = False '每次执行前重置参数
200
        randomx = RandomNumx(Alastrow) '这里参数表示生成的随机数的最大值
        arrTemp(listnum) = randomx '临时存储数组 '用于比较新生成的随机数是否出现重叠
        
        SQL = "select * from [" & TableName & "$] where 编号 = " & randomx '从存储文件读取信息
        rs.Open SQL, Conn, adOpenKeyset, adLockOptimistic
        
        If rs.BOF And rs.EOF Then
            errc = errc + 1       '假如没有获取的单词,重新进行获取,尝试5次就退出
            rs.Close
            If errc < 5 Then GoTo 200
            .Label57.Caption = "异常,退出进程"
            Set rs = Nothing
            Exit Sub
        End If
        
        sTest = rs(1) '存储试题 '答案-英文
        arrtemp3(listnum) = rs(2) '中文
        arrtemp1(listnum) = sTest
        .Label90.Caption = rs(2)
        .Label97.Caption = rs(3)
        rs.Close
        
        dicl = Len(sTest)
        tipsx = 0
        If .CheckBox6.Value = True Then tipsx = 1
        If dicl < 6 Then                        '根据单词的长度来决定时间'当然也可以修改时间的其他权重
            timex = 16000
            If tipsx = 1 Then .Label102.Caption = Left$(sTest, 2) '如果勾选提示
        ElseIf dicl > 5 And dicl < 8 Then
            timex = 19000
            If tipsx = 1 Then .Label102.Caption = Left$(sTest, 2)
        ElseIf dicl > 7 Then
            timex = 24000
            tipslen = Int(dicl * 0.4) '根据长度来产生提示
            If tipsx = 1 Then .Label102.Caption = Left$(sTest, tipslen)
        End If
        .Label84.Caption = "总时间:" & timex / 1000 & "s" '作答时间
        Sleep 100                                    '延迟执行
        time1 = timeGetTime '开始的时间
        .Label89.Caption = "已作答:" & listnum '作答数量
        If .CheckBox4.Value = True Then
            Speakvs (sTest) '发音
            timex = timex + 100 '时间补偿,由于发音造成的时间延迟会导致最后一秒的显示异常(如果直接使用application.speech,延时就会非常明显)
        End If
        
        Do
            DoEvents
            If Flagpause Then
                GoSub 1
                time1 = timeGetTime '重新的时间开始重新计算
                pausetime = sp * 1000 '重新开始的时间
            End If
            sp = Int((timeGetTime - time1 + pausetime) / 1000)
            '小心这里的运算 假如i as integer,k as integer 在进行相加,相乘等操作时,如i=100,k=400 msgbox i*k会造成数据范围的溢出,需要使用cnlg或者定义数据类型为long类型
            If sp = spx Then
                .Label85.Caption = "耗时:" & spx & "s" '计时
                spx = spx + 1
            End If
            If sp / 5 = spx1 Then '每隔5s播放一次
            spx1 = spx1 + 1 '循环三次
                If spx1 < 4 Then
                    timex = timex + 100 '注意补充的时间不要超过1000ms(1s)
                    If .CheckBox4.Value = True And .CheckBox5 = True Then Speakvs (sTest)
                    'ThisWorkbook.Application.Speech.Speak (sTest) '注意这里的application要加thisworkbook,否侧在激活界面发生切换的时候会出现错误
                    DoEvents
                End If
            End If
            If Flagnext Then GoTo 1000 '注意执行的先后顺序
            If FlagStop Then
                Set rs = Nothing
                Conn.Close
                Set Conn = Nothing
                Call QuestionOver
                .Label94.Caption = "作答未完成"
                Exit Sub
            End If
            Sleep 25 '必须控制循环的间歇,防止cpu持续运转
        Loop While timeGetTime - time1 + pausetime < timex
1000
        .Label85.Caption = "耗时:"
        .Label102.Caption = ""
        strx = Trim(.TextBox23.Text) '防止存在空格
        If Len(strx) > 0 Then  '自动获取输入框的内容
            arrtemp2(listnum) = strx '作答结果
        Else
            arrtemp2(listnum) = "No ANSWER" '空值或者空格
        End If
        .TextBox23.Text = "" '清空作答区域
    Next '注意这里再next之后listnum发生变化

End With

Set rs = Nothing
Call Answerx '作答的结果
Call QuestionOver '答题完成后的处理

Exit Sub '在调用工程级变量或者模块级别的变量时,需要跳出执行并释放变量的不是采用exit sub 而是end(结束,并且完全释放内存)'end 可以结束所有的进程,意味着可以在子sub中结束父sub,注意不能在窗口中使用
1:
    For i = 0 To 1 Step 0 '如果鼠标箭头出现转圈圈的情况可以考虑使用自调用的application,now来替换循环的方法
        Sleep 50
        DoEvents
        If Flagpause = False Then Return '继续执行
        If FlagStop Then
            Set rs = Nothing
            Call QuestionOver
            Me.Label94.Caption = "作答未完成"
            Exit Sub
        End If
    Next
End Sub

Function Speakvs(ByVal strx As String) 'vbs方式执行application.speech,如果直接执行,代码运行会在这个动作中造成约1s延迟无法在textbox中输入内容,vbs的方式在某种程度实现所谓的多线程
    Dim WshShell As Object
    Dim vbfilecm As String
    
    If vbsx = 0 Then '控制
    With ThisWorkbook.Sheets("temp")
        vbfilecm = .Range("ab17").Value
        If Len(vbfilecm) = 0 Or fso.fileexists(vbfilecm) = False Then
            Me.Label57.Caption = "配置文件丢失"
            Exit Function
        End If
        vbfilex = vbfilecm            '第一次使用就调用检查
        vbsx = 1
    End With
    End If
    vbfilecm = vbfilex & """" & strx & """"
    Set WshShell = CreateObject("Wscript.Shell")
    WshShell.Run """" & vbfilecm & ""
    Set WshShell = Nothing
End Function

Function RandomNumx(ByVal randomnum As Integer) As Integer '随机数
    Dim RndNumber, i As Byte
    
    Randomize (Timer) '初始化rnd
100
    RandomNumx = Int(randomnum * Rnd) + 1
    If listnum > 1 Then
        For i = 1 To listnum
            If RandomNumx = arrTemp(i) Then GoTo 100 '出现重复的就重新执行
        Next
    End If
End Function

Function Answerx() '单词训练-测试结果
    Dim i As Byte, k As Byte, j As Byte
    Dim arrtemp4() As String '存储判断结果
    Dim litm As Variant
    
    With Me
        j = .ComboBox10.Value
        ReDim arrtemp4(1 To j)
        .Label94.Caption = ""
        For i = 1 To j
            If arrtemp1(i) = arrtemp2(i) Then
                arrtemp4(i) = "Y"
                k = k + 1 '正确的数量
            Else
                arrtemp4(i) = "N"
            End If
            Set litm = .ListView3.ListItems.Add()
            litm.Text = arrtemp3(i)
            litm.SubItems(1) = arrtemp1(i)
            litm.SubItems(2) = arrtemp2(i)
            litm.SubItems(3) = arrtemp4(i)
        Next
        If k = 0 Then
            .Label94.Caption = "Fail"
            Exit Function
        End If
        .Label94.Caption = Int(k * 10) & "%" 'int函数表示取整数,向下取,如12.5,就取12 而不是四舍五入取13
    End With
    Set litm = Nothing
End Function

Function QuestionOver() '单词训练-作答完毕后的处理
    With Me
        .Frame1.Enabled = True
        .Frame3.Enabled = True
        .Frame10.Enabled = True
        .Label88.Caption = "状态:结束"
        .CommandButton71.Enabled = True
        .ComboBox10.Enabled = True
        .Label85.Caption = "耗时:"
        .Label86.Caption = "题数:"
        .Label89.Caption = "已作答:"
        .Label84.Caption = "总时间:"
        .Label90.Caption = ""   '试题
        .Label97.Caption = "" '提示
        .Label102.Caption = ""
    End With
    Erase arrTemp '清空数组
    Erase arrtemp1
    Erase arrtemp2
    Erase arrtemp3
End Function

Private Sub CommandButton72_Click() '单词训练-暂停/继续
    With Me
        If .Label88.Caption = "状态:结束" Or .Label88.Caption = "状态:" Then Exit Sub
        If .CommandButton72.Caption = "暂停" Then
            .CommandButton72.Caption = "继续"
            Flagpause = True
        Else
            Flagpause = False
            .CommandButton72.Caption = "暂停"
        End If
        .TextBox23.SetFocus
    End With
End Sub

Private Sub CommandButton73_Click() '单词训练-停止
    With Me
        If .Label88.Caption = "状态:结束" Or .Label88.Caption = "状态:" Then Exit Sub
        If .CommandButton73.Caption = "停止" Then FlagStop = True
        .TextBox23.SetFocus
        Call QuestionOver
        .Label94.Caption = "作答未完成"
    End With
End Sub

Private Sub CommandButton74_Click() '单词训练-下一个
    With Me
        If .Label88.Caption = "状态:" Or .Label88.Caption = "状态:结束" Then Exit Sub
        Flagnext = True
        .TextBox23.SetFocus
    End With
End Sub
'----------------------------------------------------------------------------------单词训练

Private Sub CommandButton8_Click() '编辑-添加信息
    Dim strx As String, strx1 As String, timea As Date
    Dim str As String, TableName As String, str1 As String, str2 As String, str4 As String, str3 As String
    
    With Me
        strx = Trim(.TextBox3.Text)          '先执行字符判断   '涉及到修改文件路径的操作都要进行字符的判定
        strx1 = Trim(.TextBox4.Text)
        If Len(strx) > 0 Then  '有内容 '主文件名
            If ErrCode(strx, 1) > 1 Then
                .Label57.Caption = "主文件名存在非ANSI字符"
                Exit Sub
            End If
        End If
        If Len(strx1) > 0 Then '作者
            If ErrCode(strx1, 1) > 1 Then
                .Label57.Caption = "作者名称存在非ANSI字符"
                Exit Sub
            End If
        End If
        
        TableName = "摘要记录"
        str = .Label29.Caption '统一编码
        str1 = .Label23.Caption '文件名
        str2 = strx
        str3 = .Label33.Caption '标识
        timea = Now '时间
        str4 = .TextBox2.Text '内容
        If Len(str4) > 1024 Then MsgBox "内容超出长度范围1024", vbInformation, "Tips": Exit Sub
        If RecData = True Then
            SQL = "select * from [" & TableName & "$] where 统一编码='" & str & "'"                                          '查询数据
            Set rs = New ADODB.Recordset
            rs.Open SQL, Conn, adOpenKeyset, adLockOptimistic
            If rs.BOF And rs.EOF Then '用于判断有无找到数据
                SQL = "Insert into [" & TableName & "$] (统一编码,文件名,主文件名,标识编码,时间,内容) Values ('" & str & "','" & str1 & "', '" & str2 & "','" & str3 & "',#" & timea & "#,'" & str4 & "')"
            Else
                SQL = "UPDATE [" & TableName & "$] SET 内容='" & str4 & "',时间=#" & timea & "# WHERE 统一编码='" & str & "'"
            End If
            rs.Close
            Conn.Execute (SQL)
            Call SearchFile(str)
            If Rng Is Nothing Then .Label57.Caption = "文件不存在": Exit Sub
            Rng.Offset(0, 19) = .TextBox5.Text '标签1
            Rng.Offset(0, 20) = .TextBox6.Text '标签2
            If Text3Ch <> Trim(.TextBox13.Text) Then Rng.Offset(0, 13) = strx '经过编辑才写入
            If Text4Ch <> Trim(.TextBox4.Text) Then Rng.Offset(0, 14) = strx1 '作者/主文件名
            Rng.Offset(0, 16) = .ComboBox4.Text '文本质量
            Rng.Offset(0, 17) = .ComboBox5.Text '内容评分
            Rng.Offset(0, 18) = .ComboBox2.Text '推荐指数
            Rng.Offset(0, 31) = .ComboBox12.Text '文字类型
            .Label57.Caption = "操作成功"
        Else
            .Label57.Caption = "失败,没有连接存储文件"
        End If
    End With
    Set Rng = Nothing
    Set rs = Nothing
End Sub

Private Sub CommandButton9_Click()          '工具-笔记工具-创建word笔记
    Dim wdapp As Object
    Dim filen As String, strx As String, filex As String, strx1 As String, strx2 As String
    
    On Error GoTo ErrHandle
    Set wdapp = CreateObject("Word.Application")
    filex = Me.Label29.Caption
    If Len(filex) = 0 Then '如果是空,则打开空白word
        With wdapp
            .documents.Add
            .Visible = True
            .Activate
        End With
        Set wdapp = Nothing
        Exit Sub
    End If
    
    With ThisWorkbook.Sheets("temp")
        strx = .Range("ab30").Value
        strx1 = Me.TextBox3.Text
'        strx2 = Format(Now, "yyyymmddhhmmss")
        If Len(strx) = 0 Then  '创建文件夹
           strx = ThisWorkbook.Path & "\note" '在目录下创建笔记文件夹
           If fso.folderexists(strx) = False Then
              fso.CreateFolder (strx)
              .Range("ab30") = strx
           End If
        Else
           If fso.folderexists(strx) = False Then
              strx = ThisWorkbook.Path & "\note"
              fso.CreateFolder (strx)
              .Range("ab30") = strx
           End If
        End If
    End With
'    filen = strx & "\" & filex & "-" & strx2 & ".docx"   '编号-时间
     filen = strx & "\" & filex & ".docx"   '编号
    With wdapp
        If fso.fileexists(filen) = False Then
            .documents.Add
            .ActiveDocument.Paragraphs(1).Range.InsertBefore (strx1) '在文件的第一行插入标题
            .ActiveDocument.SaveAs FileName:=filen
         End If
        .documents.Open (filen)
        .Visible = True
        .Activate                                   '打开之后显示可见/为当前的活动窗口
    End With
    
ErrHandle:                           '出现错误的时候退出word,否则可能出现word在进程没有被退出的问题
    If Err.Number <> 0 Then
        wdapp.Quit
        Me.Label57.Caption = Err.Number
        Err.Clear
    End If
    Set wdapp = Nothing
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '最近阅读-双击打开最近阅读文件
    Dim i As Integer, k As Byte
    Dim strx As String
    
    With Me.ListBox1
        If .ListCount = 0 Then Exit Sub
        k = .ListIndex
        strx = .Column(0, k)
        If CheckFileOpen(strx) = True Then Set Rng = Nothing: Exit Sub '判断文件是否存在于目录或者本地磁盘/文件是否处于打开的状态
    End With
    With Rng
    If OpenFile(strx, .Offset(0, 1).Value, .Offset(0, 2).Value, .Offset(0, 3).Value, 1, .Offset(0, 27).Value) = False Then
        Me.Label57.Caption = "异常"
        Set Rng = Nothing
        Exit Sub
    End If
    End With
    Call OpenFileOver(strx)
End Sub

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '优先阅读-双击优先列表,打开文件
    Dim n As Byte, m As Byte
    Dim strx As String, strx1 As String
    
    With Me.ListBox2 '优先阅读列表
        m = .ListCount
        If m = 0 Then Exit Sub
        n = .ListIndex
        strx = .Column(0, n)
        strx1 = .Column(1, n)
        If CheckFileOpen(strx) = True Then Set Rng = Nothing: Exit Sub '判断文件是否存在于目录或者本地磁盘/文件是否处于打开的状态
    End With
    With Rng
        If OpenFile(strx, .Offset(0, 1).Value, .Offset(0, 2).Value, .Offset(0, 3).Value, 1, .Offset(0, 27).Value) = False Then
            Me.Label57.Caption = "异常"
            Set Rng = Nothing
            Exit Sub
        End If
    End With
    Call OpenFileOver(strx)
End Sub

Private Function OpenFileOver(ByVal filecodex As String, Optional cx As Byte) '在窗体中打开文件涉及的善后操作-更新窗体的数据
    Dim itemf As ListItem '查找listview1的值
    Dim p As Integer, i As Byte
    
    Call RecentUpdate '重新获取表格的数据-最近打开
    If cx = 1 Then GoTo 100
    With Me.ListView1 '--------------------------------更新搜索框的内容
        If .ListItems.Count = 0 Then GoTo 100
        Set itemf = .FindItem(filecodex, lvwText, , lvwPartial) '查询 -更新搜索框中的信息
        If itemf Is Nothing Then
        Set itemf = Nothing
        Else
        p = itemf.Index
        With .ListItems(p) '主界面搜索框
            If .SubItems(4) = "" Then .SubItems(4) = 0 '空值和"0"还是有区别的
            .SubItems(4) = .SubItems(4) + 1 '打开次数+1
            i = Int(.SubItems(4))
        End With
        End If
    End With
100   '----------------------------------更新编辑区数据
    With Me
        If Len(.Label29.Caption) > 0 Then
            If .Label29.Caption = filecodex Then '更新编辑页面上的信息
                If Reditx = 1 Then
                    Call FileChange '大面积更新数据
                Else
                    If Len(.Label32.Caption) > 0 Then
                        i = Int(.Label32.Caption) + 1
                    Else
                        i = 1
                    End If
                    .Label32.Caption = i
                    .Label31.Caption = Recentfile '更新打开的时间 '打开次数+1
                End If
            End If
        End If
    End With
    Reditx = 0 '传到完后参数重置
    Set Rng = Nothing
    Set itemf = Nothing
End Function

Private Sub ListBox3_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '双击已添加文件夹,打开文件夹
    With Me.ListBox3
        If .ListCount = 0 Then Exit Sub
        Call OpenFileLocation(.Column(0, .ListIndex))
    End With
End Sub

Private Sub ListBox5_Click() '编辑-搜索框下拉-未完成
    Dim i As Integer
    i = Me.ListBox5.ListIndex
    If Me.CommandButton4.Caption = "查询书库" Then
        Me.TextBox1.Text = Me.ListBox5.Column(0, i) & " " & "&" & " " & Me.ListBox5.Column(1, i)
    ElseIf Me.CommandButton4.Caption = "单词查询" Then
        Me.TextBox1.Text = Me.ListBox5.Column(0, i)
    End If
End Sub

Private Sub ListView1_Click()
    With Me.ListView1
        If .ListItems.Count = 0 Then Exit Sub
        .FullRowSelect = True
        .ControlTipText = .SelectedItem.ListSubItems(1).Text
    End With
End Sub

Private Sub ListView1_DblClick() '搜索结果-在列表中直接双击打开文件
    Dim strx As String, strx1 As String, strx2 As String
    
    With Me.ListView1
        If .ListItems.Count = 0 Then Exit Sub
        strx = .SelectedItem.Text
        strx1 = .SelectedItem.ListSubItems(1).Text
        strx2 = LCase(Right$(strx1, Len(strx1) - InStrRev(strx1, ".")))
        If strx2 Like "xl*" Then
            If strx2 <> "xlsx" Then MsgBox "禁止打开此类文件", vbCritical, "Warning": Exit Sub 'excel类文件只允许xlsx格式的打开, 防止其他工作簿的宏干扰/冲突本工作薄
        End If
        '限制excel类文件的打开
        If CheckFileOpen(strx) = True Then Set Rng = Nothing: Exit Sub '判断文件是否存在于目录或者本地磁盘/文件是否处于打开的状态
        With Rng
            If OpenFile(strx, .Offset(0, 1).Value, .Offset(0, 2).Value, .Offset(0, 3).Value, 1, .Offset(0, 27).Value) = False Then
                Me.Label57.Caption = "异常"
                Set Rng = Nothing
                Exit Sub
            End If
        End With
        '----------窗体中打开文件路径(核心)都需要使用表格上的数据,防止非ansi
        With .SelectedItem
            If Len(.SubItems(4)) = 0 Then .SubItems(4) = 0 '空值和"0"还是有区别的
            .SubItems(4) = .SubItems(4) + 1 '打开次数+1
            Call OpenFileOver(strx, 1)
            End If
        End With
    End With
End Sub

Function ReSetDic() '页面切换-单词训练-重置
    With Me
        FlagStop = True '停止正在运行的进程
        .Label88.Caption = "状态:"
        .Label85.Caption = "耗时:"
        .Label84.Caption = "总时间"
        .Label90.Caption = ""     '试题
        .Label86.Caption = "题数:"
        .Label89.Caption = "已作答:"
        .Label94.Caption = "" '正确率
        .Label97.Caption = "" '提示
        .Frame3.Enabled = True
        .Frame1.Enabled = True
        .Frame10.Enabled = True
        .ComboBox10.Enabled = True
        .ComboBox10.Value = "" '测试数量
        .CommandButton71.Enabled = True '开始按钮
        .CommandButton72.Caption = "暂停" '暂停继续按钮
        .ListView3.ListItems.Clear '结果窗口清除内容
        .Label102.Caption = ""
        .CheckBox4.Value = False
        .CheckBox5.Value = False '选择按钮
        .CheckBox6.Value = False
    End With
End Function

Private Sub MultiPage1_Change()                '页面切换-操作
    Dim pgindex As Integer
    Dim arrpa() As Variant, i As Byte, ablow As Integer
    Dim url As String
    
    With Me
        pgindex = .MultiPage1.Value
        If .Label88.Caption <> "状态:" Then Call ReSetDic '只要单词页面发生变化,就执行重置
        If browser1 = 1 Then
            If pgindex <> 7 Then '页面切换,移除掉浏览器
            With Me!web
                browserkey = .LocationURL
            End With
            Me.Controls.Remove "web"
            browser1 = 0
            End If
        End If
        '----------------------------控制
        If pgindex = 1 Then '搜索
            .TextBox1.SetFocus
            
        ElseIf pgindex = 0 Then '主界面
            .TextBox8.SetFocus
            
        ElseIf pgindex = 2 And pgx2 = 0 Then '文件夹 '考虑到文件较多时运行较慢,取消自动显示列表
             With .ListView2
                .ColumnHeaders.Add , , "编码", 66, lvwColumnLeft
                .ColumnHeaders.Add , , "文件名", 116, lvwColumnLeft
                .ColumnHeaders.Add , , "类型", 32, lvwColumnLeft
                .ColumnHeaders.Add , , "文件位置", 221, lvwColumnLeft
                .View = lvwReport                            '以报表的格式显示
                .LabelEdit = lvwManual                       '使内容不可编辑
                .Gridlines = True
            End With
            pgx2 = 1
            
        ElseIf pgindex = 4 Then '工具
            .TextBox18.SetFocus
            .Label109.Caption = ""
            
        ElseIf pgindex = 3 Then '单词
            .TextBox14.SetFocus
            If pgx3 = 0 Then '注意listview不能在隐藏的状态下添加信息
                With .ListView3
                    .ColumnHeaders.Add , , "试题", 120, lvwColumnLeft
                    .ColumnHeaders.Add , , "答案", 60, lvwColumnLeft
                    .ColumnHeaders.Add , , "作答", 65, lvwColumnLeft
                    .ColumnHeaders.Add , , "Y/N", 40, lvwColumnLeft
                    .View = lvwReport                            '以报表的格式显示
                    .LabelEdit = lvwManual                       '使内容不可编辑
                    .Gridlines = True
                End With
                '单词测试数量
                .ComboBox10.List = Array(10, 15, 20)
                pgx3 = 1 '控制
            End If
            
        ElseIf pgindex = 7 Then '设置
            .TextBox11.SetFocus
            If pgx6 = 0 Then
                With ThisWorkbook.Sheets("temp")
                    ablow = .[aa65536].End(xlUp).Row
                    arrpa = .Range("ab1:ab" & ablow).Value
                End With
                If Len(arrpa(18, 1)) > 0 Then clickx = 1: .CheckBox1.Value = True '查询单词发音
                If Len(arrpa(29, 1)) > 0 Then clickx = 1: .CheckBox7.Value = True
                If Len(arrpa(31, 1)) > 0 Then clickx = 1: .CheckBox8.Value = True
                If Len(arrpa(19, 1)) > 0 Then .Label232.Caption = arrpa(19, 1) '显示文件被的最新的时间
                If Len(arrpa(27, 1)) > 0 Then .TextBox22.Text = arrpa(27, 1) '文件的备份上限
                If Len(arrpa(36, 1)) > 0 Then
                    .CheckBox11.Value = True '删除文件自动计算md5
                    If Len(arrpa(37, 1)) > 0 Then clickx = 1: .CheckBox12.Value = True
                End If
                If Len(arrpa(38, 1)) > 0 Then
                    If IsNumeric(arrpa(38, 1)) = True Then
                        i = arrpa(38, 1)
                        Select Case i
                            Case 1: clickx = 1: .CheckBox10.Value = True
                            Case 2: clickx = 1: .CheckBox11.Value = True
                            Case 3: clickx = 1: .CheckBox12.Value = True
                        End Select
                    End If
                End If
                If Len(arrpa(50, 1)) > 0 Then clickx = 1: .CheckBox19.Value = True
                If Len(arrpa(43, 1)) > 0 Then clickx = 1: .CheckBox17.Value = True 'pdf水印
                If Len(arrpa(53, 1)) > 0 Then clickx = 1: .CheckBox23.Value = True
                pgx6 = 1
            End If
        ElseIf pgindex = 6 Then
            CreateWebBrowser (browserkey)
        End If
    End With
End Sub

'Private Sub OptionButton1_Click() 'youdao -备用
'
'With ThisWorkbook.Sheets("首页")
'If Me.OptionButton1.Value = True Then .Cells(3, 1) = 1
'End With
'
'End Sub

'Private Sub OptionButton2_Click() 'baidu
'
'With ThisWorkbook.Sheets("首页")
'If Me.OptionButton2.Value = True Then .Cells(3, 1) = 2
'End With
'
'End Sub

'Private Sub OptionButton3_Click() 'bing
'With ThisWorkbook.Sheets("首页")
'If Me.OptionButton3.Value = True Then .Cells(3, 1) = 3
'End With
'End Sub

Private Sub OptionButton4_Click() '金山词霸
    With ThisWorkbook.Sheets("首页")
        If Me.OptionButton4.Value = True Then .Cells(3, 1) = 4
    End With
End Sub

Private Sub TextBox1_Change() '搜索/编辑-搜索框
    Dim arra() As Variant
    Dim arrB() As Variant
    Dim i As Integer, j As Integer, k As Integer
    Dim dic As New Dictionary
    Dim dica As New Dictionary
    Dim dicb As New Dictionary, blow As Integer
    Dim strx As String, m As Integer, n As Integer, p As Integer, xi As Integer, mi As Byte

'If Me.CommandButton4.Caption = "单词查询" Then                                           '首先查询本地的词库 -未完成
'   If Me.TextBox1.text Like "*#*" Then Exit Sub '输入的内容中包含数字,退出
'      If LenB(StrConv(Me.TextBox1, vbFromUnicode)) >= 4 Then '转换字符(vba无法区分中英文字符)
'        Me.ListBox5.Clear
'        sql = "select * from [" & TableName & "$] Where 英文 like '%" & Me.TextBox1.text & "%'or 中文 like '%" & Me.TextBox1.text & "%'or 自定义 like '%" & Me.TextBox1.text & "%'or 释义 like '%" & Me.TextBox1.text & "%'" '模糊搜索,百分号表示通配符"*"
'        Set rs = New ADODB.Recordset    '创建记录集对象
'        rs.Open sql, conn, adOpenKeyset, adLockOptimistic
'        If rs.BOF And rs.EOF Then                         '如果本地的词库为空
'           Me.ListBox5.Visible = False
'        Else
'        Me.ListBox5.Visible = True
'        For m = 1 To rs.RecordCount
'           Me.ListBox5.AddItem
'           Me.ListBox5.List(Me.ListBox5.ListCount - 1, 0) = rs(1) '英文
'           Me.ListBox5.List(Me.ListBox5.ListCount - 1, 1) = rs(3) '中文
'           Me.ListBox5.List(Me.ListBox5.ListCount - 1, 2) = rs(4) '自定义
'           Me.ListBox5.List(Me.ListBox5.ListCount - 1, 3) = rs(5) '释义
'           rs.MoveNext
'        Next
'        End If
'    Else
'    Me.ListBox5.Clear
'    End If
'End If

'    With ThisWorkbook.Sheets("书库") '注意这里的sheet被隐藏的问题
'        blow = docmx
'        'If Me.CommandButton4.Caption = "查询书库" Then
'        arra = .Range("b6:c" & blow).Value
'        arrB = .Range("r6:s" & blow).Value
'    End With
    If docmx < 8 Then
        Me.Label57.Caption = "数据库尚未存储数据"
        Exit Sub
    End If
    ArrayLoad
    strx = Me.TextBox1.Value
    strx = Replace(strx, "/", " ") '替换掉"/"符号
    With Me.ListBox5
        If Len(strx) >= 2 Then
            .Clear
            p = docmx - 5
            mi = 0
'            strx = strx & "*"
'            For j = 1 To 2 '减少循环的次数
                '暂时搜索1, 4列, 编号,文件路径(涵盖了主要的文件需要的信息),后期可以启用标签1,标签2的搜索
                For k = 1 To p
                    If InStr(1, arrax(k, 1) & "/" & arrax(k, 4), strx, vbTextCompare) > 0 Then
'                    If arrax(k, j) Like strx Then
                    '-------------------------------完整的instr函数的写法,开始字符的位置,源,比较的值,比较的方法;vbtextcompare表示不区分大小写进行比较
                        dic(arrax(k, 1)) = arrax(k, 2)
                        dica(arrax(k, 1)) = arrsx(k, 1)
                        dicb(arrax(k, 1)) = arrsx(k, 2)
                        mi = mi + 1
                        If mi > 10 Then GoTo 100
                    End If
                Next
'            Next
100
            xi = mi - 1 'dic.Count
            If xi >= 0 Then '注意dict.keys的某个值的正确写法应该dict.keys()(i), 如果已添加引用,且直接使用new dict进行初始化,就可以省略掉前面的括号
                For m = 0 To xi
                    .AddItem
'                    n = .ListCount - 1         '往最后一行写入内容
                    .List(m, 0) = dic.Keys(m)
                    .List(m, 1) = dic.Items(m)
                    .List(m, 2) = dica.Items(m)
                    .List(m, 3) = dicb.Items(m)
                Next '-----------------------当m=0经过for的时候,m会+1
            End If
            If xi >= 0 Then
                .Visible = True
            Else
                .Visible = False
            End If
        Else
            .Visible = False
        End If
    End With
End Sub

Private Sub TextBox3_Change() '主文件名修改
    Me.Label56.Caption = "" '清空这里的缓存信息,用于执行豆瓣书籍查找的判断
End Sub

Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) '防止主文件名当中输入系统禁止的字符,因为后续主文件名将被用于修改文件名
    Select Case KeyAscii
        Case asc("/"), asc("\"), asc(":"), asc("*"), asc("?"), asc("<"), asc(">"), asc("|")
        Me.Label57.Caption = "请勿输入非法字符：""/\ : * ? <> |"
        Me.TextBox3.Text = ""
        KeyAscii = 0
    Case Else
        Me.Label57.Caption = ""
    End Select
End Sub

Private Sub TextBox4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) '作者
    Select Case KeyAscii
        Case asc("/"), asc("\"), asc(":"), asc("*"), asc("?"), asc("<"), asc(">"), asc("|")
        Me.Label57.Caption = "请勿输入非法字符：""/\ : * ? <> |"
        Me.TextBox3.Text = ""
        KeyAscii = 0
    Case Else
        Me.Label57.Caption = ""
    End Select
End Sub

Private Sub CheckBox24_Click() '文本搜索
    With Me.CheckBox24
        If .Value = True Then
            searchx = 1
        Else
            searchx = 0
        End If
    End With
End Sub

Private Sub CommandButton151_Click() '查询文本
    Dim strx As String, strLen As Byte, strx1 As String * 1
    Dim i As Byte, j As Byte, k As Integer, m As Byte, p As Byte, blow As Integer, mi As Byte
    
    With Me
        If searchx = 0 Then .Label57.Caption = "未勾选文本搜索": Exit Sub
        strx = .TextBox8.Text
        strLen = Len(strx)
        If strLen < 2 Then Exit Sub
        If docmx < 8 Then
            .Label57.Caption = "数据库尚未存储数据"
            Exit Sub
        End If
        
        For i = 1 To strLen
            strx1 = Mid(strx, i, 1)
            If strx1 Like "[一-龥]" Then
                j = j + 1
            ElseIf strx1 Like "#" Then
                k = k + 1
            ElseIf strx1 Like "[a-zA-Z]" Then
                m = m + 1
            End If
        Next
        If j < 2 And k = 0 And m = 0 Then
            p = 1
        ElseIf j = 0 And k < 4 And m = 0 Then
            p = 1
        ElseIf j = 0 And k = 0 And m < 4 Then
            p = 1
        End If
        If p = 1 Then
            .Label57.Caption = "输入内容长度不满足要求'"
            Exit Sub
        End If
        '-----------------------------------前期判断
        ArrayLoad
        mi = 0
        blow = docmx - 5
        If .CheckBox26.Value = True Then '勾选极速模式
            With .ListView1.ListItems
                .Clear
                For k = 1 To blow
                    If arrax(k, 3) = "txt" Then
                        If fso.fileexists(arrax(k, 4)) = True Then
                            If FindTextInFile(arrax(k, 4), strx) > 0 Then
                                With .Add
                                    .Text = arrax(k, 1)
                                    .SubItems(1) = arrax(k, 2)
                                    .SubItems(2) = arrax(k, 3)
                                    .SubItems(3) = arrax(k, 4)
                                    .SubItems(4) = arrbx(k, 1)
                                End With
                                mi = mi + 1
                                If mi > 50 Then Exit For
                            End If
                        End If
                    End If
                Next
            End With
        Else
            If .CheckBox25.Value = True Then '不区分大小写,执行vbtext比较
                With .ListView1.ListItems
                    .Clear
                    For k = 1 To blow
                        If arrax(k, 3) = "txt" Then
                            If fso.fileexists(arrax(k, 4)) = True Then
                                If CheckFileKeyWord(arrax(k, 4), strx, 1, 0) = True Then '执行文本比较,检查文件的编码格式
                                    With .Add
                                        .Text = arrax(k, 1)
                                        .SubItems(1) = arrax(k, 2)
                                        .SubItems(2) = arrax(k, 3)
                                        .SubItems(3) = arrax(k, 4)
                                        .SubItems(4) = arrbx(k, 1)
                                    End With
                                    mi = mi + 1
                                    If mi > 50 Then Exit For
                                End If
                            End If
                        End If
                    Next
                End With
            Else
                With .ListView1.ListItems '区分大小写,二进制比较,执行的速度更快
                    .Clear
                    For k = 1 To blow
                        If arrax(k, 3) = "txt" Then
                            If fso.fileexists(arrax(k, 4)) = True Then
                                If CheckFileKeyWordC(arrax(k, 4), strx) = True Then
                                    With .Add
                                        .Text = arrax(k, 1)
                                        .SubItems(1) = arrax(k, 2)
                                        .SubItems(2) = arrax(k, 3)
                                        .SubItems(3) = arrax(k, 4)
                                        .SubItems(4) = arrbx(k, 1)
                                    End With
                                    mi = mi + 1
                                    If mi > 50 Then Exit For
                                End If
                            End If
                        End If
                    Next
                End With
            End If
        End If
        If mi = 0 Then .Label57.Caption = "未找到"
    End With
End Sub

Private Sub TextBox8_Change()   '主界面文本输入 '可考虑改成非动态的搜索,当数据的量足够大的时候 '整合多种搜索模式
    Dim i As Byte, j As Byte, k As Integer, blow As Integer, n As Byte, m As Byte, p As Byte, t As Byte
    Dim dic As New Dictionary
    Dim dica As New Dictionary
    Dim dicb As New Dictionary
    Dim dicc As New Dictionary
    Dim strx As String, strLen As Byte
    Dim strx1 As String, strx2 As String
    Dim mi As Byte
    Dim xi As Variant
    Dim strTemp As String, strtempx As String, strtempx1 As String, chk As Byte
    '需要根据英文,数字,汉字,混合内容(含空格)来进行调整
    If searchx = 1 Then Exit Sub '执行文本分析
    If docmx < 8 Then
        Me.Label57.Caption = "数据库尚未存储数据"
        Exit Sub
    End If
    strx = Me.TextBox8.Value
    strx2 = Replace(strx, "/", " ") '替换掉"/"符号, 后面使用"/"作为连接符
    strLen = Len(strx)        '限制最长的输入的长度为38
    With Me.ListView1.ListItems
        If strLen >= 2 Then             '超过两个字符就做出反应
            .Clear                 '对作业区进行清理
            blow = docmx - 5
            ArrayLoad
            mi = 0
            For k = 1 To blow '注意由于筛选的目标之间有相同的字符,将导致出现多行结果的bug,这里使用字典的方法来解决
                chk = 0
                p = 0: j = 0: n = 0
                strtempx = ""
                strtempx1 = ""
                strTemp = arrax(k, 1) & "/" & arrax(k, 4) '"/"间隔符
                If InStr(1, strTemp, strx2, vbTextCompare) > 0 Then '搜索的式还可以有很大的调整空间'制作一张近似词的表,当搜索英文的某些词语可以同步检索'将拼写错误的词替换掉进行搜索
                    chk = 1
                Else
                    If strLen >= 3 Then
                        If InStr(strx, Chr(32)) > 0 Then '输入的内容存在空格,将空格的内容拆开进行检查
                            xi = Split(strx, Chr(32)) '以空格作为分割
                            i = UBound(xi)
                            For t = 0 To i
                                strtempx = xi(t)
                                If strtempx Like "[一-龥]" Then '单个中文直接进行判断
                                    If InStr(1, strTemp, strtempx, vbTextCompare) > 0 Then
                                        p = p + 1
                                        If p >= 2 Then chk = 1: Exit For
                                    End If
                                Else
                                    If Len(strtempx) >= 2 Then
                                        If InStr(1, strTemp, strtempx, vbTextCompare) > 0 Then chk = 1: Exit For  '表示满足要求
                                    End If
                                End If
                            Next
                        Else
                            '--------- like的用法: https://analystcave.com/vba-like-operator/
                            For i = 1 To strLen '判断输入额内容是否为复合内容
                                strx1 = Mid(strx, i, 1)
                                If strx1 Like "[一-龥]" Then       '中文 '可单个处理,可集中处理 "公司财务" 可拆解成单个单词 公/司/财/务, 所以可以不处理空格
                                    If InStr(1, strTemp, strx1, vbTextCompare) > 0 Then
                                        p = p + 1
                                        If p >= 2 Then GoTo 98 '优先处理中文
                                    End If
                                ElseIf strx1 Like "[a-zA-Z]" Then  '英文字母,含大小写 '集中处理,单个字母的含义大幅下降
                                    strtempx = strtempx & strx1
                                ElseIf strx1 Like "[0-9]" Then     '数字或者使用 "#"来表示任意单个0-9数字 '集中处理,单个处理没有单词拆解的影响大
                                    strtempx1 = strtempx1 & strx1
                                End If
                            Next
                            j = Len(strtempx)
                            n = Len(strtempx1)
                            If p = 0 Then
                                If j = 0 Or n = 0 Then GoTo 99
                            Else
                                If strLen - p < 3 Then GoTo 99
                            End If
                            If j >= 3 Then
                                If InStr(1, strTemp, strtempx, vbTextCompare) > 0 Then chk = 1
                            End If
                            If n >= 3 Then
                                If InStr(1, strTemp, strtempx1, vbTextCompare) > 0 Then chk = 1
                            End If
                        End If
                    End If
                End If
                If chk = 1 Then
98
                    dic(arrax(k, 1)) = arrax(k, 2)
                    dica(arrax(k, 1)) = arrax(k, 3)
                    dicb(arrax(k, 1)) = arrax(k, 4)
                    dicc(arrax(k, 1)) = arrbx(k, 1)  '单一for循环可以不需要字典,直接使用数组即可/或者直接添加值, 多个for循环的时候才需要
                    mi = mi + 1
                    If mi > 50 Then GoTo 100 '限制搜索结果的数量
                End If
99
            Next
100
            '----------------------------------------------------------------数据写入listview
            If mi > 0 Then '获取到有效的值
                mi = mi - 1
                For m = 0 To mi
                    With .Add
                        .Text = dic.Keys(m)
                        .SubItems(1) = dic.Items(m)
                        .SubItems(2) = dica.Items(m)
                        .SubItems(3) = dicb.Items(m)
                        .SubItems(4) = dicc.Items(m)
                    End With
                Next
            Else
                .Clear
            End If
        Else
            .Clear                   '当输入的文字少于2位时,进行清理
        End If
    End With
End Sub

Private Sub UserForm_Activate() '活跃窗体
    Dim blow As Integer
    Dim strx As String, strx2 As String
    
    If Statisticsx = 1 Then Exit Sub
    If Workbooks.Count = 1 Then '只有一个工作簿的时候显示最小化窗口
        AddIcon    '添加图标
        AddMinimiseButton   '添加按钮
        AppTasklist Me    '添加任务栏
    End If
    If UF3Show = 3 Then ''变更激活的方式,假如隐藏状态, 当出现数据变化,就发送变化, 然后在从隐藏状态恢复时更新数据
        UF3Show = 1
        Call PauseRm '禁用selection事件
        If DeleFilex = 1 Then '文件被删除
            strx = Me.Label29.Caption
            If Len(strx) > 0 Then '有值
                SearchFile strx
                If Rng Is Nothing Then
                    DeleFileOverx strx
                    Exit Sub
                End If
            End If
        End If
        DataUpdate
    End If
'    Call EveSpy '让事件监听保持运行状态
'    Call RecData '保证ado处于连接状态
'    Call LockSet '保持表格的锁定处于vba可写入状态
End Sub

Sub DataUpdate() '更新窗体的数据, 使用完成,需要将参数重置
    If AddPlistx = 1 Then Call PrReadList: AddPlistx = 0 '在隐藏的状态下,优先阅读列表发生变化
    If OpenFilex = 1 Then Call RecentUpdate: OpenFilex = 0 '最近阅读
    If MDeleFilex = 1 Then Call AddFileListx: MDeleFilex = 0 '添加的文件夹 '添加的文件夹被移除 '数据库内的内容发生了变化
    If DeleFilex = 1 Then Call CwUpdate: Call Choicex: DeleFilex = 0 '更新数据库 '书库详 '筛选区域
    docmx = ThisWorkbook.Sheets("书库").[d65536].End(xlUp).Row '关键参数
End Sub

Sub ArrayLoad() '将部分的内容加载到数组,加快响应的速度
    If spyx <> docmx Then '利用模块级变量保存搜索区域的值在内存中,减少访问表格的需要,只有当表格的数据发生变化才重新获取值,加快访问的速度
        spyx = docmx '-----------初始赋值/变化在进行赋值 ' '这个值由窗体活跃来自动获取 '当执行完数据更新需要考虑如何更新这个值
        If SafeArrayGetDim(arrax) <> 0 Then '判断数组是否经过初始化, 如果发生了初始化,就将原有的数组抹掉
            Erase arrax
            Erase arrbx                        '当数据发生变化的时候,抹掉旧的数据重新获取新的
            Erase arrsx
        End If
        With ThisWorkbook.Sheets("书库")
            arrax = .Range("b6:f" & docmx).Value '编号,文件名, 扩展名, 文件路径, 文件所在位置
            arrbx = .Range("n6:n" & docmx).Value '打开次数
            arrsx = .Range("s6:t" & docmx).Value '评分/推荐指数
        End With
'            arrux = .Range("u6:v" & docmx).Value '标签1/标签2
    End If
End Sub

Private Sub UserForm_Initialize() '窗体初始化
    Dim dic As New Dictionary
'    Dim arr() As String
    Dim strx As String
    Dim i As Byte, TableName As String, k As Integer
    
    If Statisticsx = 1 Then Exit Sub
    UF3Show = 1 '标记窗体处于激活的状态
    NewM = False '用于控制textbox菜单
    With Me
        .MultiPage1.Value = 0 '打开窗口时显示主界面
        '初始化listview
        With .ListView1  '注意listview必须在可见的状态下完成初始化
            .ColumnHeaders.Add , , "编码", 66, lvwColumnLeft
            .ColumnHeaders.Add , , "文件名", 122, lvwColumnLeft
            .ColumnHeaders.Add , , "类型", 33, lvwColumnLeft
            .ColumnHeaders.Add , , "文件路径", 225, lvwColumnLeft '需要留出一定的空间,否则会出现下横滑动栏
            .ColumnHeaders.Add , , "打开次数", 50, lvwColumnLeft
            .View = lvwReport                            '以报表的格式显示
            .LabelEdit = lvwManual                       '使内容不可编辑
            .Gridlines = True
        End With
        
        With .ListBox3 '文件夹
            .MultiSelect = fmMultiSelectMulti '多选
            .ListStyle = fmListStyleOption
        End With
        '工具设置
        .ComboBox11.List = Array("浏览器", "Axure", "Mind", "Note", "PDF", "截图", "Spy++", "备份", "解压目录")
        '内容评分
        .ComboBox5.List = Array(1, 2, 3, 4, 5)
        'pdf清晰度
        .ComboBox3.List = Array(1, 2, 3)
        '文本质量
        .ComboBox4.List = Array(1, 2, 3)
        '文件操作
        .ComboBox6.List = Array("打开", "删除", "打开位置", "导出文件", "添加到导出列表")
        '推荐指数
        .ComboBox2.List = Array(1, 2, 3)
        '文字类型
        .ComboBox12.List = Array("CNS", "CNT", "EN", "JPN", "OTS", "MIX") '简体中文,繁体,英文,日文,其他,混合内容
        .ComboBox13.List = Array(6, 12, 18)
        .ComboBox14.List = Array("Start", "Reading", "Over") '阅读状态
    End With
    
    PauseRm '禁用selection事件,worksheetchange事件很容易触发,在窗体显示时,禁用掉
    
    With ThisWorkbook
        docmx = .Sheets("书库").[d65536].End(xlUp).Row
        voicex = .Sheets("temp").Range("ab18").Value
        Choicex      '筛选区域
        RecentUpdate '最近阅读
        PrReadList   '优先阅读
        AddFileListx '添加的文件夹
        CwUpdate     '书库详情
'        If .Sheets("temp").Range("ab18") = 1 Then Me.CommandButton44.Visible = True '发音按钮-未完成
    End With
    Exit Sub
    '初始化备忘录
'    If RecData = True Then
'        TableName = "备忘录"
'        Sql = "select * from [" & TableName & "$]"
'        Set Rs = New ADODB.Recordset    '创建记录集对象
'        Rs.Open Sql, Conn, adOpenKeyset, adLockOptimistic
'        If Rs.BOF And Rs.EOF Then             '用于判断有无找到数据
'            Me.TextBox10.Enabled = False
'        Else
'            k = Rs.RecordCount
''            ReDim arr(1 To k)
'            For i = 1 To k
'                dic(CStr(Rs.Fields(0))) = ""      '字典的方式存储日期
''                arr(i) = Rs(2)                   '数组的方式存储内容(注意2,是横向的数据,不是下一个数据)
'                If i = k Then strx = Rs(2)
'                Rs.MoveNext                       '数据集指针,指向下一个数据
'            Next
'            With Me
''                .TextBox10.Text = arr(k)
'                .TextBox10.Text = strx
'                .ComboBox9.List = dic.Keys
'                .ComboBox9.Text = dic.Keys(UBound(dic.Keys)) '取最后的值
'            End With
'        End If
'        Rs.Close
'        Set Rs = Nothing
'    End If
End Sub

Private Sub Choicex() '筛选区域
    Dim dicop As New Dictionary
    Dim num As Integer, n As Integer
    Dim arrop As Variant
    
    With ThisWorkbook.Sheets("书库")
        n = docmx
        If n > 6 Then
            arrop = .Range("d6:d" & n).Value
            dicop.CompareMode = TextCompare '不区分大小写
            n = n - 5
            For num = 1 To n
                dicop(arrop(num, 1)) = "" '获得字典的key(不重叠数据)
            Next
        ElseIf n = 6 Then
            dicop(.Range("d6").Value) = ""
        ElseIf n < 6 Then
            Exit Sub
        End If
    End With
    With Me
        .ComboBox1.List = dicop.Keys
        .ComboBox7.List = Array(1, 2, 3, 4, 5)
        .ComboBox8.List = Array(1, 2, 3)
    End With
End Sub

Private Function RecentUpdate() ''最近阅读(修改)
    Dim k As Byte, i As Byte
    
    With ThisWorkbook.Sheets("主界面")
        '最近阅读
        If Len(.Range("p27").Value) > 0 Then
            Me.ListBox1.Clear
            For k = 27 To 33
                If Len(.Range("p" & k).Value) > 0 Then
                   Me.ListBox1.AddItem
                   i = Me.ListBox1.ListCount - 1
                   Me.ListBox1.List(i, 0) = .Range("u" & k)
                   Me.ListBox1.List(i, 1) = .Range("p" & k)
                   Me.ListBox1.List(i, 2) = .Range("w" & k)
                Else
                   Exit For
                End If
            Next
        End If
    End With
End Function

Private Function PrReadList()  '优先阅读
    Dim m As Byte, i As Byte
    
    With ThisWorkbook.Sheets("主界面")
        '优先阅读
        If Len(.Range("i27").Value) > 0 Then                      '不添加空白的值进来
            Me.ListBox2.Clear
            For m = 27 To 33
               If Len(.Range("d" & m).Value) > 0 Then
                    Me.ListBox2.AddItem
                    i = Me.ListBox2.ListCount - 1
                    Me.ListBox2.List(i, 0) = .Range("i" & m).Value
                    Me.ListBox2.List(i, 1) = .Range("d" & m).Value
                    Me.ListBox2.List(i, 2) = .Range("k" & m).Value
                Else
                    Exit For
               End If
            Next
        End If
    End With
End Function
 
Private Function AddFileListx() '添加的文件夹
    Dim j As Byte, Elow As Byte, i As Byte, m As Byte
    Dim strx As String
    
    With ThisWorkbook.Sheets("主界面")
        If Len(.Range("e37").Value) > 0 Then '不添加空白的值进来
            Elow = .[e65536].End(xlUp).Row
            m = Elow - 37
            ReDim arraddfolder(m) '--------------用于临时存储,防止非ansi字符
            Me.ListBox3.Clear '清除之前的数据
            For j = 37 To Elow
                If Len(.Range("e" & j).Value) > 0 Then
                    Me.ListBox3.AddItem
                    i = Me.ListBox3.ListCount - 1
                    strx = .Range("e" & j).Value
                    Me.ListBox3.List(i, 0) = strx
                    arraddfolder(i) = strx
                    If fso.folderexists(strx) = False Then
                        Me.ListBox3.List(i, 1) = "该文件夹已移除" '检查文件夹是否已经被移除
                    Else
                        Me.ListBox3.List(i, 1) = .Range("i" & j).Value
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End With
End Function

Private Function CwUpdate() '窗口数据更新
    Dim str1 As Integer
    Dim str2 As String '注意这项
    Dim str3 As Integer
    Dim str4 As Integer
    Dim str5 As Integer
    Dim str6 As Integer
    Dim str7 As Integer
    Dim str8 As Integer

    With ThisWorkbook.Sheets("主界面")
        str1 = .Range("p37").Value '文件总数
        str2 = .Range("p38").Value '所有文件大小
        str3 = .Range("p40").Value 'pdf
        str4 = .Range("s40").Value 'EPUB
        str5 = .Range("p42").Value '其他
        str6 = .Range("p41").Value 'PPT
        str7 = .Range("v41").Value 'Word
        str8 = .Range("s41").Value 'Excel
    End With
    With Me
        .Label47.Caption = str1
        .Label48.Caption = str2
        .Label49.Caption = str3
        .Label50.Caption = str4
        .Label51.Caption = str5
        .Label52.Caption = str6
        .Label53.Caption = str7
        .Label54.Caption = str8
    End With
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) '隐藏窗口后显示uf4
    'https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/queryclose-constants
    If Statisticsx = 1 Then Exit Sub
    UF3Show = 3 '用于表示窗体处于隐藏的状态
    If CloseMode = vbFormControlMenu Then                                    '可以修改禁止使用"x"按钮来关闭窗口
            Cancel = True
            Me.Hide
            If Workbooks.Count = 1 And Application.Visible = False Then
                UserForm4.Hide
                UserForm4.Show 1
                UserForm4.Caption = "锁定"
            Else
                UserForm4.Caption = "Mini"
                UserForm4.Show
            End If
    End If
    Timeset = 2
    If FlagStop = False Then FlagStop = True '确保单词训练当中的计时处于停止的状态,在使用时间相关的代码多需要考虑时间停止的问题
End Sub

Private Sub UserForm_Terminate()
    If Statisticsx = 1 Then Exit Sub
    UF3Show = 0
End Sub
