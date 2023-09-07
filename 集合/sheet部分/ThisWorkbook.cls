VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Appspy As Eventspy '创建全局事件监听(模块级)
'需要注意将一些经常用的参数加载到内存中去,以加速进程的执行

'Private Sub Workbook_Activate() '隐藏
'Exit Sub
'    Application.DisplayFormulaBar = False
'End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
'    .Application.DisplayFormulaBar = True '启用公式栏
    Timeset = 2 '-----时间类进程在关闭工作簿时需要保证处于关闭的状态,以避免某些事件无法关闭,重新打开工作簿
    With ThisWorkbook
        .Close savechanges:=True '自动保存文件
        If Conn.State = adStateOpen Then Conn.Close           '断开与文件的连接
        Call ResetMenu '关闭前重置右键菜单
        Set rs = Nothing
        Set Conn = Nothing   '清空
    End With
End Sub

Private Sub Workbook_Deactivate() '显示公式栏/重置右键菜单
'    Application.DisplayFormulaBar = True '启用讲导致工作表中的控件按钮位置发生变化
    With ThisWorkbook.Sheets("temp")
    If Len(.Range("ab9")) = 0 Or .Range("ab9") <> 0 Then '0/1用于控制右键菜单的创建或者重置，避免不必要的重复动作
        Call ResetMenu
        .Range("ab9") = 0
    End If
    End With
End Sub

Private Sub Workbook_Open() '启动
    Dim lastpath As String
    Dim Aversion As Integer
    Dim userpathx As String, OSversion As String, strx As String, strx1 As String
    Dim yesno As Variant
    Dim yesno1 As Variant, drx As Drive
    Dim arrfolder() As Variant, folderx As Byte, i As Byte, k As Byte '检查文件夹是否存在
    Exit Sub
'200 '重置
'    With ThisWorkbook
'    On Error GoTo 1000
'    If Len(.Sheets("temp").Range("ab2").Value) = 0 Then '初始化校检 '尚未写入数据
'
'        Aversion = CInt(.Application.Version) 'cint函数转换数据为Integer型 '获取Excel的版本号 '禁止在2010版本之下运行
'        If Aversion < 14 Then '14对应2010版本,15 2013, 16, 2016
'            MsgBox "你的Excel版本太旧了,不能为你服务", vbOKOnly, "Sorry...!"
'            Exit Sub
'        End If
'        If CheckExcel64 = True Then MsgBox "本程序不兼容X64 Excel!", vbOKOnly, "Sorry...!": Exit Sub '判断excel是否为64位版本
'        OSversion = GetOsVersion
'        If OSversion <> "Win7" And OSversion <> "Win8" And OSversion <> "Win8.1" And OSversion <> "Win10" Then '限制运行的系统 '2010版本的office依然可以在xp上运行
'            MsgBox "你的系统版本太旧了,不能为你服务", vbOKOnly, "Sorry...!"
'            Exit Sub
'        End If
'
'        With UserForm6
'            .Show 0
'            .Label1 = "初始化中...请稍后,这个过程请勿进行其他操作(需要几十秒)!"
'        End With
'
'        '检查非系统盘是否存在,'判断磁盘的类型,2表示固定的磁盘'以防止在移动磁盘或者光驱上创建文件 Environ("SYSTEMDRIVE")=系统盘符
'        For Each drx In fso.Drives
'            If drx.DriveType = 2 Then '必须注意在多条件判断时,条件的执行后否造成错误,如,某值的判断必须不是空才能判断,如果是空就会出错,显然需要先确定这个值是否为空才进行下一步的判断
'                If drx.AvailableSpace > 20971520 And drx.Path <> Environ("SYSTEMDRIVE") Then userpathx = drx.Path & "\lbrecord": Exit For
'            End If
'        Next
'
'        If Len(userpathx) = 0 Then
'            strx1 = Environ("UserProfile") & "\documents"
'            If fso.FolderExists(strx1) = False Then fso.CreateFolder (strx1)
'            userpathx = Environ("UserProfile") & "\documents\lbrecord"
'        End If
'        If fso.FileExists(userpathx & "\lbrecord.xlsx") = True Then '检查文件是否已被创建'如果存在就退出
'            Unload UserForm6
'            MsgBox "同名的文件位于程序需要创建存储路径的位置,请移除该文件", vbCritical, "Warning"
'            Exit Sub
'        End If
'
'        With .Application
'            .EnableEvents = False '启用事件
'            .DisplayAlerts = False
'            .ScreenUpdating = False
'            .Sheets("首页").Activate
''            .Sheets("首页").ScrollArea = "a1:t42" '可滑动的区域
'        End With
'
'        If fso.FolderExists(userpathx) = False Then
'            fso.CreateFolder (userpathx) '创建文件夹
'            fso.GetFolder(userpathx).Attributes = Hidden '设置为隐藏
'        End If
'
'        lastpath = userpathx & "\lbrecord.xlsx"
'
'        Call CreateWorksheet(lastpath) '创建表格 '顺序不能乱
'        Call Assist(userpathx) '创建桌面快捷键和移动文件
'        Call CheckKeyExe
'        Call HideOption '设置窗口的显示效果
'
'        With .Sheets("temp") '等信息汇总完了才开始全部写入数据
'            .Range("ab2") = userpathx '程序的存储位置
'            .Range("ab3") = lastpath '写入内容
'            .Range("ab22") = Environ("UserProfile") '系统用户文件夹所在,用于后续程序所在发生变化的判断
'            .Range("ab23") = OSversion '系统版本
'            .Range("ab24") = Aversion '程序版本
'        End With
'        .Save '保存文件
'
'        fso.CopyFile (.fullname), userpathx & "\"      '复制文件到新的文件夹
'
'        With .Application
'            .EnableEvents = True '启用事件
'            .DisplayAlerts = True
'            .ScreenUpdating = True
'        End With
'
'        Unload UserForm6
'        UserForm11.Show '显示初始化的状态
'
'    Else '-----------------------------------------完成初始化后
'        With .Application
'            .EnableEvents = False '启用事件
'            .DisplayAlerts = False
'            .ScreenUpdating = False
'        End With
'        '---初始化载入一部分的信息
'
'        With .Sheets("temp")                    '检查文件的存在状况和处理文件丢失等情况
'            lastpath = .Range("ab3").Value
'            If fso.FileExists(lastpath) = False Then
'                lastpath = ThisWorkbook.Path & "\lbrecord.xlsx"
'                If fso.FileExists(lastpath) = False Then
'                    If Environ("UserProfile") <> .Range("ab22").Valu Then '用户的系统发生改变
'                        .Range("ab2:ab100").ClearContents
'                        GoTo 200 '重新设置文件
'                    Else
'                        Call CreateWorksheet(lastpath) '创新创建表
'                        .Range("ab3") = lastpath
'                    End If
'                End If
'            End If
'        End With
'
'        With .Sheets("书库").CommandButton21
'            If .Caption = "退出调试" Then
'                .Caption = "调试模式"
'                .ForeColor = &H80000012
'                .CommandButton1.Enabled = True
'                .CommandButton11.Enabled = True
'            End If
'        End With
'
'        With .Sheets("主界面")
'            folderx = .[e65536].End(xlUp).Row
'            If folderx = 37 Then '只有一行数据
'               strx = Trim(.Cells(37, "e").Value)
'                If Len(strx) > 0 Then
'                If fso.FolderExists(strx) = False Then MsgBox "文件夹丢失", vbCritical, "Warning!!!": GoTo 1000 '当添加的文件夹丢失
'                End If
'            ElseIf folderx > 37 Then
'                arrfolder = .Range("e37:e" & folderx).Value
'                k = folderx - 37
'                For i = 1 To k
'                    strx = Trim(arrfolder(i, 1))
'                    If Len(strx) > 0 Then
'                    If fso.FolderExists(strx) = False Then MsgBox arrfolder(i, 1) & "文件夹丢失", vbCritical, "Warning!!!": GoTo 1000
'                    End If
'                Next
'            End If
'        End With
'
'1001 '如果连接失败-进行重新连接
'        If Len(.Sheets("temp").Cells(54, "ab").Value) > 0 Then '进入调试模式
'        Conn.Open "Provider = Microsoft.ACE.OLEDB.12.0;Data Source=" & lastpath & ";extended properties=""excel 12.0;HDR=YES""" '打开数据存储文件
'        If Conn.State = adStateOpen Then        '检查连接成功
''            Call LockSet '保持表格的锁定处于vba可写入状态
'            With .Sheets("首页")
'                .Activate
'                .ScrollArea = "a1:t42"
'            End With
'
'            If Workbooks.Count = 1 Then            '当打开的文件只有一个Excel工作簿时,隐藏主窗口
'                With .Application
'                If .Visible = True Then .Visible = False
'                    .EnableEvents = True '启用事件
'                    .DisplayAlerts = True
'                    .ScreenUpdating = True
'                End With
'                UserForm3.Show
'            ElseIf Workbooks.Count > 1 Then
'                .Application.ActiveWindow.WindowState = xlMinimized  '否则就最小化窗口
'                UserForm4.Show
'                UserForm3.Show 0
'            End If
'            With .Application
'                .EnableEvents = True '启用事件
'                .DisplayAlerts = True
'                .ScreenUpdating = True
'            End With
'        Else
'            Set Conn = Nothing
'            yesno = MsgBox("存储文件没有成功连接，是否重新连接?_", vbYesNo) '如果无法正常连接存储文件的处理
'            If yesno = vbYes Then
'                GoTo 1001
'            Else
'                yesno1 = MsgBox("未能连接数据,是否退出程序?_", vbYesNo)
'                If yesno1 = vbYes Then
'                    .Close savechanges:=True
'                Else
'                    .Sheets("书库").Activate
'                    .Sheets("书库").Label1.Caption = "程序异常,建议不继续使用"
'                End If
'            End If
'        End If
'    End If
'
'1000
'        With .Application
'            .EnableEvents = True '启用事件
'            .DisplayAlerts = True
'            .ScreenUpdating = True
'        End With
'    End With
End Sub

Private Sub Workbook_SheetActivate(ByVal sH As Object) '表发生变化,鼠标右键菜单变化
    With ThisWorkbook.Sheets("temp")
        If sH.Name <> "书库" Then    '如果活跃的表格不是书库就取消掉附加的右键菜单
            If Len(.Range("ab9").Value) = 0 Or .Range("ab9").Value <> 0 Then '0/1用于控制右键菜单的创建或者重置，避免不必要的重复动作
                Call ResetMenu
                .Range("ab9") = 0 '如果打开的表不是书库这张表就重置右键菜单
            End If
        End If
    End With
End Sub

Private Function Assist(ByVal userpath As String) '创建桌面快捷方式/移动文件
    Dim MyPath As String, strx As String, myWshc As Object, mySht As Object
    
    With ThisWorkbook.Sheets("temp")
        strx = ThisWorkbook.Path & "\"
        If fso.fileexists(strx & "whitenoise.mp4") = True Then
            fso.MoveFile (strx & "whitenoise.mp4"), userpath & "\" '移动音频
            .Range("ab8") = userpath & "\whitenoise.mp4"
        End If
        
        If fso.fileexists(strx & "speak.vbs") = True Then '移动vbs脚本
            fso.MoveFile (strx & "speak.vbs"), userpath & "\"
            .Range("ab17") = userpath & "\speak.vbs"
        End If
        
        If fso.fileexists(strx & "fileisopen.ps1") = True Then '移动powershell脚本
            fso.MoveFile (strx & "fileisopen.ps1"), userpath & "\"
            .Range("ab7") = userpath & "\fileisopen.ps1"
        End If
        
        If fso.fileexists(strx & "单词表.xlsx") = True Then '移动单词表格
            fso.MoveFile (strx & "单词表.xlsx"), userpath & "\"
           .Range("ab25") = userpath & "\单词表.xlsx"
        End If
        
        If fso.fileexists(strx & "help.pdf") = True Then '移动帮助文件
            fso.MoveFile (strx & "help.pdf"), userpath & "\"
           .Range("ab39") = userpath & "\help.pdf"
        End If
        
        If fso.fileexists(strx & "lb.ico") = False Then Exit Function '检查ico文件是否存在-常见文件快捷方式设置桌面图标
        Set myWshc = CreateObject("Wscript.Shell")
        fso.MoveFile (strx & "lb.ico"), userpath & "\" '移动ico文件到新的文件夹
        MyPath = myWshc.specialfolders("Desktop")
        Set mySht = myWshc.CreateShortcut(MyPath & "\LB.lnk")
        With mySht
            .TargetPath = userpath & "\" & "\" & ThisWorkbook.Name     '指定文件的路径
            .IconLocation = userpath & "\" & "\lb.ICO"    '设定图标
            .Save
        End With
    End With
    Set mySht = Nothing
    Set myWshc = Nothing
End Function

Sub CheckKeyExe() '检查Powershell, ISE,IE的情况 '直接检查文件夹路径,也可以查看注册表(卸载出现残余),查看安装软件的情况
    Dim wsh As Object
    Dim wExec As Object
    Dim Result As String
    Dim psversion As String, psversionx As Byte
    '------------------------------------------HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall可以查看所有安装的软件
    With ThisWorkbook.Sheets("temp")
        If fso.fileexists(Environ("SystemRoot") & "\system32\WindowsPowerShell\v1.0\powershell.exe") = True Then '执行判断powershell是否存在'Environ("SystemRoot")表示Windows文件夹
            .Cells(4, "ab") = 1
            Set wsh = CreateObject("WScript.Shell")
            Set wExec = wsh.Exec("powershell Get-Host | Select-Object Version") '获取powershell版本号
            Result = wExec.StdOut.ReadAll
            psversion = Trim(Replace(Split(Result, "-")(UBound(Split(Result, "-"))), Chr(10), "")) '处理获取到的信息
            psversionx = CInt(psversion) '转换为integer数据
            If psversionx > 3 Then .Cells(4, "ac") = psversionx '只有当版本大于3的时候才写入数据
            If fso.fileexists(Environ("SystemRoot") & "\System32\WindowsPowerShell\v1.0\powershell_ise.exe") = True Then .Cells(5, "ab") = 1 '检查是否有powershell ISE
            Set wExec = Nothing
            Set wsh = Nothing
        End If
        If fso.fileexists(Environ("UserProfile") & "\AppData\Local\Google\Chrome\Application\chrome.exe") = True Then .Cells(10, "ab") = Environ("UserProfile") & "\AppData\Local\Google\Chrome\Application\chrome.exe" '检查chrome浏览器
        
        If fso.fileexists(Environ("SYSTEMDRIVE") & "\Program Files\Internet Explorer\iexplore.exe") = True Then
            .Cells(6, "ab") = 1  '检查IE是否存在
            If IEVersion = 11 Then .Cells(6, "ac") = IEVersion 'IE版本
        End If
        If fso.fileexists(Environ("SYSTEMDRIVE") & "\Program Files\7-Zip\7z.exe") = True Then
            .Cells(32, "ab") = Environ("SYSTEMDRIVE") & "\Program Files\7-Zip\7z.exe" '检查zip是否存在
        Else
            If fso.fileexists(Environ("SYSTEMDRIVE") & "\Program Files (x86)\7-Zip\7z.exe") = True Then .Cells(52, "ab") = Environ("SYSTEMDRIVE") & "\Program Files (x86)\7-Zip\7z.exe" '检查zip是否存在
        End If
        If fso.fileexists(Environ("SYSTEMDRIVE") & "\Program Files\Bandizip\bc.exe") = True Then
            .Cells(52, "ab") = Environ("SYSTEMDRIVE") & "\Program Files\Bandizip\bc.exe" '检查bandzip是否存在
        Else
            If fso.fileexists(Environ("SYSTEMDRIVE") & "\Program Files (x86)\Bandizip\bc.exe") = True Then .Cells(32, "ab") = Environ("SYSTEMDRIVE") & "\Program Files (x86)\Bandizip\bc.exe" '检查zip是否存在
        End If
    End With
End Sub

Function CheckExcel64() As Boolean '判断Excel是x86还是x64
    Dim xi As Long
    
    CheckExcel64 = False
    On Error GoTo ErrHandle
    xi = ThisWorkbook.Application.Hinstance
    Exit Function
ErrHandle:
    CheckExcel64 = True
End Function

