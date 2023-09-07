Attribute VB_Name = "外部程序调用"
Option Explicit
'外部程序的调用需要考虑一系列的干扰因素
'外部程序的路径-涉及到如何清除字符/空格干扰
'外部程序执行结果的返回
'执行结果返回的处理是否同步
'执行目标的路径干扰
'执行失败的处理
'如何等待执行结果的返回, wsh.run支持同步, 需要考虑执行的时间过长的问题,可以搭配shell控制一起使用
'------------------------------主要涉及外部程序:CMD, Powershell, 7Zip, BandZip
'---------Environ("comspec"), cmd的路径
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Function Webbrowser(ByVal url As String, Optional ByVal cmCode As Byte) '调用浏览器
    Dim exepath As String
    
    With ThisWorkbook.Sheets("temp")
        exepath = .Range("ab10").Value
        If Len(exepath) = 0 Or fso.fileexists(exepath) = False Then '不存在其他浏览器
            If Len(.Range("ab6").Value) > 0 Then
                If cmCode = 2 Then
                    Turlx = url
                    UserForm15.Show
                    Exit Function
                End If
                If cmCode = 1 Then UserForm3.Label57.Caption = "此站点不支持IE浏览器": Exit Function
                exepath = Environ("SYSTEMDRIVE") & "\Program Files\Internet Explorer\iexplore.exe" & Chr(32) '如果没有设置其他的浏览器.默认调用IE
            Else
                UserForm3.Label57.Caption = "尚未设置浏览器"
                Exit Function
            End If
        Else
            exepath = exepath & Chr(32)
        End If
    End With
    exepath = """" & exepath & """"
    Shell exepath & url, vbNormalFocus
End Function

Function CheckFileX(ByVal folderp As String) As Boolean '检查文件夹下是否存在目标文件 'https://www.cnblogs.com/zhaoqingqing/p/4620402.html
    Dim strx As String, strx2 As String, strx1 As String, strx3 As String, strx4 As String
    Dim c As Byte, i As Byte, xi As Variant, k As Byte

    CheckFileX = True
    strx4 = ThisWorkbook.Path & "\checklist.txt"
    If fso.fileexists(strx4) = True Then fso.DeleteFile (strx4): Sleep 25 '等待文件删除
    xi = Split(folderp, "\")
    c = UBound(xi)
    If fso.GetDriveName(folderp) <> Environ("SYSTEMDRIVE") Then k = 1 '判断文件的来源是否是系统盘
    strx = Left(folderp, 2) '获取盘符
    i = InStrRev(folderp, "\")
    strx1 = Left(folderp, i - 1)
    strx2 = "cd " & strx1
    strx3 = xi(c)
    If c = 1 Then
        Shell ("cmd /c " & strx & " && for /r " & strx3 & " %i in (*.epub,*.pdf,*.mobi,*.xl*,*.doc*,*.txt,*.pp*,*.ac*) do @echo %i >" & strx4), vbHide
    ElseIf k = 0 And c > 1 Then
        Shell ("cmd /c " & strx2 & " && for /r " & strx3 & " %i in (*.epub,*.pdf,*.mobi,*.xl*,*.doc*,*.txt,*.pp*,*.ac*) do @echo %i >" & strx4), vbHide
    ElseIf k = 1 And c > 1 Then
        Shell ("cmd /c " & strx & " && strx2 " & " && for /r " & strx3 & " %i in (*.epub,*.pdf,*.mobi,*.xl*,*.doc*,*.txt,*.pp*,*.ac*) do @echo %i >" & strx4), vbHide
    End If
    Sleep 200 '等待文件生成
    If fso.fileexists(strx4) = False Then CheckFileX = False '没有目标文件
End Function

Function ZipCompress(ByVal xpath As String, Optional ByVal cmbagname As String, Optional ByVal cmCode As Byte, Optional ByVal passwordx As String) '调用7zip来压缩文件 '路径'压缩包名称'压缩的类型'压缩设置的密码
    Dim exepath As String, ax As Variant, i As Byte
    Dim strx As String, strx1 As String, strx2 As String, strx3 As String, strx4 As String  '调用外部程序时,不管是cmd还是Powershell,符号需要注意
    
'    exepath = "C:\Program Files\7-Zip\7z.exe" '如果安装到其他位置
    exepath = ThisWorkbook.Sheets("temp").Cells(32, "ab").Value
    If fso.fileexists(exepath) = False Then MsgBox "尚未设置解压软件", vbOKOnly, "Tips": Exit Sub
    exepath = exepath & Chr(32)
    exepath = """" & exepath & """"
    i = InStrRev(xpath, "\") - 1
    If Len(cmbagname) = 0 Then '压缩包的名字
        If InStr(xpath, ".") > 0 Then '表示这是文件
            ax = Split(xpath, "\")
            strx1 = ax(UBound(ax))
            strx1 = Split(strx1, ".")(0)
            cmbagname = strx1
        Else  '这是文件夹
            i = i + 1
            cmbagname = Right(xpath, Len(xpath) - i) '获取文件夹名
            i = i - 1
            xpath = xpath & "\"
        End If
    End If
    
    strx2 = Left(xpath, i) '如果压缩的是文件就进入文件所在的文件夹下, 如果压缩的是文件夹就返回上一级文件夹
    cmbagname = cmbagname & ".7z" '使用7z作为压缩包的格式
    cmbagname = """" & cmbagname & """"
    xpath = """" & xpath & """"
     '不管是什么程序的外部调用,都需要注意消除变量当中的干扰项, 如空格,单引号,不可见字符等等这些干扰因素
    Select Case cmCode '压缩 '压缩后删除 '压缩后加密 '压缩后加密并删除
        Case 0
        strx3 = exepath & " a " & cmbagname & " " & xpath
        Case 1
        strx3 = exepath & " a " & "-sdel " & cmbagname & " " & xpath 'a表示add添加文件到压缩包
        Case 2
        strx3 = exepath & " a " & "-p" & passwordx & " " & cmbagname & " " & xpath
        Case 3
        If Len(passwordx) = 0 Then passwordx = "password" '如果选择3,但有没有设置密码,默认将密码设置为"password"
        strx3 = exepath & " a " & "-p" & passwordx & " " & "-sdel " & cmbagname & " " & xpath
        Case Else
        Exit Function
    End Select
    '解压文件的时候采用的方法也是类似的(包括加密文件的解压), 只是命令为" e "
    '可以拓展一下,将程序改写成破解压缩包的小程序,当然密码长度和密码使用的字符要有限制,暴力破解密码难度随着密码的复杂度成级数上升(3位数为例)10*10*10=1000 * 46.656/36*36*36=46656 *5.1=/62*62*62=238 328
    strx4 = "cd " & strx2 '进入所在目录
                         'cmd执行多个命令时,采用 "&&" 作为连接符 '注意这里和直接调用7zip有所区别
    Shell ("cmd /c " & strx4 & "&&" & strx3), vbHide '调用7zip进行压缩文件,a表示add,添加文件大压缩包, -sdel,压缩之后删除掉源文件 ,-p表示设置密码 '更多相关信息可查看7z安装目录下的chm文件
    '整理合并后可以将相关的功能合并成类模块方便调用,将压缩,sha1计算等
End Function

Function FileisOpen(ByVal FilePath As String) As Boolean  '调用powershell脚本来判断文件是否处于打开的状态,支持非ansi字符路径文件'启用powershell执行ps脚本的命令'需要管理权限'以管理员身份执行powershell 输入: set-executionpolicy remotesigned 关闭命令: Set-ExecutionPolicy Restricted
    Dim strOutput As String, Jsfilepath As String
    Dim WshShell As Object, WshShellExec As Object
    '-------------------------------------------------------------需要注意txt文件是不锁定的在打开的状态,所以可以自由控制txt文档
    Jsfilepath = ThisWorkbook.Sheets("temp").Range("ab7").Value
    If Len(Jsfilepath) = 0 Or PSexist = False Then Exit Function
    strCommand = "Powershell.exe -ExecutionPolicy ByPass " & Jsfilepath & Chr(32) & FilePath  'chr(32) 表示空格符 '注意strcommand参数中的这个空格符
    Set WshShell = CreateObject("WScript.Shell")
    Set WshShellExec = WshShell.Exec(strCommand)
    strOutput = WshShellExec.StdOut.ReadAll     '返回执行的结果,如果文件处于打开的状态就有返回值,否则就是空值
    If Len(strOutput) > 0 Then
        FileisOpen = True
    Else
        FileisOpen = False
    End If
    Set WshShell = Nothing
    Set WshShellExec = Nothing
End Function

Function HashPowershell(filepaths As String) As String '使用powershell来计算md5,注意不同方式对于符号的区别使用 'https://docs.microsoft.com/en-us/windows-server/administration/windows-commands/certutil
    Dim wsh As Object
    Dim wExec As Object
    Dim exepath As String
    Dim filepathx As String
    Dim Result As String, i As Byte, strx As String
    
    Set wsh = CreateObject("WScript.Shell")
     '这里不同的执行代码需要注意的是文件路径带有某些特殊的字符造成代码无法被执行或者执行出错的问题,假设文件路径存在'空格()"等非常有干扰的字符
    If InStr(filepaths, Chr(39)) > 0 Then
        filepathx = """" & filepaths & """"
        Set wExec = wsh.Exec("powershell certutil -hashfile " & """" & """" & filepathx & """" & """" & " md5")
    Else
        filepathx = "'" & filepaths & " '"
        Set wExec = wsh.Exec("powershell certutil -hashfile " & filepathx & " MD5")
    End If
    'Set wExec = wsh.Exec("powershell Get-FileHash " & filepathx & " -Algorithm MD5| Format-List")
    Result = wExec.StdOut.ReadAll
    If Len(Result) = 0 Then '没有获取到hash
        HashPowershell = "UC"
        Exit Function
    End If
    i = UBound(Split(Result, ":")) - 1
    strx = Replace(Replace(Split(Result, ":")(i), Chr(10), ""), " ", "")
    HashPowershell = Left$(strx, Len(strx) - 8) '1方式获取值的方式
    'Hashpowershell = UCase(Trim(Replace(Replace(Split(Split(Result, "Hash")(1), "Path")(0), ":", ""), Chr(10), ""))) 'chr10表示换行符
    Set wExec = Nothing
    Set wsh = Nothing
End Function

Function ZipHash(ByVal FilePath As String) As String '使用7zip来计算sha1 or sha256/sha1等 不支持md5-备用 'https://www.7-zip.org/
    Dim wsh As Object
    Dim wExec As Object
    Dim exepath As String, filepathx As String
    Dim filemdx As String, Result As String '更多的相关的内容可以参阅zip安装目录下的chm文档
    
'    exepath = "C:\Program Files\7-Zip\7z.exe "
    exepath = ThisWorkbook.Sheets("temp").Cells(32, "ab").Value
    If fso.fileexists(exepath) = False Then MsgBox "尚未设置解压软件", vbOKOnly, "Tips": Exit Sub
    exepath = exepath & Chr(32)
    If fso.fileexists(Trim(exepath)) = False Or fso.fileexists(Trim(FilePath)) = False Then Exit Function
    filepathx = """" & FilePath & """"                          '注意这里的符号的使用,如何表示一个变量
    Set wsh = CreateObject("WScript.Shell")
    Set wExec = wsh.Exec(exepath & "h -scrcsha256 " & filepathx) '此方法的缺陷在于无法隐藏执行窗口
    Result = wExec.StdOut.ReadAll
    If Len(Result) = 0 Then
    ZipHash = "UC"
    Set wExec = Nothing
    Set wsh = Nothing
    Exit Function
    End If
    filemdx = Trim(Split(Result, "SHA256 for data:")(UBound(Split(Result, "SHA256 for data:"))))
    ZipHash = Replace(Left$(filemdx, Len(filemdx) - 18), Chr(10), "") 'chr(10)表示ascii码的换行符
    Set wExec = Nothing
    Set wsh = Nothing
End Function

Function PowerSHForceW(ByVal FilePath As String, ByVal filez As Long) '通过Powershell强制向受保护的文件写入内容
    Dim ws As Object
    Dim strx As String, commandline As String, strx1 As String
    'Powershell其他写入内容的方法不管是out-file 还是system.IO的方式都无法直接向受保护的文件写入数据
    filez = filez \ 2
    If filez > 104857600 Then filez = 104857600 '上限100M
    PowerSHCreate filez
    strx = ThisWorkbook.Path & "\temp.txt"
    strx = "'" & strx & "'"
    If InStr(FilePath, Chr(39)) = 0 Then
        FilePath = "'" & FilePath & "'"
        strx1 = FilePath
    Else
        FilePath = """" & FilePath & """"
        strx1 = """" & """" & FilePath & """" & """"
    End If
    Set ws = CreateObject("wscript.shell")
    'https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/get-content?view=powershell-6
    'https://docs.microsoft.com/en-us/powershell/module/Microsoft.PowerShell.Utility/Out-File?view=powershell-6
    commandline = "powershell $bulk=" & strx & ";$file=" & strx1 & ";(Get-Content $bulk) | Set-Content $file -Encoding Unknown -Force" & ";Remove-Item $bulk"
    '将生成的数据填充进文件,100M的文件趋近于上限了(如果是单独的ps1脚本就无此问题)
    ws.Run (commandline), 0, True
    Set ws = Nothing
    Sleep 50
End Function

Function PowerSHCreate(ByVal filez As Long) '调用powershell快速生成一个任意大小的文件
    Dim TempFile As String, commandline As String
    Dim ws As Object
    'https://docs.microsoft.com/en-us/dotnet/api/system.io.streamwriter?view=netframework-4.8
    TempFile = ThisWorkbook.Path & "\temp.txt"
    TempFile = "'" & TempFile & "'"
    commandline = "$tempFile=" & TempFile & _
    ";$fs=New-Object System.IO.FileStream($tempFile,[System.IO.FileMode]::OpenOrCreate)" & _
    ";$fs.Seek(" & filez & " ,[System.IO.SeekOrigin]::Begin)" & _
    ";$fs.WriteByte(1)" & _
    ";$fs.Close()"
    Set ws = CreateObject("wscript.shell")
    ws.Run ("powershell " & commandline), 0, True
    Set ws = Nothing
    Sleep 50
End Function

Sub CmdOpenFile(ByVal FilePath As String) '通过cmd间接打开文件
    '--------------------------------需要有默认打开文件的管理关联程序
    FilePath = """" & FilePath & """"
    Shell ("cmd /c" & FilePath), vbHide
End Sub

Function PowerSHOpen(ByVal FilePath As String) '利用powershell 打开文件
    Dim ws As Object
    Dim commandline As String
    
    If InStr(FilePath, Chr(39)) > 0 Then
        FilePath = """" & FilePath & """"
        commandline = "powershell " & "invoke-item " & """" & """" & FilePath & """" & """"
    Else
        FilePath = "'" & FilePath & "'"
        commandline = "powershell " & "invoke-item " & FilePath
    End If
    Set ws = CreateObject("wscript.shell")
    ws.Run (commandline), 0
    Set ws = Nothing
End Function

Sub ZipExtract(ByVal FilePath As String, ByVal Folderpath As String) '由于7zip这坑神不支持在cmd commandline下输入密码,不得不换成bandzip
    Dim i As Byte
    Dim strx As String, strx1 As String, exe As String
    Dim wsh As Object
    '和7zip相比,貌似其命令行出现轻微的差异,在命令参数的顺序上,不过整体的参数完全是相同的
    Set wsh = CreateObject("WScript.Shell")
'    exe = "C:\Program Files\Bandizip\bc.exe "
    exe = ThisWorkbook.Sheets("temp").Cells(52, "ab").Value
    If fso.fileexists(exe) = False Then MsgBox "尚未设置解压软件", vbOKOnly, "Tips": Exit Sub
    exe = exe & Chr(32)
    strx = """" & FilePath & """"
    strx1 = """" & Folderpath & """"
    exe = """" & exe & """"
    wsh.Run (exe & " e " & "-aoa " & strx & " -o " & strx1), vbNormalFocus '支持在命令行中输入密码
    Set wsh = Nothing
End Sub

Function TerminateEXE(ByVal exename As String, Optional ByVal cmCode As Byte = 0) As Byte '终止特定进程
    Dim obj As Object, targetexe As Object, targetexex As Object
                                                            'https://docs.microsoft.com/en-us/windows/win32/wmisdk/wmi-tasks--processes
    On Error GoTo 100
    TerminateEXE = 0
    Set obj = GetObject("winmgmts:\\.\root\cimv2")          'https://docs.microsoft.com/en-us/windows/win32/wmisdk/connecting-to-wmi-with-vbscript
    If obj Is Nothing Then MsgBox "无法创建对象": Exit Function
    Set targetexe = obj.ExecQuery("select * from win32_process where name=" & Chr(39) & exename & Chr(39)) '("select * from win32_process where name='iexplore.exe'")' chr(39)='
    If targetexe Is Nothing Then Exit Function
    For Each targetexex In targetexe
    If cmCode = 1 Then
        targetexex.Terminate
    Else
        TerminateEXE = 1: Exit For
    End If
    Next
    Set obj = Nothing
    Set targetexe = Nothing
    Exit Function
100
    Set obj = Nothing
    Set targetexe = Nothing
    Err.Clear
End Function

Sub CopyFileClipboard(ByVal FilePath As String) '通过powershell复制文件到剪切板
    Dim commandline As String, strx As String
    
    FilePath = """" & FilePath & """"   '多文件:'filepath = "'C:\text.txt','D:\text.txt','D:\text.jpg'"
    strx = """" & """" & FilePath & """" & """"
    commandline = "powershell $filelist =" & strx & vbCrLf & _
    "$col = New-Object Collections.Specialized.StringCollection " & vbCrLf & _
    "foreach($file in $filelist){$col.add($file)}" & vbCrLf & _
    "Add-Type -AssemblyName System.Windows.Forms" & vbCrLf & _
    "[Windows.Forms.Clipboard]::setfiledroplist($col)"
    Shell commandline, vbHide
End Sub

Sub cmdFindText(ByVal FilePath As String, ByVal Keyword As String) '用cmd查找txt文本内带有关键词的行
    Keyword = """" & Keyword & """"
    FilePath = """" & FilePath & """"
    Shell "cmd /c" & "findstr /c:" & strx & " " & FilePath & " | clip", vbHide '配合shell控制一起使用
    '---------------------------将找到的内容输出到剪切板
End Sub
