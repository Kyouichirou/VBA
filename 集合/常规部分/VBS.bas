Attribute VB_Name = "VBS"
Option Explicit
'https://docs.microsoft.com/en-us/windows-server/administration/windows-commands/cscript
'https://docs.microsoft.com/en-us/windows-server/administration/windows-commands/wscript
Dim WshShell As Object
Dim i As Integer
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) '时间 -单词训练

Sub VBSmultiThread(ByVal Folderpath As String) '利用VBS来实现"多线程"实现更快计算文件md5 ,计算速度从1个文件340M/s,上升至800M(57个文件)/s(速度和硬件性能挂钩)
    Dim t As Single
    Dim fd As Folder
    Dim yesno As Variant
    
    t = Timer
    i = 0
    '----------检测脚本程序是否处于运行的状态, 先结束掉进程, 以免影响后续的判断
    If TerminateEXE("cscript.exe") = 1 Then
        yesno = MsgBox("使用前,需要关闭cscript.exe", vbYesNo, "Tips")
        If yesno = vbNo Then Exit Sub
        TerminateEXE "cscript.exe", 1
    End If
    Set WshShell = CreateObject("Wscript.Shell")
    If fso.folderexists(Folderpath) = False Then Exit Sub
    Set fd = fso.GetFolder(Folderpath)
    Search_Files fd
    Do
        Sleep 25
        If TerminateEXE("cscript.exe") = 0 Then Exit Do '判断脚本命令是否完全执行完毕
    Loop
    '--------读取写入txt文档的内容即可
    Debug.Print Timer - t, i
    Set fd = Nothing
    Set WshShell = Nothing
End Sub

Private Sub Search_Files(ByVal fd As Folder)
    Dim fl As File
    Dim sfd As Folder
    For Each fl In fd.Files
        If fl.Size > 0 Then
            If fl.Attributes <> 34 Then i = i + 1: CallVBS fl.Path, Right("000" & CStr(i), 4) '生成txt文档的文件名(长度4), 方便后续的数据的读取和使用
        End If
    Next
    If fd.SubFolders.Count = 0 Then Exit Sub
    For Each sfd In fd.SubFolders
    Testfile sfd
    Next
End Sub

Private Function CallVBS(ByVal FilePath As String, ByVal namex As String)
    Dim commandline As String
    Dim vbspath As String
    Dim strx As String
    '-----------------------https://ss64.com/vb/cscript.html
    strx = FilePath & namex
    vbspath = "C:\Users\adobe\Desktop\MD5Hash.vbs"
    commandline = vbspath & """" & strx & """"
    WshShell.Run """" & vbfilecm & """" '这里不要设置为同步(true)
End Function

'--------------------------------------------------------vbs脚本文件的内容(MD5A模块的vbs版本),需要注意的是,空格在传递参数时表示多个参数
'Dim fso, fl, flop, strx, strxa, strfolder, strxb, strxc
'strx = ""
'For Each strxa In WScript.Arguments '将参数合并成字符串
'strx = strx & strxa & " " '-----------这里还需要测试多个空格存在于路径的情况
'Next
'Set fso = CreateObject("Scripting.FileSystemObject")
'strfolder = "C:\Users\adobe\Documents\lbrecord\Hash\"
'strx = Left(strx, Len(strx) - 1) '传递两个参数进来
'strxb = Right(strx, 4)
'strx = Left(strx, Len(strx) - 4)
'strxc = strfolder & strxb & ".txt"
'fso.CreateTextFile strxc, True '无法直接创建file对象使用openstream
'Set fl = fso.GetFile(strxc)
'Set flop = fl.OpenAsTextStream(8, -2) 'TristateUseDefault 'ForAppending
'flop.WriteLine GetFileHashMD5(strx) 'WScript.Arguments(0)
'flop.Close
'Set fso = Nothing
'Set fl = Nothing
'Set flop = Nothing
'
'Function GetFileHashMD5(filepath) '计算md5速度最好的方法,支持非ansi字符路径
'    Dim Filehashx
'    Dim WDI                                        '可以计算大于2G以上的文件,单个文件测试最大12+G
'    Dim HashValue
'    Dim i       'https://docs.microsoft.com/zh-cn/windows/win32/msi/msifilehash-table
'    Dim k, j
'
'    On Error Resume Next '出错直接调用其他的模块
'    Set WDI = CreateObject("WindowsInstaller.Installer")
'    If WDI Is Nothing Then MsgBox Err.Number
'    Set Filehashx = WDI.FileHash(filepath, 0)           '核心
'    If WDI Is Nothing Or Filehashx Is Nothing Then GetFileHashMD5 = "UC1" & filepath: Set WDI = Nothing: Set Filehashx = Nothing: Exit Function
'    k = Filehashx.FieldCount '4
'    For i = 1 To k
'        HashValue = HashValue & BigEndianHex(Filehashx.IntegerData(i))
'    Next
'    GetFileHashMD5 = HashValue
'    j = Len(GetFileHashMD5)
'    If j <> 32 And j <> 2 Then GetFileHashMD5 = "UC2" & filepath
'    Set Filehashx = Nothing
'    Set WDI = Nothing
'    If Err.Number > 0 Then GetFileHashMD5 = "UC3" & filepath: Err.Clear
'End Function
'
'Function BigEndianHex(xl) 'https://blog.csdn.net/weixin_42066185/article/details/83755433
'    Dim Result
'    Dim strx1, strx2, strx3, strx4
'    '-------------------------------------https://stackoverrun.com/ja/q/8312292
'    '-----------https://docs.microsoft.com/zh-CN/office/vba/api/excel.application.worksheetfunction
'    '-----------https://docs.microsoft.com/zh-CN/office/vba/api/excel.worksheetfunction.dec2hex
'    '-----------Result = ThisWorkbook.Application.WorksheetFunction.Dec2Hex(xl, 8) '会出现超出8位的问题
'    Result = Hex(xl)
'    If Len(Result) < 8 Then Result = Right("00000000" & Result, 8) '补位
'    strx1 = Mid(Result, 7, 2)
'    strx2 = Mid(Result, 5, 2)
'    strx3 = Mid(Result, 3, 2)
'    strx4 = Mid(Result, 1, 2)
'    BigEndianHex = strx1 & strx2 & strx3 & strx4
'End Function
'-----------------------------------------------vbs
