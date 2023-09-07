Attribute VB_Name = "删除本地文件"
Option Explicit
Private Type uSHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As Long   'unicode字符串的地址(必须是包含unicode)
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long
End Type
Private Type aSHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String  '字符串
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long
End Type
'https://docs.microsoft.com/zh-cn/windows/win32/api/shellapi/nf-shellapi-shfileoperationw
'注意这里的A和W的区别, A表示参数传递进来的时候将被转换为ansi编码, w表示接受unicode编码
'https://blog.csdn.net/Giser_D/article/details/103311433
'W为唯一,CreateProcessA会做ansi到wide的转换，然后调底层函数。底层函数只有wide的版本。
Private Declare Function uSHFileOperation Lib "shell32.dll" Alias "SHFileOperationW" (ByRef lpFileOp As uSHFILEOPSTRUCT) As Long ' 用于删除包含unicode字符的文件
Private Declare Function aSHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (ByRef lpFileOp As aSHFILEOPSTRUCT) As Long
Private Const FO_Delete = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Const HWND_DESKTOP = 0
Private Const NOCONFIRMATION = &H10 '不提示
'varPtr是指变量所在处的内存地址，不管你的变量str有没有赋值，这个内存地址是不变的~
'strPtr是指字符串的地址（也可认为是值的地址），当你的变量str刚声明时，他空间中0，没有任何值的指向，所以strPtr(str)=0，
'当给变量str赋值abc后（即str=“abc”），strPtr(str)也就是指向了其字符串（也就是他的值）的地址~
'一句话，变量确定后，其varPtr就是固定不变的，但其strPtr会随着他的值的变化而变化，因为值变了，它指向的地址就变了~
Function DeleteFiles(ByVal FilePath As String, ByVal isUnicode As Boolean) '支持ansi, 也支持删除包含unicode字符的文件
    Dim aSHdele As aSHFILEOPSTRUCT
    Dim uSHdele As uSHFILEOPSTRUCT
    If isUnicode = False Then
        With aSHdele
            .hwnd = HWND_DESKTOP
            .pTo = ""
            .wFunc = FO_Delete
            .pFrom = Path + Chr(0)
            .fFlags = FOF_ALLOWUNDO + NOCONFIRMATION
        End With
        aSHFileOperation aSHdele
    Else
        With uSHdele
            .hwnd = HWND_DESKTOP
            .pTo = ""
            .wFunc = FO_Delete
            .pFrom = StrPtr(Path)    '关键在于这一步
            .fFlags = FOF_ALLOWUNDO + NOCONFIRMATION
        End With
        uSHFileOperation uSHdele
    End If
End Function

Function DeleFilepShell(ByVal strx As String) '调用powershell命令来删除文件(支持非ansi编码字符),删除文件支持到回收站 '打开的文件不会出错,会提示(备用删除工具)
    Dim ws As Object
    Dim filepathx As String
    
    If fso.fileexists(strx) = False Then Exit Function '文件不存在,直接退出
    filepathx = "'" & strx & " '"               'powershell变量采用 " '(单引号) " 符号比 " "(双引号) " 符号更好
    Set ws = CreateObject("wscript.shell")
    ws.Run ("powershell $testFile=" & filepathx & """;dir | Out-File $testFile;$shell = new-object -comobject 'Shell.Application';$item = $shell.Namespace(0).ParseName( (Resolve-Path $testFile).Path);$item.InvokeVerb('delete')"""), 0
    '0表示隐藏窗口'powershell调用多行命令的方法,使用 ";"(分号)或者 "|"
    Set ws = Nothing
End Function

Function DeleToRecycle(ByVal FilePath As String) As Boolean '移除文件到回收站-备用 'boolean将用于判断删除命令是否得到实际的执行(可能存在文件被其他程序占用而出现无法删除的问题)
    Dim objReg As Object
    Dim objShell As Object
    Dim vStateArr As Variant, vBackupState As Variant
    '---------------------------https://docs.microsoft.com/en-us/previous-versions/tn-archive/ee176985(v=technet.10)?redirectedfrom=MSDN
    On Error GoTo 100
    Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    objReg.GetBinaryValue &H80000001, "Software\Microsoft\Windows\CurrentVersion\Explorer", "ShellState", vStateArr
    vBackupState = vStateArr
    vStateArr(4) = 39
    '-------------------修改注册表,禁用弹窗提示
    objReg.SetBinaryValue &H80000001, "Software\Microsoft\Windows\CurrentVersion\Explorer", "ShellState", vStateArr
    Set objShell = CreateObject("Shell.Application")
    objShell.Namespace(0).ParseName(FilePath).InvokeVerb ("delete") '删除执行
    objReg.SetBinaryValue &H80000001, "Software\Microsoft\Windows\CurrentVersion\Explorer", "ShellState", vBackupState
    '-------------------------恢复注册表
    DeleToRecycle = True '删除命令被执行
    Set objReg = Nothing
    Set objShell = Nothing
100
    Exit Function
    If Err.Number <> 0 Then
        Set objReg = Nothing
        Set objShell = Nothing
        Err.Clear
        DeleToRecycle = False
    End If
End Function

Sub EmptyRecycle() '清空回收站
    Dim retVal As Long
    retVal = SHEmptyRecycleBin(0&, vbNullString, SHERB_NORMAL)
End Sub

Sub SendFile2Recycle(ByVal FilePath As String) '支持unicode字符,缺点,出现确认删除弹窗, 需要结合注册表取消弹窗
    '-----------------------------https://docs.microsoft.com/zh-cn/windows/win32/shell/invokeverbex
    Dim strFolderParent As Variant
    Dim strFileName As Variant
    Dim objShell As Object
    Dim objFolder As Object
    Dim objFolderItem As Object
    
    strFolderParent = fso.GetParentFolderName(FilePath)
    strFileName = fso.GetFileName(FilePath)
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.Namespace(strFolderParent)
    Set objFolderItem = objFolder.ParseName(strFileName)
    objFolderItem.InvokeVerbEx ("Delete")
    Set objShell = Nothing
    Set objFolder = Nothing
    Set objFolderItem = Nothing
End Sub
