Attribute VB_Name = "CMD"
Option Explicit
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type
'------------------------https://www.pinvoke.net/default.aspx/advapi32.createprocessasuser
Private Declare Function CreateProcessAsUser Lib "Advapi32.dll" Alias "CreateProcessAsUserA" (ByVal hToken As Long, _
ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As SECURITY_ATTRIBUTES, _
ByVal lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
ByVal lpEnvironment As String, ByVal lpCurrentDirectory As String, ByVal lpStartupInfo As STARTUPINFO, ByVal lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, _
lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, _
ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const NORMAL_PRIORITY_CLASS = &H20
Private Const STARTF_USESTDHANDLES = &H100
Private Const STARTF_USESHOWWINDOW = &H1
'除了这种方法,也可以利用cmd的管道来实现结果的取回, 如将结果输出到剪切板来实现结果取回, cmd的管道的符号 "|" ,如ping www.baidu.com | clip, 即可将结果输出到剪切板
Private Function ExecuteCommandLineOutput(commandline As String, Optional BufferSize As Long = 256, Optional TimeOut As Long) As String 'cmd命令执行结果返回
    Dim Proc As PROCESS_INFORMATION 'https://docs.microsoft.com/en-us/windows/win32/api/processthreadsapi/ns-processthreadsapi-process_information
    Dim Start As STARTUPINFO        'http://chokuto.ifdef.jp/urawaza/struct/STARTUPINFO.html
    Dim sA As SECURITY_ATTRIBUTES   'https://docs.microsoft.com/zh-cn/previous-versions/windows/desktop/legacy/aa379560(v=vs.85)
    Dim hReadPipe As Long
    Dim hWritePipe As Long
    Dim lBytesRead As Long
    Dim sBuffer As String
    Dim BeginTime As Date
    
    If Len(commandline) > 0 Then
        With sA
            .nLength = Len(sA)
            .bInheritHandle = 1&
            .lpSecurityDescriptor = 0&
        End With
        If CreatePipe(hReadPipe, hWritePipe, sA, 0) > 0 Then
            With Start
                .cb = Len(Start)
                .dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
                .hStdOutput = hWritePipe
                .hStdError = hWritePipe
            End With
            If CreateProcessA(0&, commandline, sA, sA, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, Start, Proc) = 1 Then
                CloseHandle hWritePipe
                sBuffer = String(BufferSize, Chr(0))
                'chr(0)不是null,null是什么都没有，而chr(0)的值是0。表示成16进制是0x00，表示成二进制是00000000,虽然chr(0)不会显示出什么，但是他是一个字符
                '当汉字被截断时，根据编码规则他总是要把后边的其他字符拉过来一起作为汉字解释，这就是出现乱码的原因。而值为0x81到0xff与0x00组合始终都显示为“空”,根据这一特点，在substr的结果后面补上一个chr(0)，就可以防止出现乱码了
                If TimeOut > 0 Then BeginTime = Now
                Do Until ReadFile(hReadPipe, sBuffer, BufferSize, lBytesRead, 0&) = 0
                    DoEvents
                    If TimeOut > 0 Then '设置了超时的时间
                        If DateDiff("s", BeginTime, Now) > TimeOut Then '这里可以改成其他的时间设置 如 stopwatch
                            ExecuteCommandLineOutput = "Timeout"
                            Exit Do
                        End If
                    End If
                    ExecuteCommandLineOutput = ExecuteCommandLineOutput & Left(sBuffer, lBytesRead)
                Loop
                CloseHandle Proc.hProcess
                CloseHandle Proc.hThread
                CloseHandle hReadPipe
            Else
                ExecuteCommandLineOutput = "File or command not found"
            End If
        Else
            ExecuteCommandLineOutput = "CreatePipe failed. Error: " & Err.LastDllError & "."
        End If
    End If
End Function



