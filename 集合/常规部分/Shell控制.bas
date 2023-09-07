Attribute VB_Name = "Shell控制"
Option Explicit

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
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Sub kdk()
CheckProgramRun 1704
End Sub
Function CheckProgramRun(ByVal pid As Long, Optional ByVal timeoutx As Integer) As Boolean '判断程序是否还在继续运行
    Dim i As Integer, k As Integer
    k = 100
    If timeoutx > 20 Then k = timeoutx '最长的延迟控制
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0, pid)
    Do
        GetExitCodeProcess hProcess, ExitCode
        DoEvents
        Sleep 50
        i = i + 1
    Loop While ExitCode = STILL_ALIVE And i < k
    If ExitCode = STILL_ALIVE Then CheckProgramRun = True
    CloseHandle hProcess
    CheckProgramRun = False
End Function
'-----------------------------------------'其他演示案例(下面内容需要手动修改才能用)
Private Sub Tips() '等待shell执行
Dim pid As Long
Dim ExitEvent As Long
pid = Shell(exe, vbNormalFocus)
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0, pid)

ExitEvent = WaitForSingleObject(hProcess, INFINITE)
Call CloseHandle(hProcess)
End Sub

Private Sub Tips1() '终结进程
Dim pid As Long
pid = Shell("C:\Program Files\Windows NT\Accessories\wordpad.exe", vbNormalFocus)
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0, pid)
If hProcess <> 0 Then
'GetExitCodeProcess hProcess, ExitCode
Debug.Print TerminateProcess(hProcess, 0&)
End If
CloseHandle hProcess
End Sub

Private Sub Tips2() '窗体前置
Dim pid As Long
Dim hwnd5 As Long
pid = Shell("notepad", vbNormalFocus)
hwnd5 = GetForegroundWindow()
Do While IsWindow(hwnd5)
DoEvents
Loop
End Sub


