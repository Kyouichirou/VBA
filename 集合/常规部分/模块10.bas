Attribute VB_Name = "模块10"
Option Explicit

Private bIsWinNT As Boolean
Private Declare Function SHRestartSystemMB Lib "shell32" Alias "#59" (ByVal hOwner As Long, sExtraPrompt As String, ByVal uFlags As Long) As Long
Private Declare Function SHShutDownDialog Lib "shell32" Alias "#60" (ByVal YourGuess As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Const EWX_LOGOFF = 0
Private Const EWX_SHUTDOWN = 1
Private Const EWX_REBOOT = 2
Private Const EWX_FORCE = 4
Private Const EWX_POWEROFF = 8
Private Const shrsExitNoDefPrompt = 1
Private Const shrsRebootSystem = 2
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)



Sub dldlld()
Dim i As Long
For i = 1 To 7
Cells(i, 5) = IsUnicodeStr(Cells(i, 2).Value)
Next
End Sub

Private Function IsWinNT() As Boolean

Dim osvi As OSVERSIONINFO

osvi.dwOSVersionInfoSize = Len(osvi)

GetVersionEx osvi

IsWinNT = (osvi.dwPlatformId = VER_PLATFORM_WIN32_NT)

End Function

Private Function CheckString(msg As String) As String

If bIsWinNT Then

CheckString = StrConv(msg, vbUnicode)

Else: CheckString = msg

End If

End Function

Private Function GetStrFromPtr(lpszStr As Long, nBytes As Integer) As String

ReDim ab(nBytes) As Byte

CopyMemory ab(0), ByVal lpszStr, nBytes

GetStrFromPtr = GetStrFromBuffer(StrConv(ab(), vbUnicode))

End Function

Private Function GetStrFromBuffer(szStr As String) As String

If IsUnicodeStr(szStr) Then szStr = StrConv(szStr, vbFromUnicode)

If InStr(szStr, vbNullChar) Then

GetStrFromBuffer = Left$(szStr, InStr(szStr, vbNullChar) - 1)

Else: GetStrFromBuffer = szStr

End If

End Function




Private Sub Command1_Click()

Call SHShutDownDialog(0)

End Sub

Private Sub Command2_Click()

Dim sPrompt As String

Dim uFlag As Long

Select Case Combo1.ListIndex

Case -1: uFlag = Val(Combo1.Text)

Case 0: uFlag = shrsExitNoDefPrompt

Case 1: uFlag = shrsRebootSystem

End Select

If SHRestartSystemMB(hwnd, sPrompt, uFlag) = vbYes Then

End If

End Sub

Private Sub Form_Load()

bIsWinNT = IsWinNT()

If bIsWinNT Then 'WinNT操作系统

With Combo1

.AddItem "0 - 关闭程序并以其它用户身份登陆"

.AddItem "1 - 关闭计算机"

.AddItem "2 - 重新启动计算机"

.Text = ""

End With

Else 'Win95/98操作系统

With Combo1

.AddItem "1 - 关闭计算机"

.AddItem "2 - 重新启动计算机"

.Text = ""

End With

End If

Command1.Caption = "关闭系统对话框"

Command2.Caption = "关闭或重新启动计算机"

End Sub




