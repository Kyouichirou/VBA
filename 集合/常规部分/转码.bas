Attribute VB_Name = "转码"
' @status: clear
' @level: important
' @description: 处理转码问题, 特别是将unicode的vba字符串转为uft-8
' @ document
' https://docs.microsoft.com/en-us/windows/win32/api/Stringapiset/nf-stringapiset-widechartomultibyte
' https://support.microsoft.com/zh-cn/help/138813/how-to-convert-from-ansi-to-unicode-unicode-to-ansi-for-ole
' https://docs.microsoft.com/en-us/windows/win32/intl/unicode
Option Explicit
' 在64位下的api的声明方式, 增加PtrSaf关键词
#If Win64 And VBA7 Then
    Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32 " (ByVal CodePage As Long, _
            ByVal dwFlags As Long, _
            ByVal lpMultiByteStr As LongPtr, ByVal cchMultiByte As Long, _
            ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long _
            )       As Long
    Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32 " (ByVal CodePage As Long, _
            ByVal dwFlags As Long, _
            ByVal lpWideCharStr As LongPtr, _
            ByVal cchWideChar As Long, _
            ByVal lpMultiByteStrPtr As LongPtr, _
            ByVal cchMultiByte As Long, _
            ByVal lpDefaultChar As LongPtr, _
            ByVal lpUsedDefaultChar As LongPtr _
            )       As Long
#Else
    ' 32位下的api声明发横
    Private Declare Function MultiByteToWideChar Lib "kernel32 " (ByVal CodePage As Long, _
            ByVal dwFlags As Long, _
            ByVal lpMultiByteStr As Long, _
            ByVal cchMultiByte As Long, _
            ByVal lpWideCharStr As Long, _
            ByVal cchWideChar As Long _
            )       As Long
    Private Declare Function WideCharToMultiByte Lib "kernel32 " (ByVal CodePage As Long, _
            ByVal dwFlags As Long, _
            ByVal lpWideCharStr As Long, _
            ByVal cchWideChar As Long, _
            ByVal lpMultiByteStr As Long, _
            ByVal cchMultiByte As Long, _
            ByVal lpDefaultChar As Long, _
            ByVal lpUsedDefaultChar As Long _
            )       As Long
#End If
' default to ANSI code page
Private Const CP_ACP As Byte = 0
' default to UTF-8 code page       
Private Const CP_UTF8 As Long = 65001

Function EncodeToBytes(ByVal sData As String) As Byte()
    '字符转 UTF8

    Dim aRetn()     As Byte
    Dim nSize       As Long
    
    If Len(sData) = 0 Then Exit Function
    nSize = WideCharToMultiByte(CP_ACP, 0, StrPtr(sData), -1, 0, 0, 0, 0) - 1
    If nSize = 0 Then Exit Function
    ReDim aRetn(0 To nSize - 1) As Byte
    WideCharToMultiByte CP_ACP, 0, StrPtr(sData), -1, VarPtr(aRetn(0)), nSize, 0, 0
    EncodeToBytes = aRetn
    Erase aRetn
End Function

Function DecodeToBytes(ByVal sData As String) As Byte()

    '解码

    Dim aRetn()     As Byte
    Dim nSize       As Long
    
    If Len(sData) = 0 Then Exit Function
    nSize = MultiByteToWideChar(CP_UTF8, 0, StrPtr(sData), -1, 0, 0) - 1
    If nSize = 0 Then Exit Function
    ReDim aRetn(0 To 2 * nSize - 1) As Byte
    MultiByteToWideChar CP_UTF8, 0, StrPtr(sData), -1, VarPtr(aRetn(0)), nSize
    DecodeToBytes = aRetn
    Erase aRetn
End Function