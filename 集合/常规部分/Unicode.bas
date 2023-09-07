Attribute VB_Name = "Unicode"
Option Explicit
'https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-istextunicode
Private Declare Function IsTextUnicode Lib "advapi32" (lpBuffer As Any, ByVal cb As Long, lpi As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Const IS_TEXT_UNICODE_ASCII16 = &H1
Private Const IS_TEXT_UNICODE_REVERSE_ASCII16 = &H10
Private Const IS_TEXT_UNICODE_STATISTICS = &H2
Private Const IS_TEXT_UNICODE_REVERSE_STATISTICS = &H20
Private Const IS_TEXT_UNICODE_CONTROLS = &H4
Private Const IS_TEXT_UNICODE_REVERSE_CONTROLS = &H40
Private Const IS_TEXT_UNICODE_SIGNATURE = &H8
Private Const IS_TEXT_UNICODE_REVERSE_SIGNATURE = &H80
Private Const IS_TEXT_UNICODE_ILLEGAL_CHARS = &H100
Private Const IS_TEXT_UNICODE_ODD_LENGTH = &H200
Private Const IS_TEXT_UNICODE_DBCS_LEADBYTE = &H400
Private Const IS_TEXT_UNICODE_NULL_BYTES = &H1000
Private Const IS_TEXT_UNICODE_UNICODE_MASK = &HF
Private Const IS_TEXT_UNICODE_REVERSE_MASK = &HF0
Private Const IS_TEXT_UNICODE_NOT_UNICODE_MASK = &HF00
Private Const IS_TEXT_UNICODE_NOT_ASCII_MASK = &HF000

'此函数很不准, 而且对字符有长度要求, 最起码长度大于1
Function IsUnicodeStr(sBuffer As String) As Long
Dim dwRtnFlags As Long
dwRtnFlags = IS_TEXT_UNICODE_UNICODE_MASK
Dim arr() As Byte
arr = sBuffer
Dim k As Long
k = StrPtr(sBuffer)
IsUnicodeStr = IsTextUnicode(ByVal k, LenB(sBuffer), dwRtnFlags) '×??°??????
End Function
