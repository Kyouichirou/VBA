Attribute VB_Name = "复制文件"
Option Explicit
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
'----------------------------------https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-globallock
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal flags As Long, ByVal Size As Long) As Long
'------------------------------------https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-globalalloc
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private Declare Function RegisterClipboardFormat Lib "user32.dll" Alias "RegisterClipboardFormatW" (ByVal lpString As Long) As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long   '----------------剪切板控制

Private Const CF_HDROP As Long = 15&
Private Const DROPEFFECT_COPY As Long = 1
Private Const DROPEFFECT_MOVE As Long = 2
Private Const GMEM_ZEROINIT As Long = &H40
Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_DDESHARE As Long = &H2000

Private Type dropFiles
    pFiles  As Long
    pt      As POINTAPI
    fNC     As Long
    fWide   As Long
End Type

Function CutOrCopyFiles(FileList As Variant, Optional ByVal CopyMode As Boolean = True) As Boolean 'API的方式复制文件 ,支持非ansi
    Dim uDropEffect As Long, i As Long
    Dim dropFiles   As dropFiles
    Dim uGblLen     As Long, uDropFilesLen  As Long
    Dim hGblFiles   As Long, hGblEffect As Long
    Dim mPtr        As Long
    Dim FileNames   As String

If OpenClipboard(0) Then
    EmptyClipboard
    FileNames = GetFileListString(FileList)
    If Len(FileNames) Then
        uDropEffect = RegisterClipboardFormat(StrPtr("Preferred DropEffect"))
        hGblEffect = GlobalAlloc(GMEM_ZEROINIT Or GMEM_MOVEABLE Or GMEM_DDESHARE, Len(uDropEffect))
        mPtr = GlobalLock(hGblEffect)
        i = IIf(CopyMode, DROPEFFECT_COPY, DROPEFFECT_MOVE)
        CopyMemory ByVal mPtr, i, Len(i)
        GlobalUnlock hGblEffect
        uDropFilesLen = Len(dropFiles)
        With dropFiles
            .pFiles = uDropFilesLen
            .fWide = CLng(True)
        End With
        uGblLen = uDropFilesLen + Len(FileNames) * 2 + 8
        hGblFiles = GlobalAlloc(GMEM_ZEROINIT Or GMEM_MOVEABLE Or GMEM_DDESHARE, uGblLen)
        mPtr = GlobalLock(hGblFiles)
        CopyMemory ByVal mPtr, dropFiles, uDropFilesLen
        CopyMemory ByVal (mPtr + uDropFilesLen), ByVal StrPtr(FileNames), LenB(FileNames)
        GlobalUnlock hGblFiles
        SetClipboardData CF_HDROP, hGblFiles
    End If
    CloseClipboard
End If
End Function

Private Function GetFileListString(FileList As Variant) As String
    Dim i As Byte, j As Byte, k As Byte
    
    On Error GoTo GetFileListStringLOOP
    Select Case VarType(FileList)
        Case vbString
            GetFileListString = Trim$(FileList)
        Case &H2008
            j = LBound(FileList)
            k = UBound(FileList)
            For i = j To k
            FileList(i) = Trim$(FileList(i))
            If Len(FileList(i)) Then GetFileListString = GetFileListString & FileList(i) & vbNullChar
            Next i
            If Len(GetFileListString) Then GetFileListString = Left$(GetFileListString, Len(GetFileListString) - 1)
    End Select
GetFileListStringLOOP:
End Function
