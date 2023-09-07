Attribute VB_Name = "ÎÄ¼þÏÂÔØ"
Option Explicit
Option Compare Text

Private Enum DownloadFileDisposition
    OverwriteKill = 0
    OverwriteRecycle = 1
    DoNotOverwrite = 2
    PromptUser = 3
End Enum

Private Declare Function SHFileOperation Lib "shell32.dll" Alias _
    "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long   'https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/bb776426(v=vs.85)
                                                              'https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/bb776779(v=vs.85)
Private Declare Function PathIsNetworkPath Lib "shlwapi.dll" _
    Alias "PathIsNetworkPathA" ( _
    ByVal pszPath As String) As Long                        'https://docs.microsoft.com/en-us/windows/win32/api/shlwapi/nf-shlwapi-pathfindfilenamea

Private Declare Function GetSystemDirectory Lib "kernel32" _
    Alias "GetSystemDirectoryA" ( _
    ByVal lpBuffer As String, _
    ByVal nSize As Long) As Long

Private Declare Function SHEmptyRecycleBin _
    Lib "shell32" Alias "SHEmptyRecycleBinA" _
    (ByVal hwnd As Long, _
     ByVal pszRootPath As String, _
     ByVal dwFlags As Long) As Long

Private Const FO_Delete = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_NOCONFIRMATION = &H10
Private Const MAX_PATH As Long = 260

Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
  "URLDownloadToFileA" ( _
  ByVal pCaller As Long, _
  ByVal szURL As String, _
  ByVal szFileName As String, _
  ByVal dwReserved As Long, _
  ByVal lpfnCB As Long) As Long

Private Function DownloadFile(UrlFileName As String, DestinationFileName As String, Overwrite As DownloadFileDisposition, ErrorText As String) As Boolean
    Dim Disp As DownloadFileDisposition
    Dim Res As VbMsgBoxResult
    Dim b As Boolean
    Dim s As String
    Dim l As Long
    
    ErrorText = vbNullString
    
    If Dir(DestinationFileName, vbNormal) <> vbNullString Then
        Select Case Overwrite
            Case OverwriteKill
                On Error Resume Next
                Err.Clear
                Kill DestinationFileName
                If Err.Number <> 0 Then
                    ErrorText = "Error Kill'ing file '" & DestinationFileName & "'." & vbCrLf & Err.Description
                    DownloadFile = False
                    Exit Function
                End If
        
            Case OverwriteRecycle
                On Error Resume Next
                Err.Clear
                b = RecycleFileOrFolder(DestinationFileName)
                If b = False Then
                    ErrorText = "Error Recycle'ing file '" & DestinationFileName & "." & vbCrLf & Err.Description
                    DownloadFile = False
                    Exit Function
                End If
            
            Case DoNotOverwrite
                DownloadFile = False
                ErrorText = "File '" & DestinationFileName & "' exists and disposition is set to DoNotOverwrite."
                Exit Function
                
            'Case PromptUser
            Case Else
                s = "The destination file '" & DestinationFileName & "' already exists." & vbCrLf & _
                    "Do you want to overwrite the existing file?"
                Res = MsgBox(s, vbYesNo, "Download File")
                If Res = vbNo Then
                    ErrorText = "User selected not to overwrite existing file."
                    DownloadFile = False
                    Exit Function
                End If
                b = RecycleFileOrFolder(DestinationFileName)
                If b = False Then
                    ErrorText = "Error Recycle'ing file '" & DestinationFileName & "." & vbCrLf & Err.Description
                    DownloadFile = False
                    Exit Function
                End If
        End Select
    End If
    
    l = URLDownloadToFile(0&, UrlFileName, DestinationFileName, 0&, 0&)
    If l = 0 Then
        DownloadFile = True
    Else
        ErrorText = "Buffer length invalid or not enough memory."
        DownloadFile = False
    End If
End Function
                            
Private Function RecycleFileOrFolder(FileSpec As String) As Boolean
    Dim FileOperation As SHFILEOPSTRUCT
    Dim lReturn As Long

    If (Dir(FileSpec, vbNormal) = vbNullString) And _
        (Dir(FileSpec, vbDirectory) = vbNullString) Then
        RecycleFileOrFolder = True
        Exit Function
    End If

    With FileOperation
        .wFunc = FO_Delete
        .pFrom = FileSpec
        .fFlags = FOF_ALLOWUNDO
        ' Or
        .fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION
    End With

    lReturn = SHFileOperation(FileOperation)
    If lReturn = 0 Then
        RecycleFileOrFolder = True
    Else
        RecycleFileOrFolder = False
    End If
End Function

Function DownloadFilex(ByVal url As String, ByVal FileName As String) As Boolean
    Dim LocalFileName As String
    Dim ErrorText As String
    
    DownloadFilex = DownloadFile(UrlFileName:=url, _
                     DestinationFileName:=FileName, _
                     Overwrite:=OverwriteRecycle, _
                     ErrorText:=ErrorText)
End Function
