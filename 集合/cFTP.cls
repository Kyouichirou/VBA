VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const BufferSize = 255

Private hOpen As Long
Private hConnection As Long
Private hFile As Long
Private dwType As Long '数据传输的类型
Private dwSeman As Long

Private szErrorMessage As String

Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As Currency
        ftLastAccessTime As Currency
        ftLastWriteTime As Currency
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * 260
        cAlternate As String * 14
End Type

Private Const MAX_PATH = 260
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const NO_ERROR = 0
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_OFFLINE = &H1000
Private Const INTERNET_FLAG_PASSIVE = &H8000000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000

Private Const ERROR_NO_MORE_FILES = 18
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_INVALID_PORT_NUMBER = 0
Private Const INTERNET_SERVICE_FTP = 1
Private Const FTP_TRANSFER_TYPE_BINARY = &H2
Private Const FTP_TRANSFER_TYPE_ASCII = &H1

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, ByVal lpSource As Long, _
ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Private Declare Function FTPGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, _
ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, _
ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, _
ByVal lpszOldName As String, ByVal lpszNewName As String) As Boolean
'------------------------------------------- ------------------ https://docs.microsoft.com/en-us/windows/win32/wininet/portal?redirectedfrom=MSDN

Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszName As String) As Boolean
Private Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal lpszName As String) As Boolean

Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean
Private Declare Function FtpOpenFile Lib "wininet.dll" Alias "FtpOpenFileA" (ByVal hFtpSession As Long, ByVal sBuff As String, _
ByVal Access As Long, ByVal flags As Long, ByVal Context As Long) As Long

Private Declare Function FTPPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hFtpSession As Long, ByVal lpszLocalFile As String, _
ByVal lpszRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, _
ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, _
lpdwCurrentDirectory As Long) As Boolean

Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, _
ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpLibFileName As String) As Long

Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
Private Declare Function InternetWriteFile Lib "wininet.dll" (ByVal hFile As Long, ByRef sBuffer As Byte, ByVal lNumBytesToWrite As Long, _
dwNumberOfBytesWritten As Long) As Integer
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByRef sBuffer As Byte, _
ByVal lNumBytesToRead As Long, dwNumberOfBytesRead As Long) As Integer

Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Long
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, _
ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long

Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (ByRef lpdwError As Long, _
ByVal lpszErrorBuffer As String, ByRef lpdwErrorBufferLength As Long) As Boolean

Public Event FileTransferProgress(lCurrentBytes As Long, lTotalBytes As Long)

Private Sub ErrorOut(ByVal dwError As Long, ByRef szFunc As String)
    Dim dwRet As Long
    Dim dwTemp As Long
    Dim szString As String * 2048
    dwRet = FormatMessage(FORMAT_MESSAGE_FROM_HMODULE, _
                      GetModuleHandle("wininet.dll"), dwError, 0, _
                      szString, 256, 0)
    szErrorMessage = szFunc & " error code: " & dwError & " Message: " & szString
    If (dwError = 12003) Then
        ' Extended error information was returned
        dwRet = InternetGetLastResponseInfo(dwTemp, szString, 2048)
        szErrorMessage = szString
    End If
End Sub

Property Get GetLastErrorMessage() As String
    GetLastErrorMessage = szErrorMessage
End Property

Public Sub CloseConnection()
    If hConnection <> 0 Then
        InternetCloseHandle hConnection
    End If
    hConnection = 0
End Sub

Public Function SimpleFTPPutFile(sLocal As String, sRemote As String) As Boolean
    If (FTPPutFile(hConnection, sLocal, sRemote, dwType, 0) = False) Then
        ErrorOut Err.LastDllError, "SimpleFtpPutFile"
        SimpleFTPPutFile = False
        Exit Function
    Else
        SimpleFTPPutFile = True
    End If
End Function
 
Public Function RenameFTPFile(sExisting As String, sNewName As String) As Boolean
    If (FtpRenameFile(hConnection, sExisting, sNewName) = False) Then
        ErrorOut Err.LastDllError, "RenameFTPFile"
        RenameFTPFile = False
        Exit Function
    Else
        RenameFTPFile = True
    End If
End Function

Public Function CreateFTPDirectory(sDirectory As String) As Boolean
    If (FtpCreateDirectory(hConnection, sDirectory) = False) Then
        ErrorOut Err.LastDllError, "CreateFTPDirectory"
        CreateFTPDirectory = False
        Exit Function
    Else
        CreateFTPDirectory = True
    End If
End Function

Public Function RemoveFTPDirectory(sDirectory As String) As Boolean
    If (FtpRemoveDirectory(hConnection, sDirectory) = False) Then
        ErrorOut Err.LastDllError, "RemoveFTPDirectory"
        RemoveFTPDirectory = False
        Exit Function
    Else
        RemoveFTPDirectory = True
    End If
End Function
 
Public Function DeleteFTPFile(sRemote As String) As Boolean
    If (FtpDeleteFile(hConnection, sRemote) = False) Then
        ErrorOut Err.LastDllError, "DeleteFTPFile"
        DeleteFTPFile = False
        Exit Function
    Else
        DeleteFTPFile = True
    End If
End Function

Public Function OpenConnection(sServer As String, sUser As String, sPassword As String, Optional lPort As Long = 21) As Boolean
    If hConnection <> 0 Then
        InternetCloseHandle hConnection
    End If
    If lPort = 21 Then
        hConnection = InternetConnect(hOpen, sServer, INTERNET_INVALID_PORT_NUMBER, sUser, sPassword, INTERNET_SERVICE_FTP, dwSeman, 0)
    Else
        hConnection = InternetConnect(hOpen, sServer, lPort, sUser, sPassword, INTERNET_SERVICE_FTP, dwSeman, 0)
    End If
    If hConnection = 0 Then
        ErrorOut Err.LastDllError, "InternetConnect"
        OpenConnection = False
        Exit Function
    Else
        OpenConnection = True
    End If
End Function

Public Function FTPUploadFile(sLocal As String, sRemote As String, ByVal FileType As Byte) As Boolean
    Dim Data(BufferSize - 1) As Byte '每次读取数据大大小
    Dim Written As Long
    Dim Size As Long
    Dim Sum As Long
    Dim lBlock As Long
    
    Sum = 0
    lBlock = 0
    sLocal = Trim(sLocal)
    sRemote = Trim(sRemote)
    dwType = FileType
    If sLocal <> "" And sRemote <> "" Then
      hFile = FtpOpenFile(hConnection, sRemote, GENERIC_WRITE, dwType, 0)
      If hFile = 0 Then
          ErrorOut Err.LastDllError, "FtpOpenFile:PutFile"
          FTPUploadFile = False
          Exit Function
      End If
      
      '---------------------------------------这里需要修改为ado模式读取数据,否则路径中带有非ansi编码字符的文件将会无法读取(参照Md5B模块的代码)
      Open sLocal For Binary Access Read As #1
      Size = LOF(1)
      For lBlock = 1 To Size \ BufferSize
          Get #1, , Data
          If (InternetWriteFile(hFile, Data(0), BufferSize, Written) = 0) Then
              FTPUploadFile = False
              ErrorOut Err.LastDllError, "InternetWriteFile"
              Exit Function
          End If
          DoEvents
          Sum = Sum + BufferSize
          RaiseEvent FileTransferProgress(Sum, Size)
      Next lBlock
      
      Get #1, , Data
      If (InternetWriteFile(hFile, Data(0), Size Mod BufferSize, Written) = 0) Then
          FTPUploadFile = False
          ErrorOut Err.LastDllError, "InternetWriteFile2"
          Exit Function
      End If
      
      Sum = Sum + (Size Mod BufferSize)
      Size = Sum
      RaiseEvent FileTransferProgress(Sum, Size)

      Close #1
      InternetCloseHandle (hFile)
      FTPUploadFile = True
   End If
End Function

Public Function FTPDownloadFile(sLocal As String, sRemote As String, ByVal FileType As Byte) As Boolean
    Dim Data(BufferSize - 1) As Byte ' array of 100 elements 0 to 99
    Dim Written As Long
    Dim Size As Long
    Dim Sum As Long
    Dim lBlock As Long
    FTPDownloadFile = False
            
    Sum = 0
    lBlock = 0
    
    sLocal = Trim(sLocal)
    sRemote = Trim(sRemote)
    
    If sLocal <> "" And sRemote <> "" Then
      Size = GetFTPFileSize(sRemote)
      DoEvents
      dwType = FileType
      If Size > 0 Then
          hFile = FtpOpenFile(hConnection, sRemote, GENERIC_READ, dwType, 0)
          DoEvents
          If hFile = 0 Then
              ErrorOut Err.LastDllError, "FtpOpenFile:GetFile"
              Exit Function
          End If
          '---------------------------------------这里需要修改为ado模式读取数据,否则路径中带有非ansi编码字符的文件将会无法读取(参照Md5B模块的代码)
          Open sLocal For Binary Access Write As #1
          DoEvents
          Seek #1, 1
          Sum = 1
          For lBlock = 1 To Size \ BufferSize
              If (InternetReadFile(hFile, Data(0), BufferSize, Written) = 0) Then
                  ErrorOut Err.LastDllError, "InternetReadFile"
                  Close #1
                  Exit Function
              End If
              Put #1, , Data
              DoEvents
              Sum = Sum + BufferSize
              RaiseEvent FileTransferProgress(Sum, Size)
              DoEvents
          Next lBlock
          
          DoEvents
          ReDim Data2((Size Mod BufferSize) - 1) As Byte
          If (InternetReadFile(hFile, Data2(0), Size Mod BufferSize, Written) = 0) Then
              ErrorOut Err.LastDllError, "InternetReadFile2"
              Close #1
              Exit Function
          End If
          Put #1, , Data2
          Sum = Sum + (Size Mod BufferSize)
          Size = Sum
          RaiseEvent FileTransferProgress(Sum, Size)
          Close #1
          InternetCloseHandle (hFile)
          DoEvents
          FTPDownloadFile = True
          DoEvents
      End If
   End If
End Function

Public Function SimpleFTPGetFile(sLocal As String, sRemote As String) As Boolean
   ' add INTERNET_FLAG_NO_CACHE_WRITE to avoid local caching 0x04000000 (hex)
    If (FTPGetFile(hConnection, sRemote, sLocal, False, FILE_ATTRIBUTE_NORMAL, dwType Or INTERNET_FLAG_RELOAD, 0) = False) Then
        ErrorOut Err.LastDllError, "SimpleFtpGetFile"
        SimpleFTPGetFile = False
        Exit Function
    Else
        SimpleFTPGetFile = True
    End If
End Function

Public Function GetFTPDirectory() As String
    Dim szDir As String
    szDir = String(1024, Chr$(0))
    If (FtpGetCurrentDirectory(hConnection, szDir, 1024) = False) Then
        ErrorOut Err.LastDllError, "FtpGetCurrentDirectory"
        Exit Function
    Else
        GetFTPDirectory = Left(szDir, InStr(1, szDir, String(1, 0), vbBinaryCompare) - 1)
    End If
End Function

Public Function SetFTPDirectory(sDir As String)
    If (FtpSetCurrentDirectory(hConnection, sDir) = False) Then
        ErrorOut Err.LastDllError, "FtpSetCurrentDirectory"
        SetFTPDirectory = False
        Exit Function
    Else
        SetFTPDirectory = True
    End If
End Function

Public Function GetFTPFileSize(sFile As String) As Long
    Dim szDir As String
    Dim hFind As Long
    Dim nLastError As Long
    Dim pData As WIN32_FIND_DATA
    
    hFind = FtpFindFirstFile(hConnection, sFile, pData, 0, 0)
    nLastError = Err.LastDllError
    If hFind = 0 Then
        If (nLastError = ERROR_NO_MORE_FILES) Then
            GetFTPFileSize = -1  ' File not found
        Else
            GetFTPFileSize = -2  ' Other error
            ErrorOut Err.LastDllError, "FtpFindFirstFile"
        End If
        Exit Function
    End If
           
    GetFTPFileSize = pData.nFileSizeLow
    InternetCloseHandle (hFind)
End Function

Public Function GetFTPDirectoryContent() As String
    Dim hFind As Long
    Dim nLastError As Long
    Dim pData As WIN32_FIND_DATA
    Dim tmp As String
    GetFTPDirectoryContent = ""
    hFind = FtpFindFirstFile(hConnection, "*.*", pData, 0, 0)
    nLastError = Err.LastDllError
    If hFind = 0 Then
        If (nLastError = ERROR_NO_MORE_FILES) Then
            GetFTPDirectoryContent = "False"  ' File not found
        Else
            GetFTPDirectoryContent = "False"  ' Other error
            ErrorOut Err.LastDllError, "FtpFindFirstFile"
        End If
        Exit Function
    Else
        Do
            tmp = pData.cFileName
            If InStr(tmp, Chr$(0)) > 0 Then
                tmp = Left(tmp, InStr(tmp, Chr$(0)) - 1)
            End If
            If Len(tmp) Then
                If pData.dwFileAttributes And vbDirectory Then
                    If GetFTPDirectoryContent = "" Then
                        GetFTPDirectoryContent = tmp & "/"
                    Else
                        GetFTPDirectoryContent = GetFTPDirectoryContent & vbCrLf & tmp & "/"
                    End If
                Else
                    If GetFTPDirectoryContent = "" Then
                        GetFTPDirectoryContent = tmp & "(" & Format(pData.nFileSizeLow / 1024, "0.00") & "kB)"
                    Else
                        GetFTPDirectoryContent = GetFTPDirectoryContent & vbCrLf & tmp & "(" & Format(pData.nFileSizeLow / 1024, "0.00") & "kB)"
                    End If
                End If
            End If
        Loop While InternetFindNextFile(hFind, pData)
    End If
    InternetCloseHandle (hFind)
End Function

Public Sub SetTransferASCII()
    dwType = FTP_TRANSFER_TYPE_ASCII
End Sub

Public Sub SetTransferBinary()
    dwType = FTP_TRANSFER_TYPE_BINARY
End Sub

Public Sub SetModeActive()
    dwSeman = 0
End Sub

Public Sub SetModePassive()
    dwSeman = INTERNET_FLAG_PASSIVE
End Sub

Private Sub Class_Initialize()
    hOpen = InternetOpen("eDIY FTP Client", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    If hOpen = 0 Then
        ErrorOut Err.LastDllError, "InternetOpen"
    End If
    dwType = FTP_TRANSFER_TYPE_ASCII '默认使用asc, 在执行过程中根据实际的文件类型进行调整
'    dwType = FTP_TRANSFER_TYPE_BINARY 'FTP在传送文件时分为ASC和Bin两种格式，只有文字文件(例如 html 文件)使用ASC(涉及到系统差异Unix/windows)，其他的通通使用Bin格式(例如图像文件、压缩文件、可执行文件等等)。
    dwSeman = 0
    hConnection = 0
End Sub

Private Sub Class_Terminate()
    InternetCloseHandle hOpen
End Sub

