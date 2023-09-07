Attribute VB_Name = "MD5S"
Option Explicit
Private Const HashTypeMD5 As String = "MD5" ' https://msdn.microsoft.com/en?us/library/system.security.cryptography.md5cryptoserviceprovider(v=vs.110).aspx
Private Const HashTypeSHA1 As String = "SHA1" ' https://msdn.microsoft.com/en?us/library/system.security.cryptography.sha1cryptoserviceprovider(v=vs.110).aspx
Private Const HashTypeSHA256 As String = "SHA256" ' https://msdn.microsoft.com/en?us/library/system.security.cryptography.sha256cryptoserviceprovider(v=vs.110).aspx
Private Const HashTypeSHA384 As String = "SHA384" ' https://msdn.microsoft.com/en?us/library/system.security.cryptography.sha384cryptoserviceprovider(v=vs.110).aspx
Private Const HashTypeSHA512 As String = "SHA512" ' https://msdn.microsoft.com/en?us/library/system.security.cryptography.sha512cryptoserviceprovider(v=vs.110).aspx

Private uFileSize As Double ' Comment out if not testing performance by FileHashes()
Public Tracenumx As Byte '标记追踪错误返回

Function FileHashes(oTestFile As String) As String
    Dim tStart As Date, tFinish As Date, oBlockSize As Variant
    Dim blockSize As Double
    
    oBlockSize = "2^17-1"             '块大小,可以调整,计算的速度会发生变化
    blockSize = Evaluate(oBlockSize)
    '----------------------------------https://www.cnblogs.com/gongyanxu/p/8637965.html
    '------------------------------------------https://docs.microsoft.com/zh-cn/office/vba/api/Excel.Application.Evaluate
    FileHashes = GetFileHash(oTestFile, blockSize, HashTypeMD5)
End Function

Private Function GetFileHash(ByVal sFile As String, ByVal uBlockSize As Double, ByVal sHashType As String) As String
    Dim oCSP As Object ' One of the "CryptoServiceProvider"
    Dim oRnd As MdHash ' "MdHash" Class by Microsoft, must be in the same file
    Dim uBytesRead As Double, uBytesToRead As Double, bDone As Boolean
    Dim aBlock() As Byte, aBytes As Variant ' Arrays to store bytes
    Dim AHash() As Byte, sHash As String, i As Long
    
    Set oRnd = New MdHash '调用类模块
    Set oCSP = CreateObject("System.Security.Cryptography." & sHashType & "CryptoServiceProvider")
    uFileSize = fso.GetFile(sFile).Size  '注意这里不要使用filelen, 多种约束条件
    Tracenumx = 0 '工程级变量,使用之前进行初始化
    On Error GoTo Cleanup
    If oRnd Is Nothing Or oCSP Is Nothing Then GetFileHash = "UC": GoTo Cleanup
    uBytesRead = 0
    bDone = False
    sHash = String(oCSP.HashSize / 4, "0") ' Each hexadecimal has 4 bits
    ' Process the file in chunks of uBlockSize or less
    If uFileSize = 0 Then
       ReDim aBlock(0)
       oCSP.TransformFinalBlock aBlock, 0, 0
       bDone = True
    Else
       With oRnd
           .OpenFile sFile
           If Tracenumx = 1 Then GetFileHash = "UC": GoTo Cleanup '无法有效打开文件
           Do
              If uBytesRead + uBlockSize < uFileSize Then
                  uBytesToRead = uBlockSize
              Else
                  uBytesToRead = uFileSize - uBytesRead
                  bDone = True
              End If
              ' Read in some bytes
              aBytes = .ReadBytes(uBytesToRead)
              aBlock = aBytes
              If bDone Then
                  oCSP.TransformFinalBlock aBlock, 0, uBytesToRead
                  uBytesRead = uBytesRead + uBytesToRead
              Else
                  uBytesRead = uBytesRead + oCSP.TransformBlock(aBlock, 0, uBytesToRead, aBlock, 0)
              End If
              DoEvents
           Loop Until bDone
           .CloseFile
       End With
    End If
    If bDone Then
        ' convert Hash byte array to an hexadecimal string
        AHash = oCSP.Hash
        For i = 0 To UBound(AHash)
            Mid$(sHash, i * 2 + (AHash(i) > 15) + 2) = Hex(AHash(i))
        Next
    End If
    GetFileHash = sHash
Cleanup:
    Set oRnd = Nothing
    Set oCSP = Nothing
End Function
