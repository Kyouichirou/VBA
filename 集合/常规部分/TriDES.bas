Attribute VB_Name = "TriDES"
Option Explicit
'------------------https://docs.microsoft.com/zh-cn/dotnet/api/system.security.cryptography.tripledes?view=netframework-4.8
Private Const iniVector As String * 8 = "12345678" '指定偏移向量
Private Const pubKEY As String * 16 = "testpassword0000"
Private Const dUTF8 As String = "System.Text.UTF8Encoding"
Private Const tDES As String = "System.Security.Cryptography.TripleDESCryptoServiceProvider"

Function EncryptStringTripleDES(ByVal strText As String) As String '加密
    Dim objDES As Object
    Dim objUTF8 As Object
    Dim Bytes() As Byte, kBytes() As Byte, iBytes() As Byte
    Dim Hash() As Byte
    Dim Result As String
    Dim i As Integer
    
    Set objUTF8 = CreateObject(dUTF8)
    Set objDES = CreateObject(tDES)
    With objUTF8
        Bytes = .GetBytes_4(strText)
        kBytes = .GetBytes_4(pubKEY)
        iBytes = .GetBytes_4(iniVector)
    End With
    i = UBound(Bytes) + 1
    With objDES
        '.blockSize = 64
        '.keysize = 192
        .key = kBytes
        .Iv = iBytes
        Hash = .CreateEncryptor().TransformFinalBlock(Bytes, 0, i)
        .Clear
    End With
    Result = BytesToBase64(Hash)
    EncryptStringTripleDES = Result
    Set objUTF8 = Nothing
    Set objDES = Nothing
End Function

Function DecryptStringTripleDES(ByVal strText As String) As String '解密
    Dim objDES As Object
    Dim objUTF8 As Object
    Dim Bytes() As Byte, kBytes() As Byte, iBytes() As Byte
    Dim Hash() As Byte
    Dim Result As String
    Dim i As Integer
    
    Bytes = Base64toBytes(strText)
    Set objUTF8 = CreateObject(dUTF8)
    Set objDES = CreateObject(tDES)
    With objUTF8
        kBytes = .GetBytes_4(pubKEY)
        iBytes = .GetBytes_4(iniVector)
    End With
    i = UBound(Bytes) + 1
    With objDES
        .key = kBytes
        .Iv = iBytes
        Hash = .CreateDecryptor().TransformFinalBlock(Bytes, 0, i)
        .Clear
    End With
    Result = objUTF8.GetString(Hash)
    DecryptStringTripleDES = Result
    Set objUTF8 = Nothing
    Set objDES = Nothing
End Function

Private Function BytesToBase64(ByRef varBytes() As Byte) As String
    With CreateObject("MSXML2.DomDocument").createElement("b64")
        .DataType = "bin.base64"
        .nodeTypedValue = varBytes
        BytesToBase64 = Replace(.Text, vbLf, "")
    End With
End Function

Private Function Base64toBytes(ByVal varStr As String) As Byte()
    With CreateObject("MSXML2.DOMDocument").createElement("b64")
         .DataType = "bin.base64"
         .Text = varStr
         Base64toBytes = .nodeTypedValue
    End With
End Function


