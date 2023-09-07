VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cHMAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const mMD5 As String = "System.Security.Cryptography.HMACMD5"
Private Const mSHA1 As String = "System.Security.Cryptography.HMACSHA1"
Private Const mSHA256 As String = "System.Security.Cryptography.HMACSHA256"
Private Const mSHA384 As String = "System.Security.Cryptography.HMACSHA384"
Private Const mSHA512 As String = "System.Security.Cryptography.HMACSHA512"
Private Const sUTF8 As String = "System.Text.UTF8Encoding"
Private Const sXML2 As String = "MSXML2.DOMDocument"
Private objUTF8 As Object
Private objHash As Object
Private objXML As Object

Friend Function HMACMD5(ByVal strText As String, ByVal Password As String, Optional ByVal IsHex As Boolean = False) As String
    HMACMD5 = GetHMACString(strText, Password, mMD5, IsHex, rType)
End Function

Friend Function HMACSH1(ByVal strText As String, ByVal Password As String, Optional ByVal IsHex As Boolean = False, Optional ByVal rType As Boolean) As String
    HMACSHA1 = GetHMACString(strText, Password, mSHA1, IsHex, rType)
End Function

Friend Function HMACSHA256(ByVal strText As String, ByVal Password As String, Optional ByVal IsHex As Boolean = False, Optional ByVal rType As Boolean) As String
    HMACSHA256 = GetHMACString(strText, Password, mSHA256, IsHex, rType)
End Function

Friend Function HMACSHA384(ByVal strText As String, ByVal Password As String, Optional ByVal IsHex As Boolean = False, Optional ByVal rType As Boolean) As String
    HMACSHA384 = GetHMACString(strText, Password, mSHA384, IsHex, rType)
End Function

Friend Function HMACSHA512(ByVal strText As String, ByVal Password As String, Optional ByVal IsHex As Boolean = False, Optional ByVal rType As Boolean) As String
    HMACSHA512 = GetHMACString(strText, Password, mSHA512, IsHex, rType)
End Function

Private Function GetHMACString(ByVal strText As String, ByVal Password As String, ByVal Algorithmx As String, Optional ByVal IsHex As Boolean, Optional ByVal rType As Boolean) As String
    Dim SecretKey() As Byte
    Dim Bytes() As Byte
    
    If LenB(strText) = 0 Then GetHMACString = "": Exit Function
    If LenB(Password) = 0 Then Password = "Password"
    If objHash Is Nothing Then Set objHash = CreateObject(Algorithmx)
    Bytes = objUTF8.GetBytes_4(strText)
    SecretKey = objUTF8.GetBytes_4(Password)
    objHash.key = SecretKey
    Bytes = objHash.ComputeHash_2((Bytes))
    If IsHex = False Then
       GetHMACString = ConvToBase64String(Bytes)
    Else
       GetHMACString = ConvToHexString(Bytes)
    End If
    If rType = True Then GetHMACString = UCase$(GetHMACString)
End Function

Private Function ConvToBase64String(ByRef Arrayx() As Byte) As String '×ª»»ÎªBase64×Ö·û´®
    With objXML
      .LoadXML "<root />"
      .DocumentElement.DataType = "bin.base64"
      .DocumentElement.nodeTypedValue = Arrayx
      ConvToBase64String = Replace$(.DocumentElement.Text, vbLf, "", 1, -1, vbBinaryCompare)
    End With
End Function

Private Function ConvToHexString(ByRef Arrayx() As Byte) As String '×ª»»Îªhex×Ö·û´®
    With objXML
      .LoadXML "<root />"
      .DocumentElement.DataType = "bin.Hex"
      .DocumentElement.nodeTypedValue = Arrayx
      ConvToHexString = Replace$(.DocumentElement.Text, vbLf, "", 1, -1, vbBinaryCompare)
    End With
End Function

Private Sub Class_Initialize()
    Set objUTF8 = CreateObject(sUTF8)
    Set objXML = CreateObject(sXML2)
End Sub

Private Sub Class_Terminate()
    objHash.Clear
    Set objUTF8 = Nothing
    Set objHash = Nothing
    Set objXML = Nothing
End Sub
