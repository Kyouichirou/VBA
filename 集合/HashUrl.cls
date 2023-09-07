VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HashUrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const CRC32x As String = "https://en.wikipedia.org/wiki/Cyclic_redundancy_check"
Private Const MD5 As String = "https://docs.microsoft.com/en-us/dotnet/api/system.security.cryptography.md5?view=netframework-4.8"
Private Const SHA1 As String = "https://docs.microsoft.com/en-us/dotnet/api/system.security.cryptography.sha1?view=netframework-4.8"
Private Const SHA256 As String = "https://docs.microsoft.com/en-us/dotnet/api/system.security.cryptography.sha256?view=netframework-4.8"
Private Const SHA384 As String = "https://docs.microsoft.com/en-us/dotnet/api/system.security.cryptography.sha384?view=netframework-4.8"
Private Const SHA512 As String = "https://docs.microsoft.com/en-us/dotnet/api/system.security.cryptography.sha512?view=netframework-4.8"

Friend Property Get CRC32URL() As String
    CRC32URL = CRC32x
End Property

Friend Property Get MD5URL() As String
    MD5URL = MD5
End Property

Friend Property Get SHA1URL() As String
   SHA1URL = SHA1
End Property

Friend Property Get SHA256URL() As String
    SHA256URL = SHA256
End Property

Friend Property Get SHA384URL() As String
    SHA384URL = SHA384
End Property

Friend Property Get SHA512URL() As String
    SHA512URL = SHA512
End Property
