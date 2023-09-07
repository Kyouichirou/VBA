Attribute VB_Name = "RSA����"
Option Explicit
'����ϸ�ڲ�������Ҫ��һ��ϸ��
'---------------https://docs.microsoft.com/zh-cn/dotnet/api/system.security.cryptography.rsaencryptionpadding?view=netframework-4.8
Private Const RSA As String = "System.Security.Cryptography.RSACryptoServiceProvider"
Dim objRSA As Object

Function AlgorithmRSA(ByVal strText As String, ByVal IsEncrypt As Boolean, Optional ByVal privateKey As String) As String() 'ѡ����ܻ��߽���
    Dim arr() As String
    Dim publicKey As String
    
    If Len(strText) = 0 Then Exit Function
    Set objRSA = CreateObject(RSA)
    If IsEncrypt = True Then
        ReDim arr(2)
        ReDim AlgorithmRSA(2)
        '--------------https://docs.microsoft.com/zh-cn/dotnet/api/system.security.cryptography.rsa.fromxmlstring?redirectedfrom=MSDN&view=netframework-4.8
        With objRSA
            publicKey = .ToXmlString(False) '������Կ
            privateKey = .ToXmlString(True) '������Կ��˽Կ
        End With
        arr(1) = publicKey
        arr(2) = privateKey
        arr(0) = Encrypt(strText, publicKey) '��Կ����
    Else
        If Len(privateKey) = 0 Then Exit Function
        ReDim arr(0)
        ReDim AlgorithmRSA(0)
        arr(0) = Decrypt(strText, privateKey) '˽Կ����
    End If
    AlgorithmRSA = arr
    Set objRSA = Nothing
End Function

Private Function Encrypt(ByVal strText As String, ByVal publicKey As String) As String
    Dim Bytes() As Byte
    Dim Hash() As Byte
    Dim Result As String
    Dim i As Integer, k As Integer
    '--------------------https://docs.microsoft.com/zh-cn/dotnet/api/system.security.cryptography.rsa.encrypt?view=netframework-4.8
    Bytes = strText
    objRSA.FromXmlString publicKey
    Hash = objRSA.Encrypt(Bytes, False)
    i = UBound(Hash)
    For k = 0 To i
        Result = Result & Right("00" & Hex(Hash(k)), 2)
    Next
    Encrypt = Result
End Function

Private Function Decrypt(ByVal strText As String, ByVal privateKey As String) As String
    Dim bLen As Integer
    Dim Hash() As Byte
    Dim i As Integer
    '-----------------https://docs.microsoft.com/zh-cn/dotnet/api/system.security.cryptography.rsa.decrypt?view=netframework-4.8
    bLen = Len(strText) \ 2
    bLen = bLen - 1
    ReDim Hash(bLen)
    For i = 0 To bLen
        Hash(i) = CByte("&H" & Mid$(strText, i * 2 + 1, 2)) 'CByte������תΪbyte����
    Next
    objRSA.FromXmlString privateKey
    Decrypt = objRSA.Decrypt(Hash, False)
End Function
