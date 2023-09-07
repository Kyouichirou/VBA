Attribute VB_Name = "rAES����"
Option Explicit
'https://docs.microsoft.com/zh-cn/dotnet/api/system.security.cryptography.symmetricalgorithm.createencryptor?view=netframework-4.8#System_Security_Cryptography_SymmetricAlgorithm_CreateEncryptor_System_Byte___System_Byte___
'https://docs.microsoft.com/zh-cn/dotnet/api/system.security.cryptography.icryptotransform?view=netframework-4.8
'����AES�㷨��һ��(Rijndael was the winner of the NIST competition to select the algorithm that would become AES.However,
'------------------there are some differences between Rijndael and the official FIPS-197 specification for AES.)
'---------'https://docs.microsoft.com/en-us/archive/blogs/shawnfa/the-differences-between-rijndael-and-aes
'CipherMode
Private Const CipherMode_CBC As Byte = 1
Private Const CipherMode_ECB As Byte = 2
Private Const CipherMode_OFB As Byte = 3
Private Const CipherMode_CFB As Byte = 4
Private Const CipherMode_CTS As Byte = 5
'------------------------------------------//
Private Const rAES As String = "System.Security.Cryptography.RijndaelManaged"
Private Const rUTF8 As String = "System.Text.UTF8Encoding"
Private Const rSHA256 As String = "System.Security.Cryptography.SHA256Managed"
Private Const eBlockSize As Long = 130172 '���ܶ�ȡ���ݿ�Ĵ�С
Private Const dBlockSize As Long = 130176 '���ܺ�д��Ŀ��С/���ܶ�ȡ���С '����޸Ĵ˲���,��Ҫ����ͬ���޸�
'���� --------final block
'inputBuffer--------------ע�������final block �� fromblock��hash�����ϸ΢����
'Byte[]
'ҪΪ�����ת��������
'inputOffset
'Int32
'�ֽ������е�ƫ�������Ӹ�λ�ÿ�ʼʹ�����ݡ�
'inputCount
'Int32
'�ֽ��������������ݵ��ֽ���
'//------------------------------------------------------------//

'-------------------��Ҫע����Ǿ��ܶ��ļ������˼���,�����ļ�����opentextasstream��Ȼ����д���ƻ����ļ�������
Function AESEncrypt(ByVal strText As String, ByVal key As String, ByVal IsFile As Boolean, Optional ByVal IsRemove As Boolean = False, Optional ByVal filez As Long, _
Optional ByVal OutPut As String) As String() '����
    Dim Hash() As Byte
    Dim Bytes() As Byte
    Dim FsoA As Object
    Dim ado As Object, Adob As Object
    Dim TempFile As String, i As Long, k As Long
    Dim objRijndael As Object
    Dim objUTF8 As Object
    Dim arrTemp() As String
    
    If Len(strText) = 0 Then Exit Function
    If IsFile = True Then 'ѡ���ļ�/�ַ���
        Set FsoA = CreateObject("Scripting.FileSystemObject")
        If FsoA.fileexists(strText) = False Then Exit Function
        If filez = 0 Then
            filez = FsoA.GetFile(strText).Size
            If filez = 0 Then Exit Function
        End If
        key = SHA256(key)
        Set objRijndael = CreateObject(rAES)
        Set ado = CreateObject("adodb.stream")
        With objRijndael
            .Mode = CipherMode_CBC
            .blockSize = 128
            .keysize = 256
            .Iv   'GenerateIV        'ƫ��ֵ
            'In general, there is no reason to use this method, because CreateEncryptor() or CreateEncryptor(null, null)
            'automatically generates both an initialization vector and a key. However, you may want to use the GenerateIV method to reuse a symmetric algorithm instance
            'with a different initialization vector.
            '���㴴��һ�� SymmetricAlgorithm �����ʵ�����ֶ����� GenerateIV ����ʱ��IV ���Ի��Զ�����Ϊ�µ����ֵ�� IV ���ԵĴ�С������� BlockSize ���Գ���8��
            '������ SymmetricAlgorithm �����ʹ�ó�Ϊ���ܿ����ӣ�CBC��������ģʽ������Ҫһ����Կ��һ����ʼ��������������ִ�м���ת����
            '��Ҫ��ʹ�� SymmetricAlgorithm ��֮һ���ܵ����ݽ��н��ܣ����뽫 Key ���Ժ� IV ��������Ϊ���ڼ��ܵ���ֵͬ��
            .key = Hex2Bytes(key)
        End With
        '----------------------------------------ǰ��׼��
        If Len(OutPut) = 0 Then
            OutPut = strText & ".aes" '���ܺ���ļ�·��
        Else
            If FsoA.folderexists(OutPut) = False Then
                OutPut = strText & ".aes" '���ܺ���ļ�·��
            Else
                If Right(OutPut, 1) = "\" Then
                    OutPut = OutPut & Right$(strText, Len(strText) - InStrRev(strText, "\")) & ".aes"
                Else
                    OutPut = OutPut & "\" & Right$(strText, Len(strText) - InStrRev(strText, "\")) & ".aes"
                End If
            End If
        End If
        '---------------ȷ�����ܺ���ļ��Ĵ��λ��
        TempFile = ThisWorkbook.Path & "\temp" & Format(Now, "yyyymmddhhmmss") '��ʱ�ļ�,,��ֹ����temp���ļ���
        FsoA.CreateTextFile(TempFile, True, False).Write (Space(filez)) 'ռλ
        With ado
            .Mode = 3  '����ѡ1,ֻ��ģʽ/,3 ��дģʽ
            .type = 1  'adTypeBinary 'adTypeText=2
            .Open
            .LoadFromFile (strText)
            .Position = 0
            If filez > eBlockSize Then
                k = 0
                Set Adob = CreateObject("adodb.stream")
                With Adob
                    .Mode = 3
                    .type = 1
                    .Open
                    .LoadFromFile (TempFile)
                End With
                Do Until .EOS = True
                    Bytes = .Read(eBlockSize)
                    i = UBound(Bytes) + 1
                    Hash = objRijndael.CreateEncryptor.TransformFinalBlock(Bytes, 0, i)
                    i = UBound(Hash) + 1
                    Adob.Position = k
                    Adob.Write Hash '���д������ݳ����˵�ǰ EOS λ�ã��� Stream �� Size �����ӣ��԰����������ֽڣ����� EOS ���ƶ��� Stream ���µ����һ���ֽڡ�
                    k = k + i
                Loop
                .Close
                Adob.SaveToFile OutPut 'https://docs.microsoft.com/zh-cn/office/client-developer/access/desktop-database-reference/saveoptionsenum
                Adob.Close
                Set Adob = Nothing
            Else
                Bytes = .Read(filez)
                Hash = objRijndael.CreateEncryptor.TransformFinalBlock((Bytes), 0, filez) 'LenB(Bytes)
                .LoadFromFile (TempFile)
                .Write Hash
                .SaveToFile OutPut
                .Close
            End If
        End With
        FsoA.DeleteFile TempFile
        If IsRemove = True Then FsoA.DeleteFile strText 'ɾ��Դ�ļ�
        Set FsoA = Nothing
        Set ado = Nothing
        ReDim AESEncrypt(1)
        ReDim arrTemp(1)
    Else
        key = SHA256(key)
        Set objUTF8 = CreateObject(rUTF8)
        Set objRijndael = CreateObject(rAES)
        With objRijndael
            .Mode = CipherMode_CBC
            .blockSize = 128
            .keysize = 256
            .Iv   'GenerateIV
            .key = Hex2Bytes(key)
        End With
        Bytes = objUTF8.GetBytes_4(strText)
        ReDim AESEncrypt(2)
        ReDim arrTemp(2)
        k = UBound(Bytes) + 1
        Hash = objRijndael.CreateEncryptor.TransformFinalBlock((Bytes), 0, k)
        arrTemp(2) = Bytes2Hex(Hash) '���ܵ�����
        Set objUTF8 = Nothing
    End If
    Hash = objRijndael.Iv
    arrTemp(1) = Bytes2Hex(Hash) 'ƫ��ֵ
    arrTemp(0) = "SHA256:" & key '�µ�key
    objRijndael.Clear
    Set objRijndael = Nothing
    AESEncrypt = arrTemp
    Erase arrTemp
End Function

Function AESDecrypt(ByVal strText As String, ByVal IsFile As Boolean, ByVal Ivx As String, ByVal key As String, Optional ByVal filez As Long, _
Optional ByVal IsRemove As Boolean = False, Optional ByVal OutPut As String) As String '����
    Dim Hash() As Byte
    Dim Bytes() As Byte
    Dim FsoA As Object
    Dim ado As Object, Adob As Object
    Dim TempFile As String, i As Long, k As Long, blockSizex As Long
    Dim objRijndael As Object
    Dim objUTF8 As Object
    Dim strx As String, strx1 As String
    
    If Len(strText) = 0 Then Exit Function
    If IsFile = True Then 'ѡ���ļ�/�ַ���
        Set FsoA = CreateObject("Scripting.FileSystemObject")
        If FsoA.fileexists(strText) = False Then Exit Function
        If filez = 0 Then
            filez = FsoA.GetFile(strText).Size
            If filez = 0 Then Exit Function
        End If
        If LCase$(Right$(strText, 3)) <> "aes" Then Exit Function '�����չ������aes���˳�
        Set objRijndael = CreateObject(rAES)
        Set ado = CreateObject("adodb.stream")
        With objRijndael
            .Mode = CipherMode_CBC
            .blockSize = 128
            .keysize = 256
            .Iv = Hex2Bytes(Ivx)
            .key = Hex2Bytes(key)
        End With
        '----------------------------------------ǰ��׼��
        strx1 = Format(time, "hhmmss")
        strx = Right$(strText, Len(strText) - InStrRev(strText, "\"))
        strx = strx1 & Left$(strx, Len(strx) - 4) '---�µ��ļ���
        If Len(OutPut) = 0 Then
            OutPut = Left$(strText, InStrRev(strText, "\"))
            OutPut = OutPut & strx1 & strx
        Else
            If FsoA.folderexists(OutPut) = False Then
                OutPut = Left$(strText, InStrRev(strText, "\"))
                OutPut = OutPut & strx1 & strx
            Else
                If Right(OutPut, 1) = "\" Then
                    OutPut = OutPut & strx1 & strx
                Else
                    OutPut = OutPut & "\" & strx1 & strx
                End If
            End If
        End If
        '---------------ȷ�����ܺ���ļ��Ĵ��λ��
        'Environ("temp"),ϵͳ�����ļ��е�λ��
        TempFile = ThisWorkbook.Path & "\temp" & Format(Now, "yyyymmddhhmmss") '��ʱ�ļ�,,��ֹ����temp���ļ���
        FsoA.CreateTextFile(TempFile, True, False).Write (Space(filez)) 'ռλ
        With ado
            .Mode = 3  '����ѡ1,ֻ��ģʽ/,3 ��дģʽ
            .type = 1  'adTypeBinary 'adTypeText=2
            .Open
            .LoadFromFile (strText)
            .Position = 0
            If filez > dBlockSize Then
                k = 0
                Set Adob = CreateObject("adodb.stream")
                With Adob
                    .Mode = 3
                    .type = 1
                    .Open
                    .LoadFromFile (TempFile)
                End With
                Do Until .EOS = True
                    Bytes = .Read(dBlockSize)
                    i = UBound(Bytes) + 1
                    Hash = objRijndael.CreateDecryptor.TransformFinalBlock(Bytes, 0, i)
                    '��Ҫע����ܺ�����ݿ�ǰ��д��Ĵ�С�����仯, ֮ǰ���ܺ�����ݴ�130172���130176, �����ܶ�ȡ��ʱ����Ҫ����д��(130176)�Ŀ��С����ȡ
                    '�������ֳ��Ȳ���������
                    i = UBound(Hash) + 1
                    Adob.Position = k
                    Adob.Write Hash '���д������ݳ����˵�ǰ EOS λ�ã��� Stream �� Size �����ӣ��԰����������ֽڣ����� EOS ���ƶ��� Stream ���µ����һ���ֽڡ�
                    k = k + i
                Loop
                .Close
                Adob.SaveToFile OutPut
                Adob.Close
                Set Adob = Nothing
            Else
                Bytes = .Read(filez)
                Hash = objRijndael.CreateDecryptor.TransformFinalBlock((Bytes), 0, filez)
                .LoadFromFile (TempFile)
                .Write Hash
                .SaveToFile OutPut
                .Close
            End If
        End With
        FsoA.DeleteFile TempFile
        If IsRemove = True Then FsoA.DeleteFile strText 'ɾ��Դ�ļ�
        Set FsoA = Nothing
        Set ado = Nothing
    Else
        Set objUTF8 = CreateObject(rUTF8)
        Set objRijndael = CreateObject(rAES)
        Bytes = Hex2Bytes(strText)
        With objRijndael
            .Mode = CipherMode_CBC
            .blockSize = 128
            .keysize = 256
            .Iv = Hex2Bytes(Ivx)
            .key = Hex2Bytes(key)
        End With
        i = UBound(Bytes) + 1
        Hash = objRijndael.CreateDecryptor.TransformFinalBlock((Bytes), 0, i)
        AESDecrypt = objUTF8.GetString(Hash) '���ܵ�����
        Set objUTF8 = Nothing
    End If
    objRijndael.Clear
    Set objRijndael = Nothing
End Function

Private Function Bytes2Hex(ByRef Arrayx() As Byte) As String '������תΪ16�����ַ���
    With CreateObject("Microsoft.XMLDOM").createElement("dummy")
        .DataType = "bin.hex"
        .nodeTypedValue = Arrayx
        Bytes2Hex = .Text
    End With
End Function
'-----------------------------------------------------------https://www.cnblogs.com/hnxxcxg/p/11126688.html
Private Function Hex2Bytes(ByVal strText As String) As Byte() '��16���Ƶ��ַ���תΪ����
    With CreateObject("Microsoft.XMLDOM").createElement("dummy")
        .DataType = "bin.hex"
        .Text = strText
        Hex2Bytes = .nodeTypedValue
    End With
End Function

Private Function SHA256(ByVal strText As String) As String '���ַ���ת��ΪSHA256 Hash
    Dim Bytes() As Byte
    Dim Hash() As Byte
    Dim objSHA256 As Object
    Dim objUTF8 As Object
    
    Set objSHA256 = CreateObject(rSHA256)
    Set objUTF8 = CreateObject(rUTF8)
    Bytes = objUTF8.GetBytes_4(strText)
    Hash = objSHA256.ComputeHash_2((Bytes))
    SHA256 = Bytes2Hex(Hash)
    Set objSHA256 = Nothing
    Set objUTF8 = Nothing
End Function
