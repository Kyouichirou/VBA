Attribute VB_Name = "MD5B"
Option Explicit
Option Base 0
Private Type MD5_CTX
    z(1) As Long
    buf(3) As Long
    inc(63) As Byte
    digest(15) As Byte
End Type
'----------------------------------------------------------------------------------'Any �������з��յ�
'-------------longptr�����Ǿ�����������ͣ�ֻ��һ�ֱ�ǣ���x64office�±�ת��Ϊlonglong�������ͣ� x32����תΪlong�� ����Ϊ�˱���x64��x32֮��ļ���
'---------------'https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/aa366535(v=vs.85)
#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Sub MD5Init Lib "Cryptdll.dll" (ByVal pContex As LongPtr)
    Private Declare PtrSafe Sub MD5Final Lib "Cryptdll.dll" (ByVal pContex As LongPtr)
    Private Declare PtrSafe Sub MD5Update Lib "Cryptdll.dll" (ByVal pContex As Long, ByVal lPtr As Long, ByVal nSize As LongPtr)
#Else
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Private Declare Sub MD5Init Lib "Cryptdll.dll" (ByVal pContex As Long)
    Private Declare Sub MD5Final Lib "Cryptdll.dll" (ByVal pContex As Long)
    Private Declare Sub MD5Update Lib "Cryptdll.dll" (ByVal pContex As Long, ByVal lPtr As Long, ByVal nSize As Long)
#End If
Dim xi() As Byte
Dim Ccx As Long 'ƴ������
 
Private Function ConvBytesToBinaryString(bytesIn() As Byte) As String
    Dim z As Long
    Dim nSize As Long
    Dim strRet As String
    
    nSize = UBound(bytesIn)
    For z = 0 To nSize
         strRet = strRet & Right$("0" & Hex(bytesIn(z)), 2)
    Next
    ConvBytesToBinaryString = strRet
End Function
 
Private Function GetMD5Hash(bytesIn() As Byte) As Byte()
    Dim ctx As MD5_CTX
    Dim nSize As Long
    
    nSize = UBound(bytesIn) + 1
    MD5Init VarPtr(ctx)
    MD5Update ByVal VarPtr(ctx), ByVal VarPtr(bytesIn(0)), nSize
    MD5Final VarPtr(ctx)
    GetMD5Hash = ctx.digest
End Function

Private Function GetMD5Hash_Bytes(bytesIn() As Byte) As String
    GetMD5Hash_Bytes = ConvBytesToBinaryString(GetMD5Hash(bytesIn))
End Function

Function GetMD5Hash_String(ByVal strIn As String, Optional ByVal cType As Byte = 0) As String        '�ַ���md5
    Dim Bytes() As Byte
    '------------------������utf8Ϊ׼
    Select Case cType
        Case 1: Bytes = StrConv(strIn, vbFromUnicode) 'ansi
        Case 2: Bytes = strIn                         'unicode
        Case Else: Bytes = EncodeToBytes(strIn)       'utf-8
    End Select
    GetMD5Hash_String = GetMD5Hash_Bytes(Bytes) '
    '-----------------------------------StrConv(strIn, vbFromUnicode, 2052)�����ڼ����ʱ����Ҫע�ⲻͬ����Լ����Ӱ��
    '��unicode�ַ�������,ͳһ2���ֽ�,ansi,1Ӣ��ռ1���ֽ�, 1����2�ֽ�, utf8
    'UTF-8ʹ�ô�1~4�ֽ�Ϊÿ���ַ������ڣ�
    '1��һ��US-ASCIl�ַ�ֻ��1�ֽڱ��루Unicode��Χ��U+0000~U+007F����
    '2�����б������ݺŵ������ġ�ϣ���ġ��������ĸ�����������ϣ�����ġ��������ġ��������ĵ���ĸ����Ҫ2�ֽڱ��루Unicode��Χ��U+0080~U+07FF����
    '3���������Ե��ַ����������պ����֡����������֡��ж����ֵȣ������˴󲿷ֳ����֣�ʹ��3�ֽڱ��롣
    '4����������ʹ�õ������ַ�ʹ��4�ֽڱ��롣
End Function

Function GetMD5Hash_File(ByVal strFile As String, ByVal Px As Byte, Optional ByVal filez As Long) As String '֧�ַ�ansi����,�����ٶ��Ѿ��õ�����������
    Dim zx() As Byte, i As Long, j As Long, p As Long, n As Long, m As Long
    Dim obj As Object, iFile As Byte
    
    On Error GoTo 100
    If filez = 0 Then
        i = fso.GetFile(strFile).Size
        If i = 0 Then GetMD5Hash_File = "UC": Exit Function
    Else
        i = filez
    End If
    If Px = 1 Then     '·�������ڷ�ansi����
        iFile = FreeFile(0) '�����ļ���� 'https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/freefile-function
        ReDim zx(i - 1)
        Open strFile For Binary As iFile
        Get iFile, , zx
        Close iFile
        GetMD5Hash_File = GetMD5Hash_Bytes(zx)
    Else             '·�����ڷ�ansi����
        Ccx = 0                       '��������
        p = 131071   '΢����Ա���128k�Ķ�ȡЧ�������-'��ʵ�ʵ�С��ģ������ʾ128k��1024k��  '��������Դ-https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/readtext-method?view=sql-server-ver15
        j = p + 1    '128*1024
        m = i Mod j  '�Ƿ�����
        If m = 0 Then
            n = i \ j
        Else
            n = i \ j + 1   'ѭ���Ĵ��� ע��\��/����, \��ʾȡ��
        End If
        i = i - 1    '��0��ʼ
        ReDim xi(i)   '��ǰ���������Ŀռ�,����Ҫ��combinearr��redim preserve
        ReDim zx(p)
'        Set obj = CreateObject("ADODB.Stream")                      'https://support.microsoft.com/ms-my/help/276488/how-to-use-the-adodb-stream-object-to-send-binary-files-to-the-browser
        Set obj = New ADODB.Stream
        If obj Is Nothing Then GoTo 100 '����û�гɹ���������,��������md5ģ�����
        With obj
            .Mode = 3
            .type = adTypeBinary
            .Open
            .LoadFromFile (strFile)
            Do While n > 0 '������Ҫ�ĳ� endofstream
               '��ȫ�����е����ݵ�ʱ�����ִ���                       '�ڶ�ȡ��ɺ�,���᷵��"NUll"����ɴ���
                zx = .Read(j)          '���򵥵ķ����ǲ������read(1)����ȡ������'���ǲ������ַ��������ļ�hash�ٶ�̫�� ,Ŀǰ���õ��·����ܹ����ļ�md�ļ��������ϰٱ�
                Call CombineArr(zx)   '�����j���ļ���Сʱ,����������open����,ֱ�Ӷ�ȡ���е�����, ������j=i ������Open strFile For Binary As lFile 'Get lFile, , bytes,������adoģʽ��,���������׻���ִ���
                n = n - 1
            Loop                     '��300M+���ļ��ͳ����������/�洢�ռ䲻�������
            .Close
        End With
        GetMD5Hash_File = GetMD5Hash_Bytes(xi)
        Erase xi
        Set obj = Nothing
    End If
    Erase zx
Exit Function
100
    If Px = 1 Then
        GetMD5Hash_File = FileHashes(strFile) '����md5
    Else
        GetMD5Hash_File = HashPowershell(strFile) '����Powershell
    End If
    Erase xi
    Erase zx
    Set obj = Nothing
    Err.Clear
End Function

Private Function CombineArr(xcx() As Byte) '����ȡ������Ϣƴ����һ��
    Dim i As Long
    
    i = UBound(xcx)                   '����copymemory�ķ�ʽ������ʹ��ѭ����ɵ�����ʱ���˷�
    CopyMemory xi(Ccx), xcx(0), i + 1 'ƴ�ӵ����������(��ʼ)λ��,ƴ������Ŀ�ʼλ��,ƴ�ӵĳ���
    Ccx = Ccx + i + 1
End Function


