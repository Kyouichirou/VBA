Attribute VB_Name = "��ѧ"
Option Explicit
Private Declare Function Hash Lib "ntdll.dll" Alias "RtlComputeCrc32" (ByVal Start As Long, ByVal Data As Long, ByVal Size As Long) As Long 'crc32
Dim CRC32Table(255) As Long 'crc32

Function CRC32API(ByVal strx As String) As String 'ͨ������api�ķ�ʽ��ʵ�ּ����ַ���crc32
    Dim strx1 As String, i As Long
    'hexʹ�����Χ��long����
    strx1 = StrConv(strx, vbFromUnicode)        'https://wenku.baidu.com/view/af430b3310661ed9ad51f376.html
    i = Hash(0, StrPtr(strx1), LenB(strx1))     'https://source.winehq.org/WineAPI/RtlComputeCrc32.html
    CRC32API = Hex$(i)                          'https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/hex-function
End Function

Function PoissonRand(ByVal lambda As Double) As Integer '���ɷֲ�
    Dim Rand As Single
    Dim k As Integer
    Dim p As Single
    Dim sump As Single
    
    Randomize
    Rand = Rnd
    k = 0
    'p(0)����
    p = 1 / exp(lambda)
    Do While Rand > sump
        k = k + 1
        '��p(k)ת����p(k+1)
        p = p * lambda / k
        sump = sump + p
    Loop
    PoissonRand = k
End Function

Sub Matchx1() '���㷨����ʾ
    Dim yesno As Variant
    Dim a As Integer, b As Integer, c As Long, strx As String
    Dim t As Single
    
    yesno = MsgBox("������Ҫ�ϳ�ʱ��,�Ƿ�����(About:20+s)?_", vbYesNo) '����޷��������Ӵ洢�ļ��Ĵ���
    If yesno = vbYes Then
    With UserForm3
    .Label57.Caption = "������..."
    .Label100.Caption = ""
    .Label101.Caption = ""
    t = Timer
    DoEvents
    For a = 1 To 1000
        For b = 1 To 1000
            For c = 1 To 1000    '�޸Ĳ���
                If a + b + c = 1000 And CLng(a) * a + CLng(b) * b = CLng(c) * c Then strx = str(a) + str(b) + str(c)   'Debug.Print a, b, c
            Next
        Next
    Next
        .Label100.Caption = strx
        .Label101.Caption = Format(Timer - t, "0.0000") & "s"
        .Label57.Caption = "�������"
    End With
    End If
    'Debug.Print Timer - t
End Sub

Sub Matchx2() '�㷨����ʾ
    Dim a As Integer, b As Integer, c As Long
    Dim t As Single, strx As String
    
    t = Timer        'timer����, �������ҹ��ʼ�����ڵ�ʱ��
    With UserForm3
    .Label100.Caption = ""
    .Label101.Caption = ""
    For a = 1 To 1000
        For b = 1 To 1000
            c = 1000 - a - b 'ֻ��Ҫ�򵥵��޸�,����ʵ�����������ٶ�����
            If c > 0 And CLng(a) * a + CLng(b) * b = CLng(c) * c Then strx = str(a) + str(b) + str(c) 'Debug.Print a, b, c '������ʱ,ע�������ֵ��������ֵ������integer�����ݷ�Χ
        Next
    Next
        .Label100.Caption = strx
        .Label101.Caption = Format(Timer - t, "0.0000") & "s"
        .Label57.Caption = "�������"
    End With
    'Debug.Print Timer - t
End Sub

Function CheckPN(ByVal numx As Integer) As Boolean '�ж�10000���ڵ�����/���Ƹ���ļ����˷�
    Dim n As Integer, i As Integer
    
    If numx < 2 Or numx > 10000 Then Exit Function
    CheckPN = True
    n = Int(Sqr(numx))                              'https://docs.microsoft.com/zh-CN/office/vba/Language/Reference/User-Interface-Help/sqr-function
    For i = 2 To n
        If numx / i = numx \ i Then CheckPN = False: Exit Function
    Next
End Function

Function PasswordGR(ByVal modex As Byte, Optional ByVal numx As Byte) As String 'ģʽ-����,����+��ĸ(��Сд),����+�ַ�+��ĸ
    Dim i As Byte, k As Byte, p As Byte, xi As Byte, j As Byte, m As Byte, n As Byte, q As Long, c As Byte, r As Byte
    Dim Password As String, strx As String
    
    If numx <> 6 And numx <> 12 And numx <> 18 Then PasswordGR = "Err": Exit Function
    Select Case modex
        Case 0 '������
        For p = 1 To numx
100
            r = r + 1 '����ִ�еĴ���
            i = Int((9 - 0 + 1) * Rnd + 0)
            strx = CStr(i)
            If numx = 6 And r < 255 Then             '��ֻ����6λ����ʱ,����������ظ�ֵ
                If InStr(Password, strx) > 0 Then GoTo 100
            End If
            Password = Password & strx
        Next
        Case 1 '����+��ĸ(��Сд)
101
        r = r + 1
        For p = 1 To numx
            q = RandNumx(2000000000) '�������Χ����һ�����������
            xi = q Mod 3 '��ȡ����
            If xi = 0 Then
                i = Int((9 - 0 + 1) * Rnd + 0)
                m = m + 1
            ElseIf xi = 1 Then
                k = 1
                i = Int((90 - 65 + 1) * Rnd + 65) '��д��ĸ
                n = n + 1
            ElseIf xi = 2 Then
                k = 1
                i = Int((122 - 97 + 1) * Rnd + 97) 'Сд��ĸ
                j = j + 1
            End If
            If k = 0 Then
                strx = CStr(i)
            Else
                strx = Chr(i)
                k = 0
            End If
            Password = Password & strx
        Next
        If r < 255 Then
            If m = 0 Or n = 0 Or j = 0 Then m = 0: n = 0: j = 0: strx = "": Password = "": GoTo 101   'ÿ�����Ͷ�Ҫ��
        End If
        Case 2         '����+�����ַ�+��ĸ(��Сд)
102
        r = r + 1
        For p = 1 To numx
            q = RandNumx(200000000) '�������Χ����һ�����������
            xi = q Mod 4 '��ȡ����
            If xi = 0 Then
                i = Int((9 - 0 + 1) * Rnd + 0)
                m = m + 1
            ElseIf xi = 1 Then
                k = 1
                i = Int((90 - 65 + 1) * Rnd + 65) '��д��ĸ
                n = n + 1
            ElseIf xi = 2 Then
                k = 1
                i = Int((122 - 97 + 1) * Rnd + 97) 'Сд��ĸ
                j = j + 1
            ElseIf xi = 3 Then
                c = c + 1
                k = 1
                xi = q Mod 5
                Select Case xi
                    Case 0: i = 33
                    Case 1: i = 35 '�����޸����Ӹ���Ŀ�ѡ�ַ�(�⼸���ǽ����׼ǵ��ַ�)
                    Case 2: i = 36
                    Case 3: i = 37
                    Case 4: i = 64
                End Select
            End If
            If k = 0 Then
                strx = CStr(i)
            Else
                strx = Chr(i)
                k = 0
            End If
            Password = Password & strx
        Next
        If r < 255 Then
            If m = 0 Or n = 0 Or j = 0 Or c = 0 Then m = 0: n = 0: j = 0: c = 0: strx = "": Password = "": GoTo 102
        '-------ÿ�����Ͷ�Ҫ�� -(ֻ��������"��"������"����",������ʵ��(����Զ���ַ�������Ҫ)������ȫ��,�����ƽ֪ⷽ����������ɷ�ʽ,���ή�Ͱ�ȫ��,���ʱ��x*(x-1)*...(x-n)<x^n)
        End If
        Case Else
        PasswordGR = "Err"
        Exit Function
    End Select
    PasswordGR = Password '�������function���Ժ�ZipCompress���һ���Զ�ѹ���ļ��ͼ��ܵ�С����
    'ѹ����������ƽ���Կ��еķ���Ϊ�����ƽ�,��֪��ѹ������ĳ���ļ�(crc32),��������ļ������ƽ�ѹ����������,���Լ����ļ���ʱ��Ӧ�ü����ļ�ͷ(�޷�ֱ�Ӵ�ѹ��������ѹ���ڵ��ļ���Ϣ)
End Function

Function Base64_HMACSHA1(ByVal sTextToHash As String, ByVal sSharedSecretKey As String) As String '-����
    Dim asc As Object, enc As Object
    Dim TextToHash() As Byte
    Dim Bytes() As Byte
    Dim SharedSecretKey() As Byte
    
    Set asc = CreateObject("System.Text.UTF8Encoding")
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA1")
    If asc Is Nothing Or enc Is Nothing Then Base64_HMACSHA1 = "�����쳣": Exit Function
    TextToHash = asc.GetBytes_4(sTextToHash)
    SharedSecretKey = asc.GetBytes_4(sSharedSecretKey)
    enc.key = SharedSecretKey
    Bytes = enc.ComputeHash_2((TextToHash))
    Base64_HMACSHA1 = EncodeBase64(Bytes)
    Set asc = Nothing
    Set enc = Nothing
End Function

Private Function EncodeBase64(ByRef arrData() As Byte) As String
    Dim objXML As MSXML2.DOMDocument60
    Dim objNode As MSXML2.IXMLDOMElement
    'https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms764622%28v%3dvs.85%29
    'https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms753804(v=vs.85)
    Set objXML = New MSXML2.DOMDocument60
    ' byte array to base64
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.Text

    Set objNode = Nothing
    Set objXML = Nothing
End Function

'-----------------------------------------------------------------------crc32
Function CRC32(ByVal item As String) As String '-crc32У��-����
    Dim i As Long, iCRC As Long, lngA As Long, ret As Long
    Dim b() As Byte, bytT As Byte, bytC As Byte
    
    b = StrConv(item, vbFromUnicode)
    InitCrc32
    iCRC = &HFFFFFFFF
    For i = 0 To UBound(b)
        bytC = b(i)
        bytT = (iCRC And &HFF) Xor bytC
        lngA = ((iCRC And &HFFFFFF00) / &H100) And &HFFFFFF
        iCRC = lngA Xor CRC32Table(bytT)
    Next
    ret = iCRC Xor &HFFFFFFFF
    CRC32 = Right("00000000" & Hex(ret), 8)
End Function

Private Function InitCrc32(Optional ByVal Seed As Long = &HEDB88320, Optional ByVal Precondition As Long = &HFFFFFFFF) As Long
    Dim i As Integer, j As Integer, CRC32 As Long, temp As Long
    
    For i = 0 To 255
        CRC32 = i
        For j = 0 To 7
            temp = ((CRC32 And &HFFFFFFFE) / &H2) And &H7FFFFFFF
            If (CRC32 And &H1) Then CRC32 = temp Xor Seed Else CRC32 = temp
        Next
        CRC32Table(i) = CRC32
    Next
    InitCrc32 = Precondition
End Function
'-----------------------------------------------------------------------crc32

Function SHA256Function(ByVal sMessage As String) As String '����sha256
    Dim clsX As CSHA256
    
    Set clsX = New CSHA256
    If clsX Is Nothing Then SHA256Function = "�����쳣": Exit Function
    SHA256Function = clsX.SHA256(sMessage)
    Set clsX = Nothing
End Function

Function RandNumx(ByVal numx As Long) As Long '���������'https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/randomize-statement
    Randomize (Timer)
    RandNumx = Int((numx - 0 + 1) * Rnd + 0)
End Function

Sub Crimx() '����������Ѱ���ض���Ŀ�������
'����100��, ���б���,�������޳�,ʣ�¼�������,�ظ�֮ǰ����,���ʣ��һ��,�ĸ������ʣ�µ�?
    Dim i As Byte, k As Byte, m As Byte, n As Byte, p As Byte, xi As Byte, j As Byte
    k = 1
    xi = 1
    For i = 1 To 100
        p = i
        Do
            k = k + 1
            m = Int(p / 2)
            n = p Mod 2
            p = m
            If n = 0 Then xi = xi + 1
        Loop Until m = 0 And n = 1
        If k - xi = 1 Then
            If i > j Then j = i
        End If
        k = 1
        xi = 1
    Next
    Debug.Print j
End Sub
