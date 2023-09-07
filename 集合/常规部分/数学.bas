Attribute VB_Name = "数学"
Option Explicit
Private Declare Function Hash Lib "ntdll.dll" Alias "RtlComputeCrc32" (ByVal Start As Long, ByVal Data As Long, ByVal Size As Long) As Long 'crc32
Dim CRC32Table(255) As Long 'crc32

Function CRC32API(ByVal strx As String) As String '通过调用api的方式来实现计算字符串crc32
    Dim strx1 As String, i As Long
    'hex使用最大范围是long类型
    strx1 = StrConv(strx, vbFromUnicode)        'https://wenku.baidu.com/view/af430b3310661ed9ad51f376.html
    i = Hash(0, StrPtr(strx1), LenB(strx1))     'https://source.winehq.org/WineAPI/RtlComputeCrc32.html
    CRC32API = Hex$(i)                          'https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/hex-function
End Function

Function PoissonRand(ByVal lambda As Double) As Integer '泊松分布
    Dim Rand As Single
    Dim k As Integer
    Dim p As Single
    Dim sump As Single
    
    Randomize
    Rand = Rnd
    k = 0
    'p(0)化简
    p = 1 / exp(lambda)
    Do While Rand > sump
        k = k + 1
        '从p(k)转换到p(k+1)
        p = p * lambda / k
        sump = sump + p
    Loop
    PoissonRand = k
End Function

Sub Matchx1() '简单算法的演示
    Dim yesno As Variant
    Dim a As Integer, b As Integer, c As Long, strx As String
    Dim t As Single
    
    yesno = MsgBox("运算需要较长时间,是否运行(About:20+s)?_", vbYesNo) '如果无法正常连接存储文件的处理
    If yesno = vbYes Then
    With UserForm3
    .Label57.Caption = "处理中..."
    .Label100.Caption = ""
    .Label101.Caption = ""
    t = Timer
    DoEvents
    For a = 1 To 1000
        For b = 1 To 1000
            For c = 1 To 1000    '修改部分
                If a + b + c = 1000 And CLng(a) * a + CLng(b) * b = CLng(c) * c Then strx = str(a) + str(b) + str(c)   'Debug.Print a, b, c
            Next
        Next
    Next
        .Label100.Caption = strx
        .Label101.Caption = Format(Timer - t, "0.0000") & "s"
        .Label57.Caption = "处理完成"
    End With
    End If
    'Debug.Print Timer - t
End Sub

Sub Matchx2() '算法的演示
    Dim a As Integer, b As Integer, c As Long
    Dim t As Single, strx As String
    
    t = Timer        'timer函数, 计算从午夜开始到现在的时间
    With UserForm3
    .Label100.Caption = ""
    .Label101.Caption = ""
    For a = 1 To 1000
        For b = 1 To 1000
            c = 1000 - a - b '只需要简单的修改,即可实现量级计算速度提升
            If c > 0 And CLng(a) * a + CLng(b) * b = CLng(c) * c Then strx = str(a) + str(b) + str(c) 'Debug.Print a, b, c '在运算时,注意运算的值超出运算值的类型integer的数据范围
        Next
    Next
        .Label100.Caption = strx
        .Label101.Caption = Format(Timer - t, "0.0000") & "s"
        .Label57.Caption = "处理完成"
    End With
    'Debug.Print Timer - t
End Sub

Function CheckPN(ByVal numx As Integer) As Boolean '判断10000以内的质数/限制更多的计算浪费
    Dim n As Integer, i As Integer
    
    If numx < 2 Or numx > 10000 Then Exit Function
    CheckPN = True
    n = Int(Sqr(numx))                              'https://docs.microsoft.com/zh-CN/office/vba/Language/Reference/User-Interface-Help/sqr-function
    For i = 2 To n
        If numx / i = numx \ i Then CheckPN = False: Exit Function
    Next
End Function

Function PasswordGR(ByVal modex As Byte, Optional ByVal numx As Byte) As String '模式-数字,数字+字母(大小写),数字+字符+字母
    Dim i As Byte, k As Byte, p As Byte, xi As Byte, j As Byte, m As Byte, n As Byte, q As Long, c As Byte, r As Byte
    Dim Password As String, strx As String
    
    If numx <> 6 And numx <> 12 And numx <> 18 Then PasswordGR = "Err": Exit Function
    Select Case modex
        Case 0 '纯数字
        For p = 1 To numx
100
            r = r + 1 '控制执行的次数
            i = Int((9 - 0 + 1) * Rnd + 0)
            strx = CStr(i)
            If numx = 6 And r < 255 Then             '当只产生6位密码时,不允许产生重复值
                If InStr(Password, strx) > 0 Then GoTo 100
            End If
            Password = Password & strx
        Next
        Case 1 '数字+字母(大小写)
101
        r = r + 1
        For p = 1 To numx
            q = RandNumx(2000000000) '在这个范围产生一个任意随机数
            xi = q Mod 3 '获取余数
            If xi = 0 Then
                i = Int((9 - 0 + 1) * Rnd + 0)
                m = m + 1
            ElseIf xi = 1 Then
                k = 1
                i = Int((90 - 65 + 1) * Rnd + 65) '大写字母
                n = n + 1
            ElseIf xi = 2 Then
                k = 1
                i = Int((122 - 97 + 1) * Rnd + 97) '小写字母
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
            If m = 0 Or n = 0 Or j = 0 Then m = 0: n = 0: j = 0: strx = "": Password = "": GoTo 101   '每种类型都要有
        End If
        Case 2         '数字+特殊字符+字母(大小写)
102
        r = r + 1
        For p = 1 To numx
            q = RandNumx(200000000) '在这个范围产生一个任意随机数
            xi = q Mod 4 '获取余数
            If xi = 0 Then
                i = Int((9 - 0 + 1) * Rnd + 0)
                m = m + 1
            ElseIf xi = 1 Then
                k = 1
                i = Int((90 - 65 + 1) * Rnd + 65) '大写字母
                n = n + 1
            ElseIf xi = 2 Then
                k = 1
                i = Int((122 - 97 + 1) * Rnd + 97) '小写字母
                j = j + 1
            ElseIf xi = 3 Then
                c = c + 1
                k = 1
                xi = q Mod 5
                Select Case xi
                    Case 0: i = 33
                    Case 1: i = 35 '可以修改增加更多的可选字符(这几个是较容易记的字符)
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
        '-------每种类型都要有 -(只是让密码"看"起来更"复杂",并不会实质(长度远比字符类型重要)提升安全性,假设破解方知道密码的生成方式,还会降低安全性,概率变成x*(x-1)*...(x-n)<x^n)
        End If
        Case Else
        PasswordGR = "Err"
        Exit Function
    End Select
    PasswordGR = Password '利用这个function可以和ZipCompress组成一个自动压缩文件和加密的小程序
    '压缩包密码的破解相对可行的方法为明文破解,即知道压缩包的某个文件(crc32),利用这个文件反向破解压缩包的密码,所以加密文件的时候应该加密文件头(无法直接打开压缩包看到压缩内的文件信息)
End Function

Function Base64_HMACSHA1(ByVal sTextToHash As String, ByVal sSharedSecretKey As String) As String '-备用
    Dim asc As Object, enc As Object
    Dim TextToHash() As Byte
    Dim Bytes() As Byte
    Dim SharedSecretKey() As Byte
    
    Set asc = CreateObject("System.Text.UTF8Encoding")
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA1")
    If asc Is Nothing Or enc Is Nothing Then Base64_HMACSHA1 = "程序异常": Exit Function
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
Function CRC32(ByVal item As String) As String '-crc32校检-备用
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

Function SHA256Function(ByVal sMessage As String) As String '调用sha256
    Dim clsX As CSHA256
    
    Set clsX = New CSHA256
    If clsX Is Nothing Then SHA256Function = "程序异常": Exit Function
    SHA256Function = clsX.SHA256(sMessage)
    Set clsX = Nothing
End Function

Function RandNumx(ByVal numx As Long) As Long '生成随机数'https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/randomize-statement
    Randomize (Timer)
    RandNumx = Int((numx - 0 + 1) * Rnd + 0)
End Function

Sub Crimx() '二进制用于寻找特定的目标的例子
'假设100人, 排列报号,奇数者剔除,剩下继续报号,重复之前步奏,最后剩下一个,哪个是最后剩下的?
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
