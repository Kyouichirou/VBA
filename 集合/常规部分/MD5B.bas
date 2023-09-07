Attribute VB_Name = "MD5B"
Option Explicit
Option Base 0
Private Type MD5_CTX
    z(1) As Long
    buf(3) As Long
    inc(63) As Byte
    digest(15) As Byte
End Type
'----------------------------------------------------------------------------------'Any 类型是有风险的
'-------------longptr并不是具体的数据类型，只是一种标记，在x64office下被转换为longlong数据类型， x32下则被转为long， 这是为了便于x64和x32之间的兼容
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
Dim Ccx As Long '拼接数组
 
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

Function GetMD5Hash_String(ByVal strIn As String, Optional ByVal cType As Byte = 0) As String        '字符串md5
    Dim Bytes() As Byte
    '------------------这里以utf8为准
    Select Case cType
        Case 1: Bytes = StrConv(strIn, vbFromUnicode) 'ansi
        Case 2: Bytes = strIn                         'unicode
        Case Else: Bytes = EncodeToBytes(strIn)       'utf-8
    End Select
    GetMD5Hash_String = GetMD5Hash_Bytes(Bytes) '
    '-----------------------------------StrConv(strIn, vbFromUnicode, 2052)这里在计算的时候需要注意不同编码对计算的影响
    '在unicode字符编码下,统一2个字节,ansi,1英文占1个字节, 1中文2字节, utf8
    'UTF-8使用答1~4字节为每个字符编码内：
    '1，一个US-ASCIl字符只需1字节编码（Unicode范围由U+0000~U+007F）。
    '2，带有变音符容号的拉丁文、希腊文、西里尔字母、亚美尼亚语、希伯来文、阿拉伯文、叙利亚文等字母则需要2字节编码（Unicode范围由U+0080~U+07FF）。
    '3，其他语言的字符（包括中日韩文字、东南亚文字、中东文字等）包含了大部分常用字，使用3字节编码。
    '4，其他极少使用的语言字符使用4字节编码。
End Function

Function GetMD5Hash_File(ByVal strFile As String, ByVal Px As Byte, Optional ByVal filez As Long) As String '支持非ansi编码,计算速度已经得到量级的提升
    Dim zx() As Byte, i As Long, j As Long, p As Long, n As Long, m As Long
    Dim obj As Object, iFile As Byte
    
    On Error GoTo 100
    If filez = 0 Then
        i = fso.GetFile(strFile).Size
        If i = 0 Then GetMD5Hash_File = "UC": Exit Function
    Else
        i = filez
    End If
    If Px = 1 Then     '路径不存在非ansi编码
        iFile = FreeFile(0) '生成文件编号 'https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/freefile-function
        ReDim zx(i - 1)
        Open strFile For Binary As iFile
        Get iFile, , zx
        Close iFile
        GetMD5Hash_File = GetMD5Hash_Bytes(zx)
    Else             '路径存在非ansi编码
        Ccx = 0                       '参数重置
        p = 131071   '微软测试表明128k的读取效率是最佳-'在实际的小规模测试显示128k比1024k好  '参数的来源-https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/readtext-method?view=sql-server-ver15
        j = p + 1    '128*1024
        m = i Mod j  '是否整除
        If m = 0 Then
            n = i \ j
        Else
            n = i \ j + 1   '循环的次数 注意\和/区别, \表示取商
        End If
        i = i - 1    '从0开始
        ReDim xi(i)   '提前分配好数组的空间,不需要在combinearr中redim preserve
        ReDim zx(p)
'        Set obj = CreateObject("ADODB.Stream")                      'https://support.microsoft.com/ms-my/help/276488/how-to-use-the-adodb-stream-object-to-send-binary-files-to-the-browser
        Set obj = New ADODB.Stream
        If obj Is Nothing Then GoTo 100 '假设没有成功创建对象,调用其他md5模块计算
        With obj
            .Mode = 3
            .type = adTypeBinary
            .Open
            .LoadFromFile (strFile)
            Do While n > 0 '这里需要改成 endofstream
               '完全读所有的数据的时候会出现错误                       '在读取完成后,将会返回"NUll"会造成错误
                zx = .Read(j)          '更简单的方法是采用逐个read(1)来读取二进制'但是采用这种方法计算文件hash速度太慢 ,目前采用的新方法能够让文件md的计算提升上百倍
                Call CombineArr(zx)   '这里的j在文件较小时,可以类似于open命令,直接读取所有的数据, 即设置j=i 类似于Open strFile For Binary As lFile 'Get lFile, , bytes,但是在ado模式下,这样很容易会出现错误
                n = n - 1
            Loop                     '在300M+的文件就出现内容溢出/存储空间不足的问题
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
        GetMD5Hash_File = FileHashes(strFile) '备用md5
    Else
        GetMD5Hash_File = HashPowershell(strFile) '调用Powershell
    End If
    Erase xi
    Erase zx
    Set obj = Nothing
    Err.Clear
End Function

Private Function CombineArr(xcx() As Byte) '将读取到的信息拼接在一起
    Dim i As Long
    
    i = UBound(xcx)                   '采用copymemory的方式将避免使用循环造成的严重时间浪费
    CopyMemory xi(Ccx), xcx(0), i + 1 '拼接到所在数组的(开始)位置,拼接数组的开始位置,拼接的长度
    Ccx = Ccx + i + 1
End Function


