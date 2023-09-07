VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HashA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------------------------https://walkccc.github.io/CLRS/
'-------------------------https://blog.csdn.net/Beyond_2016/article/details/81286360
'.NET Core
'3.1 3.0 2.2 2.1 2.0 1.1 1.0
'.NET Framework
'4.8 4.7.2 4.7.1 4.7 4.6.2 4.6.1 4.6 4.5.2 4.5.1 4.5 4.0 3.5 3.0 2.0 1.1
'.NET Standard
'2.1 2.0 1.6 1.4 1.3
'--------------------存在一个很大的问题, 无法有效释放, 在整体计算大文件时候会出现内存溢出的问题(改成计算小块(不需要dim大型数组), 就算无法有效释放内存也不会造成内存溢出的问题)
Private Declare Function Hash Lib "ntdll.dll" Alias "RtlComputeCrc32" (ByVal Start As Long, ByVal Data As Long, ByVal Size As Long) As Long 'crc32
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'--------------------------------https://docs.microsoft.com/en-us/dotnet/api/system.security.cryptography.hashalgorithm?view=netframework-4.8
Private Const MD5 As String = "System.Security.Cryptography.MD5CryptoServiceProvider"
Private Const SHA1 As String = "System.Security.Cryptography.SHA1Managed"
Private Const SHA256 As String = "System.Security.Cryptography.SHA256Managed"
Private Const SHA384 As String = "System.Security.Cryptography.SHA384Managed"
Private Const SHA512 As String = "System.Security.Cryptography.SHA512Managed"
Private Const UTF8x As String = "System.Text.UTF8Encoding" '处理字符
Private Const blockSize As Long = 131072  '(2^17) '131071'(2^17-1) '可修改参数(不同的计算硬件可能计算速度不一样)
'Private Bytes() As Byte
'Private Ccx As Long '拼接数组
Private objHash As Object
Private objUTF8 As Object
Private ado As Object

Friend Function CRC32Hash(ByVal strText As String, ByVal IsFile As Boolean, Optional ByVal rType As Boolean = False) As String '通过调用api的方式来实现计算字符串crc32
    Dim strx1 As String, i As Long
    'hex使用最大范围是long类型
    If IsFile = False Then
        strx1 = StrConv(strText, vbFromUnicode)        'https://wenku.baidu.com/view/af430b3310661ed9ad51f376.html
        i = Hash(0, StrPtr(strx1), LenB(strx1))     'https://source.winehq.org/WineAPI/RtlComputeCrc32.html
        If rType = True Then
            CRC32Hash = LCase$(Hex$(i))
        Else
            CRC32Hash = Hex$(i)                      'https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/hex-function
        End If
    Else
        CRC32Hash = "不支持文件类型"
    End If
End Function
'可以进一步修改, 将属性设置单独分离出来, 作为整体使用, byval 用 byref 替换掉

Friend Function MD5Hash(ByVal strText As String, ByVal IsFile As Boolean, Optional ByVal IsExist As Boolean = False, Optional ByVal filez As Long, _
Optional ByVal cmCodex As Integer = 0, Optional ByVal rType As Boolean = False) As String

    MD5Hash = GetHashx(strText, IsFile, MD5, IsExist, filez, cmCodex, rType)
End Function

Friend Function SHA1Hash(ByVal strText As String, ByVal IsFile As Boolean, Optional ByVal IsExist As Boolean = False, Optional ByVal filez As Long, _
Optional ByVal cmCodex As Integer = 0, Optional ByVal rType As Boolean = False) As String

    SHA1Hash = GetHashx(strText, IsFile, SHA1, IsExist, filez, cmCodex, rType)
End Function

Friend Function SHA256Hash(ByVal strText As String, ByVal IsFile As Boolean, Optional ByVal IsExist As Boolean = False, Optional ByVal filez As Long, _
Optional ByVal cmCodex As Integer = 0, Optional ByVal rType As Boolean = False) As String

    SHA256Hash = GetHashx(strText, IsFile, SHA256, IsExist, filez, cmCodex, rType)
End Function

Friend Function SHA384Hash(ByVal strText As String, ByVal IsFile As Boolean, Optional ByVal IsExist As Boolean = False, Optional ByVal filez As Long, _
Optional ByVal cmCodex As Integer = 0, Optional ByVal rType As Boolean = False) As String

    SHA384Hash = GetHashx(strText, IsFile, SHA384, IsExist, filez, cmCodex, rType)
End Function

Friend Function SHA512Hash(ByVal strText As String, ByVal IsFile As Boolean, Optional ByVal IsExist As Boolean = False, Optional ByVal filez As Long, _
Optional ByVal cmCodex As Integer = 0, Optional ByVal rType As Boolean = False) As String

    SHA512Hash = GetHashx(strText, IsFile, SHA512, IsExist, filez, cmCodex, rType)
End Function

Private Function GetHashx(ByVal strText As String, ByVal IsFile As Boolean, ByVal Algorithmx As String, Optional ByVal IsExist As Boolean, Optional filez As Long, _
Optional ByVal cmCodex As Integer, Optional ByVal rType As Boolean) As String
    Dim Hash() As Byte, hFile As Byte, i As Integer, k As Integer
    Dim Buffer() As Byte, Ublock As Long
    Dim Result As String
    
    On Error GoTo ErrHandle
    If LenB(strText) = 0 Then GetHashx = "UC": Exit Function
    If objHash Is Nothing Then Set objHash = CreateObject(Algorithmx)
    If IsFile = True Then '---------------------------------------------选择文件或者字符串
        If IsExist = False Then
            If fso.fileexists(strText) = False Then GetHashx = "UC": Exit Function '判断文件是否存在
        End If
        If cmCodex = 0 Then
            cmCodex = ErrCode(strText, 1) '-------------------检查是否包含非ansi字符
            If cmCodex = -1 Then GetHashx = "UC": Exit Function
        End If
        If cmCodex = 1 Then
         '-----------------------------前期检查
            hFile = FreeFile(0)
            Open strText For Binary Access Read As hFile 'https://docs.microsoft.com/zh-cn/office/vba/Language/Reference/User-Interface-Help/open-statement
            If filez = 0 Then
                filez = LOF(hFile) '--------------------------返回文件大小'https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/lof-function
                If filez = 0 Then GetHashx = "UC": Exit Function
            End If
            If filez > blockSize Then
                ReDim Buffer(blockSize - 1)
                k = filez \ blockSize
                For i = 1 To k
                    Get hFile, , Buffer
                    objHash.TransformBlock Buffer, 0, blockSize, Buffer, 0 '计算整块
                    'Computes the hash value for the specified region of the input byte array and copies the specified region of the input byte array to
                    'the specified region of the output byte array.
                Next
            Else
                ReDim Buffer(filez - 1)
            End If
            Get hFile, , Buffer
            Close hFile
            objHash.TransformFinalBlock Buffer, 0, filez Mod blockSize '计算剩余部分--(Computes the hash value for the specified region of the specified byte array.)
            '---------------------https://docs.microsoft.com/en-us/dotnet/api/system.security.cryptography.hashalgorithm.transformblock?view=netcore-3.1
            'Parameters-TransformBlock (byte[] inputBuffer, int inputOffset, int inputCount, byte[] outputBuffer, int outputOffset
            'inputBuffer
            'Byte[]
            'The input to compute the hash code for.
            'inputOffset
            'Int32
            'The offset into the input byte array from which to begin using data.
            'inputCount
            'Int32
            'The number of bytes in the input byte array to use as data.
            'outputBuffer
            'Byte[]
            'A copy of the part of the input array used to compute the hash code.
            'outputOffset
            'Int32
            'The offset into the output byte array from which to begin writing data.
            'TransformBlock-----------------------------------------------------参数设置
            'inputBuffer
            'Byte[]
            'The input to compute the hash code for.
            'inputOffset
            'Int32
            'The offset into the byte array from which to begin using data.
            'inputCount
            'Int32
            'The number of bytes in the byte array to use as data.
            'TransformFinalBlock-----------------------------------------------参数设置
        Else
            If filez = 0 Then
                filez = fso.GetFile(strText).Size '---------------注意不要使用filelen函数,这个函数的局限性很多,不支持非ansi字符,有文件大小的上限(好像几百M)
                If filez = 0 Then GetHashx = "UC": Exit Function
            End If
            If ado Is Nothing Then Set ado = CreateObject("adodb.stream")
            With ado '---支持(.size)
                .Mode = 3  '不能选1,只读模式/,3 读写模式
                .type = 1  'adTypeBinary 'adTypeText=2
                .Open
                .LoadFromFile (strText)
                .Position = 0 '-------------开始位置
                If filez > blockSize Then
                    ReDim Buffer(blockSize - 1)
                    Do Until .EOS = True '-----读取数据的末尾为止
                        Buffer = .Read(blockSize) 'read读取后将自动读取到后续的内容,如果需要重新开始读取,需要用position来调整位置
                        Ublock = UBound(Buffer) + 1
                        If Ublock = blockSize Then
                            objHash.TransformBlock Buffer, 0, blockSize, Buffer, 0
                        Else
                            objHash.TransformFinalBlock Buffer, 0, Ublock
                        End If
                    Loop
                Else
                    ReDim Buffer(filez - 1)
                    Buffer = .Read(filez) '如果文件比128k小/或等于,就直接读取全部数据
                    objHash.TransformFinalBlock Buffer, 0, filez
                End If
                .Close
            End With
        End If
    Else
        If objUTF8 Is Nothing Then Set objUTF8 = CreateObject(UTF8x) '计算字符串
        Buffer = objUTF8.GetBytes_4(strText)
        objHash.ComputeHash_2 Buffer
    End If
    '-------------------------------处理返回的结果
    Hash = objHash.Hash
    k = UBound(Hash)
    objHash.Clear
    For i = 0 To k             '---"0"作为补位
        Result = Result & Right$("0" & Hex(Hash(i)), 2) '转换为16进制 $符号表示强制字符串 ,mid/left/right同样可以处理非字符串类型的数据 'hash(i) > 15(16进制,之前都是1位数表示)
    Next
    GetHashx = Result
    If rType = True Then GetHashx = LCase(GetHashx) '是否输出小写
    Erase Hash
    Erase Buffer
    Exit Function
ErrHandle:
    If IsFile = True Then
        GetHashx = GetFileHashMD5(strText, cmCodex) '调用其他的模块来计算
    Else
        GetHashx = GetMD5Hash_String(strText)
    End If
    Erase Buffer
End Function

Private Sub Class_Terminate() '将对象设置为模块级别的变量,减少反复创建对象的时间,在任务完成之后才进行销毁
    objHash.Clear
    Set objHash = Nothing
    If Not objUTF8 Is Nothing Then Set objUTF8 = Nothing
    If Not ado Is Nothing Then Set ado = Nothing
End Sub
'----------------------------------------------------------------------------------------------------------------------------------废弃
'Private Function GetHashx(ByVal strText As String, ByVal IsFile As Boolean, ByVal Algorithmx As String, Optional ByVal IsExist As Boolean, Optional Filez As Long, _
'Optional ByVal cmCodex As Integer, Optional ByVal RType As Boolean) As String
'    Dim Hash() As Byte
'    Dim Arrx() As Byte
'    Dim i As Integer, k As Integer
'    Dim Result As String
'    Dim iFile As Byte, BlockSizex As Long
'
'    If Len(strText) = 0 Then GetHashx = "UC": Exit Function
'    If IsFile = True Then '选择文件或者字符串
'        If IsExist = False Then
'            If fso.FileExists(strText) = False Then GetHashx = "UC": Exit Function '判断文件是否存在
'        End If
'        If Filez = 0 Then
'            Filez = fso.GetFile(strText).Size '---------------注意不要使用filelen函数,这个函数的局限性很多,不支持非ansi字符,有文件大小的上限(好像几百M)
'            If Filez = 0 Then GetHashx = "UC": Exit Function
'        End If
'        If cmCodex = 0 Then
'            cmCodex = ErrCode(strText, 1) '-------------------检查是否包含非ansi字符
'            If cmCodex = -1 Then GetHashx = "UC": Exit Function
'        End If
'        ReDim Bytes(Filez - 1)
'        '-------------------------------------------------前期判断
'        If cmCodex = 1 Then '----是否包含非ansi字符
'            iFile = FreeFile(0) '生成文件编号 0生成 1-255, 1生成 256-511文件号 'https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/freefile-function
'            Open strText For Binary As iFile
'            Get iFile, , Bytes   'open方法的好处是,可以一次读取所有数据/但是不支持非ansi 'https://docs.microsoft.com/zh-cn/office/vba/Language/Reference/User-Interface-Help/open-statement
'            Close iFile
'        Else
'            BlockSizex = blockSize + 1 '读取128K
'            If Ado Is Nothing Then Set Ado = CreateObject("adodb.stream")
'            With Ado
'                .Mode = 3  '不能选1,只读模式/,3 读写模式
'                .type = 1  'adTypeBinary 'adTypeText=2
'                .Open
'                .LoadFromFile (strText)
'                .Position = 0 '-------------开始位置
'                If Filez > BlockSizex Then
'                    ReDim Arrx(blockSize)
'                    Do Until .EOS = True '-----读取数据的末尾为止
'                        Arrx = .Read(BlockSizex) 'read读取后将自动读取到后续的内容,如果需要重新开始读取,需要用position来调整位置
'                        CombineArr Arrx
'                    Loop
'                    Ccx = 0
'                Else
'                    Bytes = .Read(Filez) '如果文件比128k小/或等于,就直接读取全部数据
'                End If
'                .Close
'            End With
'        End If
'    Else
'        If objUTF8 Is Nothing Then Set objUTF8 = CreateObject(UTF8x) '--https://docs.microsoft.com/en-us/dotnet/api/system.text.utf8encoding?view=netframework-4.8
'        Bytes = objUTF8.GetBytes_4(strText) '---------------------------https://docs.microsoft.com/en-us/dotnet/api/system.text.utf8encoding.getbytes?view=netframework-4.8
'    End If
'    '-------------------------------------处理方式
'    If objHash Is Nothing Then Set objHash = CreateObject(Algorithmx)
'    Hash = objHash.ComputeHash_2((Bytes)) '----------ComputeHash_1(inputStream As stream),ComputeHash_2(buffer As byte()),ComputeHash_3(buffer As byte(), offset As int, count As int)
'    objHash.Clear
'    k = UBound(Hash) + 1                                               'ascii编码字节数=字符串长度=字符个数，一个英文字母占一个字节的空间，一个中文汉字占两个字节的空间
                                                                        '在Unicode下，字节数/2=字符串长度=字符个数，一个英文等于两个字节，一个中文（含繁体）等于两个字节
''                                                                      'AscB函数作用于包含在字符串中的字节数据，AscB提取字符串string第一个字节，返回第一个字节的字符代码，而非字符的字符代码。
'    For i = 1 To k                                                     'Returns a String representing the hexadecimal value of a number.-hex
'        Result = Result & Right$("0" & Hex(AscB(MidB(Hash, i, 1))), 2) 'GetHashx = String$(32, "0"),Mid$(GetHashx, i + i + 2 + (hash(i) > 15)) = Hex(hash(i))'http://www.officetanaka.net/excel/vba/function/AscB.htm
'        '---------------------------------------------------------------https://docs.microsoft.com/en-us/office/vba/Language/Reference/user-interface-help/hex-function
'        '---------------------------------------------------------------https://www.engram9.info/visual-basic-vba/mid-mid-midb-midb-functions.html
'    Next
'    If RType = True Then '是否输出小写
'        GetHashx = LCase(Result)
'    Else
'        GetHashx = Result
'    End If
'    Erase Hash
'    Erase Bytes
'End Function
'
'Private Function CombineArr(ByRef xcx() As Byte) '将读取到的信息拼接在一起, ado模式读取数据,对于大文件无法一次完整读取(300M多出现错误),需要多次读取
'    Dim i As Long
'    '-------------------------------------注意这里i不能直接使用blocksize(ReDim Arrx(BlockSize)),xcx会动态变化
'    i = UBound(xcx)                      '采用copymemory的方式将避免使用循环造成的严重时间浪费(计算300M的文件,可以压缩从7s压倒1s左右)
'    CopyMemory Bytes(Ccx), xcx(0), i + 1 '拼接到所在数组的(开始)位置,拼接数组的开始位置,拼接的长度
'    Ccx = Ccx + i + 1
'End Function
