Attribute VB_Name = "Hash"
Option Explicit
'---------------和类模块Hashx实现的功能是相同的,只是写法的不同
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'--------------------------------https://docs.microsoft.com/en-us/dotnet/api/system.security.cryptography.hashalgorithm?view=netframework-4.8
Private Const md5x As String = "System.Security.Cryptography.MD5CryptoServiceProvider"
Private Const SHA1x As String = "System.Security.Cryptography.SHA1Managed"
Private Const SHA256x As String = "System.Security.Cryptography.SHA256Managed"
Private Const SHA384x As String = "System.Security.Cryptography.SHA384Managed"
Private Const SHA512x As String = "System.Security.Cryptography.SHA512Managed"
Private Const UTF8x As String = "System.Text.UTF8Encoding"
'---------------------------------------https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/enum-statement
Public Enum Algorithm '选择的算法类型
    MD5 = 0
    SHA1 = 1
    SHA256 = 2
    SHA384 = 3
    SHA512 = 4
End Enum
Private Const blockSize As Long = 131071
Dim Bytex() As Byte
Dim Ccx As Long '拼接数组

Private Function GetHash(ByVal strText As String, Optional ByVal Algorithmx As Byte = Algorithm.MD5, Optional ByVal IsFile As Boolean = False, Optional ByVal filez As Long, Optional ByVal cmCodex As Boolean = True, Optional ByVal rType As Boolean = True) As String
    Dim objHash As Object
    Dim objUTF8 As Object
    Dim Hash() As Byte
    Dim arrx() As Byte
    Dim i As Integer, k As Integer
    Dim Result As String, AlgorithmT As String
    Dim ado As Object, iFile As Byte, blockSizex As Long
    
    Select Case Algorithmx
        Case 0: AlgorithmT = md5x
        Case 1: AlgorithmT = SHA1x
        Case 2: AlgorithmT = SHA256x
        Case 3: AlgorithmT = SHA384x
        Case 4: AlgorithmT = SHA512x
    End Select
    ReDim Hash(15)
    If IsFile = True Then '选择文件或者字符串
        If filez = 0 Then
            filez = fso.GetFile(strText).Size
            If filez = 0 Then GetHash = "UC": Exit Function
        End If
        ReDim Bytex(filez - 1)
        If cmCodex = False Then '是否包含非ansi字符
            iFile = FreeFile(0) '生成文件编号 'https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/freefile-function
            Open strText For Binary As iFile
            Get iFile, , Bytex
            Close iFile
        Else
            ReDim arrx(blockSize)
            Set ado = CreateObject("adodb.stream")
            With ado
                .Mode = 3 '不能选1,只读模式/,3 读写模式
                .type = 1  'adTypeBinary 'adTypeText=2
                .Open
                .LoadFromFile (strText)
                .Position = 0
                blockSizex = blockSize + 1 '读取128K
                Do Until .EOS = True '-----读取数据的末尾为止
                    arrx = .Read(blockSizex)
                    CombineArr arrx
                Loop
                .Close
            End With
            Ccx = 0
            Set ado = Nothing
        End If
    Else
        Set objUTF8 = CreateObject(UTF8x) '--https://docs.microsoft.com/en-us/dotnet/api/system.text.utf8encoding?view=netframework-4.8
        Bytex = objUTF8.GetBytes_4(strText) 'https://docs.microsoft.com/en-us/dotnet/api/system.text.utf8encoding.getbytes?view=netframework-4.8
        Set objUTF8 = Nothing
    End If
    '-------------------------------------处理方式
    Set objHash = CreateObject(AlgorithmT)
    objHash.Initialize
    Hash = objHash.ComputeHash_2(Bytex) 'ComputeHash_1(inputStream As stream),ComputeHash_2(buffer As byte()),ComputeHash_3(buffer As byte(), offset As int, count As int)
    objHash.Clear
    k = UBound(Hash) + 1
    For i = 1 To k
        Result = Result & Right$("0" & Hex(AscB(MidB(Hash, i, 1))), 2)
    Next i
    If rType = False Then '是否输出小写
        GetHash = LCase(Result)
    Else
        GetHash = Result
    End If
    Erase Hash
    Erase Bytes
    Set objHash = Nothing
End Function

Private Function CombineArr(xcx() As Byte) '将读取到的信息拼接在一起
    Dim i As Long
    
    i = UBound(xcx)                  '采用copymemory的方式将避免使用循环造成的严重时间浪费
    CopyMemory Bytex(Ccx), xcx(0), i + 1 '拼接到所在数组的(开始)位置,拼接数组的开始位置,拼接的长度
    Ccx = Ccx + i + 1
End Function
