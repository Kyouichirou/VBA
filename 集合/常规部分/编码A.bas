Attribute VB_Name = "编码A"

'
'''' WinApi function that maps a UTF-16 (wide character) string to a new character string
'Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
'    ByVal CodePage As Long, _
'    ByVal dwFlags As Long, _
'    ByVal lpWideCharStr As Long, _
'    ByVal cchWideChar As Long, _
'    ByVal lpMultiByteStr As Long, _
'    ByVal cbMultiByte As Long, _
'    ByVal lpDefaultChar As Long, _
'    ByVal lpUsedDefaultChar As Long) As Long
'
'' CodePage constant for UTF-8
'Private Const CP_UTF8 = 65001

'''' Return byte array with VBA "Unicode" string encoded in UTF-8
'Public Function Utf8BytesFromString(strInput As String) As Byte()
'    Dim nBytes As Long
'    Dim abBuffer() As Byte
'    ' Catch empty or null input string
'    Utf8BytesFromString = vbNullString
'    If Len(strInput) < 1 Then Exit Function
'    ' Get length in bytes *including* terminating null
'    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, 0&, 0&, 0&, 0&)
'    ' We don't want the terminating null in our byte array, so ask for `nBytes-1` bytes
'    ReDim abBuffer(nBytes - 2)  ' NB ReDim with one less byte than you need
'    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(abBuffer(0)), nBytes - 1, 0&, 0&)
'    Utf8BytesFromString = abBuffer
'End Function
'
'Sub testAdodbStream()
'    Dim bg(1) As Byte, bu(1) As Byte, butf(2) As Byte
''    bg(0) = &HD6                                         'D6D0是"中"的ANSI/gb2312编码
''    bg(1) = &HD0
''    MsgBox "转换GB2312字节数组为BSTR字符串：" & BytesToBstr(bg, "GB2312")         '"中"
'
''    bu(0) = &H2D                                         '4E2D是"中"的Unicode编码,低位在前
''    bu(1) = &H4E
''    MsgBox "转换Unicode字节数组为BSTR字符串：" & BytesToBstr(bu, "Unicode")       '"中"
''
'    butf(0) = &HE4                                       'E4 B8 AD 是"中"的UTF-8编码
'    butf(1) = &HB8
'    butf(2) = &HAD
'    MsgBox "转换UTF-8字节数组为BSTR字符串：" & BytesToBstr(butf, "utf-8")       '"中"
''
''    Dim utfbody() As Byte, str As String, buni() As Byte
''    str = "中"
''    buni = str
''    utfbody = BytesToBytesNoBom(buni, "Unicode", "utf-8")     '转换unicode字节数组为utf-8的字节数组
''    utfbody = BytesToBytesNoBom(bg, "gb2312", "utf-8")        '转换gb2312字节数组为utf-8的字节数组
''
''    Dim strBase64
''    strBase64 = Base64Encode(bg, "gb2312")                                               '转换gb2312字节数组为Base64字符串
''    MsgBox "解码Base64字符串为gb2312字符串:" & Base64Decode(strBase64, "gb2312")         '解码Base64字符串为gb2312字符串
''
''    strBase64 = Base64Encode(utfbody, "utf-8")                                           '转换uft-8字节数组为Base64字符串
''    MsgBox "解码Base64字符串为utf-8字符串:" & Base64Decode(strBase64, "utf-8")           '解码Base64字符串为utf-8字符串
'End Sub
'
'
'Function BytesToBstr(ByRef arrBody() As Byte, ByVal CodeBase As String) As String '- ------------------------------------------- -
''  函数说明：字节数组转换成Unicode字符串（BSTR）
''- ------------------------------------------- -
'    Dim objStream As Object
'
'    Set objStream = CreateObject("ADODB.Stream")
'    objStream.type = 1     'adTypeBinary=1  adTypeText=2
'    objStream.Mode = 3     'adModeRead=1  adModeWrite=2  adModeReadWrit=3  adModeUnknown=0
'    objStream.Open
'    objStream.Write arrBody
'    objStream.Position = 0
'    objStream.type = 2      'adTypeBinary=1  adTypeText=2
'    objStream.CharSet = CodeBase
'    BytesToBstr = objStream.ReadText
'    objStream.Close
'    Set objStream = Nothing
'End Function
'
'Function BytesToBytesNoBom(ByRef arrBody() As Byte, ByVal SCodeBase As String, ByVal DCodeBase As String) As Byte()
''- ------------------------------------------- -
''  函数说明：不同编码的字节数组转换
''- ------------------------------------------- -
'    Dim objStream As Object
'    Dim SText$, Dtext$
'
'    Set objStream = CreateObject("ADODB.Stream")
'    objStream.type = 1     'adTypeBinary=1  adTypeText=2
'    objStream.Mode = 3     'adModeRead=1  adModeWrite=2  adModeReadWrit=3  adModeUnknown=0
'    objStream.Open
'    objStream.Write arrBody
'    objStream.Position = 0
'
'    objStream.type = 2                  'adTypeText = 2
'    objStream.CharSet = SCodeBase
'    SText = objStream.ReadText          '读取文本到sCode(Unicode)
'
'    objStream.Position = 0              ' 这只是定位到文件头
'    objStream.SetEOS                    'Position=0,更新 EOS 属性的值。使EOS的位置为0（也就是把结尾设成开头的位置）
'    objStream.type = 2                  'adTypeText = 2
'    objStream.CharSet = DCodeBase       '指定输出编码
'    objStream.WriteText SText           '写入文本数据到Adodb.Stream
'    'objStream.SaveToFile ThisWorkbook.Path & Application.PathSeparator & "out.bin", 2    '输出成文件
'
'    objStream.Position = 0              '切换type之前，要先重置指针为0
'    objStream.type = 1                  'adTypeBinary=1  adTypeText=2
'    If InStr(1, DCodeBase, "utf-8", vbTextCompare) > 0 Then                               '去掉BOM
'        objStream.Position = 3
'    ElseIf InStr(1, DCodeBase, "unicode", vbTextCompare) > 0 Then
'        objStream.Position = 2
'    End If
'    BytesToBytesNoBom = objStream.Read
'    objStream.Close
'    Set objStream = Nothing
'End Function
'
'
'Function Base64Encode(varIn As Variant, CodeBase As String) As String
''- ------------------------------------------- -
''  函数说明：BASE64编码
''- ------------------------------------------- -
'    Dim adoStream As Object
'    Dim xmlDoc As Object
'    Dim xmlNode As Object
'
'    Set adoStream = CreateObject("ADODB.Stream")
'    adoStream.CharSet = CodeBase   '文本编码
'    If VarType(varIn) = vbString Then
'        adoStream.type = 2    'adTypeText
'        adoStream.Open
'        adoStream.WriteText varIn
'    ElseIf VarType(varIn) = vbByte Or vbArray Then
'        adoStream.type = 1    'adTypeBinary
'        adoStream.Open
'        adoStream.Write varIn
'    Else
'        Exit Function
'    End If
'    adoStream.Position = 0
'    adoStream.type = 1        'adTypeBinary
'
'    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
'    Set xmlNode = xmlDoc.createElement("MyNode")
'    xmlNode.DataType = "bin.base64"
'    xmlNode.nodeTypedValue = adoStream.Read
'    Base64Encode = xmlNode.Text
'    adoStream.Close
'End Function
'
'Public Function Base64Decode(varIn As Variant, CodeBase As String, Optional ByVal ReturnValueType As VbVarType = vbString) As Byte()
''- ------------------------------------------- -
''  函数说明：BASE64解码
''- ------------------------------------------- -
'    Dim adoStream As Object
'    Dim xmlDoc As Object
'    Dim xmlNode As Object
'
'    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
'    Set xmlNode = xmlDoc.createElement("MyNode")
'    xmlNode.DataType = "bin.base64"
'    If VarType(varIn) = vbString Then
'        xmlNode.Text = Replace(varIn, vbCrLf, "")
'    ElseIf VarType(varIn) = vbByte Or vbArray Then
'        xmlNode.Text = Replace(StrConv(varIn, vbUnicode), vbCrLf, "")
'    Else
'        Exit Function
'    End If
'
'    Set adoStream = CreateObject("ADODB.Stream")
'    adoStream.CharSet = CodeBase  '文本的编码
'    adoStream.type = 1            'adTypeBinary
'    adoStream.Open
'    adoStream.Write xmlNode.nodeTypedValue
'    adoStream.Position = 0
'    If ReturnValueType = vbString Then
'        adoStream.type = 2       'adTypeText
'        Base64Decode = adoStream.ReadText
'    Else
'        Base64Decode = adoStream.Read
'    End If
'    adoStream.Close
'End Function

