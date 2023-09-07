Attribute VB_Name = "����A"

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
''    bg(0) = &HD6                                         'D6D0��"��"��ANSI/gb2312����
''    bg(1) = &HD0
''    MsgBox "ת��GB2312�ֽ�����ΪBSTR�ַ�����" & BytesToBstr(bg, "GB2312")         '"��"
'
''    bu(0) = &H2D                                         '4E2D��"��"��Unicode����,��λ��ǰ
''    bu(1) = &H4E
''    MsgBox "ת��Unicode�ֽ�����ΪBSTR�ַ�����" & BytesToBstr(bu, "Unicode")       '"��"
''
'    butf(0) = &HE4                                       'E4 B8 AD ��"��"��UTF-8����
'    butf(1) = &HB8
'    butf(2) = &HAD
'    MsgBox "ת��UTF-8�ֽ�����ΪBSTR�ַ�����" & BytesToBstr(butf, "utf-8")       '"��"
''
''    Dim utfbody() As Byte, str As String, buni() As Byte
''    str = "��"
''    buni = str
''    utfbody = BytesToBytesNoBom(buni, "Unicode", "utf-8")     'ת��unicode�ֽ�����Ϊutf-8���ֽ�����
''    utfbody = BytesToBytesNoBom(bg, "gb2312", "utf-8")        'ת��gb2312�ֽ�����Ϊutf-8���ֽ�����
''
''    Dim strBase64
''    strBase64 = Base64Encode(bg, "gb2312")                                               'ת��gb2312�ֽ�����ΪBase64�ַ���
''    MsgBox "����Base64�ַ���Ϊgb2312�ַ���:" & Base64Decode(strBase64, "gb2312")         '����Base64�ַ���Ϊgb2312�ַ���
''
''    strBase64 = Base64Encode(utfbody, "utf-8")                                           'ת��uft-8�ֽ�����ΪBase64�ַ���
''    MsgBox "����Base64�ַ���Ϊutf-8�ַ���:" & Base64Decode(strBase64, "utf-8")           '����Base64�ַ���Ϊutf-8�ַ���
'End Sub
'
'
'Function BytesToBstr(ByRef arrBody() As Byte, ByVal CodeBase As String) As String '- ------------------------------------------- -
''  ����˵�����ֽ�����ת����Unicode�ַ�����BSTR��
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
''  ����˵������ͬ������ֽ�����ת��
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
'    SText = objStream.ReadText          '��ȡ�ı���sCode(Unicode)
'
'    objStream.Position = 0              ' ��ֻ�Ƕ�λ���ļ�ͷ
'    objStream.SetEOS                    'Position=0,���� EOS ���Ե�ֵ��ʹEOS��λ��Ϊ0��Ҳ���ǰѽ�β��ɿ�ͷ��λ�ã�
'    objStream.type = 2                  'adTypeText = 2
'    objStream.CharSet = DCodeBase       'ָ���������
'    objStream.WriteText SText           'д���ı����ݵ�Adodb.Stream
'    'objStream.SaveToFile ThisWorkbook.Path & Application.PathSeparator & "out.bin", 2    '������ļ�
'
'    objStream.Position = 0              '�л�type֮ǰ��Ҫ������ָ��Ϊ0
'    objStream.type = 1                  'adTypeBinary=1  adTypeText=2
'    If InStr(1, DCodeBase, "utf-8", vbTextCompare) > 0 Then                               'ȥ��BOM
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
''  ����˵����BASE64����
''- ------------------------------------------- -
'    Dim adoStream As Object
'    Dim xmlDoc As Object
'    Dim xmlNode As Object
'
'    Set adoStream = CreateObject("ADODB.Stream")
'    adoStream.CharSet = CodeBase   '�ı�����
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
''  ����˵����BASE64����
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
'    adoStream.CharSet = CodeBase  '�ı��ı���
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

