Attribute VB_Name = "编码"
Option Explicit
'https://www.zhihu.com/question/23374078
Private Const cstBase64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/" '码表 'https://www.cnblogs.com/feixiangmanon/p/10474462.html

Function Base64Encode(ByVal strSource As String) As String 'base64编码
   Dim arrBase64() As String
   Dim arrB() As Byte, bTmp(2) As Byte, bT As Byte
   Dim i As Long, j As Long
   On Error Resume Next
   If UBound(arrBase64) = -1 Then
       arrBase64 = Split(StrConv(cstBase64, vbUnicode), vbNullChar)
   End If
   arrB = StrConv(strSource, vbFromUnicode)
  
   j = UBound(arrB)
   For i = 0 To j Step 3
       Erase bTmp
       bTmp(0) = arrB(i + 0)
       bTmp(1) = arrB(i + 1)
       bTmp(2) = arrB(i + 2)
      
       bT = (bTmp(0) And 252) / 4
       Base64Encode = Base64Encode & arrBase64(bT)
      
       bT = (bTmp(0) And 3) * 16
       bT = bT + bTmp(1) \ 16
       Base64Encode = Base64Encode & arrBase64(bT)
      
       bT = (bTmp(1) And 15) * 4
       bT = bT + bTmp(2) \ 64
       If i + 1 <= j Then
           Base64Encode = Base64Encode & arrBase64(bT)
       Else
           Base64Encode = Base64Encode & "="
       End If
    
       bT = bTmp(2) And 63
       If i + 2 <= j Then
           Base64Encode = Base64Encode & arrBase64(bT)
       Else
           Base64Encode = Base64Encode & "="
       End If
   Next
End Function

Function Base64Decode(strEncoded As String) As String 'base64解码
   On Error Resume Next
   Dim arrB() As Byte, bTmp(3)  As Byte, bT As Long, bRet() As Byte
   Dim i As Long, j As Long
   arrB = StrConv(strEncoded, vbFromUnicode)
   j = InStr(strEncoded & "=", "=") - 2
   ReDim bRet(j - j \ 4 - 1)
   For i = 0 To j Step 4
       Erase bTmp
       bTmp(0) = (InStr(cstBase64, Chr(arrB(i))) - 1) And 63
       bTmp(1) = (InStr(cstBase64, Chr(arrB(i + 1))) - 1) And 63
       bTmp(2) = (InStr(cstBase64, Chr(arrB(i + 2))) - 1) And 63
       bTmp(3) = (InStr(cstBase64, Chr(arrB(i + 3))) - 1) And 63
       bT = bTmp(0) * 2 ^ 18 + bTmp(1) * 2 ^ 12 + bTmp(2) * 2 ^ 6 + bTmp(3)
       bRet((i \ 4) * 3) = bT \ 65536
       bRet((i \ 4) * 3 + 1) = (bT And 65280) \ 256
       bRet((i \ 4) * 3 + 2) = bT And 255
   Next
   Base64Decode = StrConv(bRet, vbUnicode)
End Function

Function ToUnicode(ByVal str As String) As String 'Unicode编码 '带符号的内容会出现问题
    Dim code As String
    Dim obj As Object
    Set obj = CreateObject("MSScriptControl.ScriptControl")
    obj.AllowUI = True
    obj.Language = "JavaScript"
    
    code = code & "function ToUnicode(str)"
    code = code & "{"
    code = code & "return escape(str).replace(/%/g," & Chr(34) & "\\" & Chr(34) & ").toLowerCase();"
    code = code & "}"
    code = code & "ToUnicode (" & Chr(34) & str & Chr(34) & ")" '注意这里的单双引号的使用， 当原文本同时带有单双引号时， 需要加反斜杠\作为转义
    
    ToUnicode = obj.eval(code) '输出结果
    Set obj = Nothing
End Function



Function UnUnicode(ByVal str As String) As String 'Unicode解码
    Dim code As String
    Dim obj As Object
    Set obj = CreateObject("MSScriptControl.ScriptControl")
    obj.AllowUI = True
    obj.Language = "JavaScript"
    code = code & "function UnUnicode(str)"
    code = code & "{"
    code = code & "return unescape(str.replace(/\\/g, " & Chr(34) & "%" & Chr(34) & "));"
    code = code & "}"
    code = code & "UnUnicode (" & Chr(34) & str & Chr(34) & ")"
    UnUnicode = obj.eval(code) '输出结果
    Set obj = Nothing
End Function

Function UTF8_URLEncoding(szInput) As String 'UTF-8 URL编码
    Dim wch, uch, szRet
    Dim x
    Dim nAsc, nAsc2, nAsc3
    If szInput = "" Then
        UTF8_URLEncoding = szInput
        Exit Function
    End If
    For x = 1 To Len(szInput)
        wch = Mid(szInput, x, 1)
        nAsc = AscW(wch)
        If nAsc < 0 Then nAsc = nAsc + 65536
        If (nAsc And &HFF80) = 0 Then
            szRet = szRet & wch
        Else
            If (nAsc And &HF000) = 0 Then
                uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            Else
                uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
                Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
                Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            End If
        End If
    Next
    UTF8_URLEncoding = szRet
End Function

Function UTF8_UrlDecode(ByVal url As String) 'UTF-8 URL解码
    Dim b, ub   ''中文字的Unicode码(2字节)
    Dim UtfB    ''Utf-8单个字节
    Dim UtfB1, UtfB2, UtfB3 ''Utf-8码的三个字节
    Dim i, n, s
    n = 0
    ub = 0
    For i = 1 To Len(url)
        b = Mid(url, i, 1)
        Select Case b
        Case "+"
            s = s & " "
        Case "%"
            ub = Mid(url, i + 1, 2)
            UtfB = CInt("&H" & ub)
            If UtfB < 128 Then
                i = i + 2
                s = s & ChrW(UtfB)
            Else
                UtfB1 = (UtfB And &HF) * &H1000   ''取第1个Utf-8字节的二进制后4位
                UtfB2 = (CInt("&H" & Mid(url, i + 4, 2)) And &H3F) * &H40      ''取第2个Utf-8字节的二进制后6位
                UtfB3 = CInt("&H" & Mid(url, i + 7, 2)) And &H3F      ''取第3个Utf-8字节的二进制后6位
                s = s & ChrW(UtfB1 Or UtfB2 Or UtfB3)
                i = i + 8
            End If
        Case Else    ''Ascii码
            s = s & b
        End Select
    Next
    UTF8_UrlDecode = s
End Function

Function UrlEncode(ByVal strURL As String) As String 'GBK URL编码
    Dim i As Long
    Dim tempStr As String
    For i = 1 To Len(strURL)
        If asc(Mid(strURL, i, 1)) < 0 Then
            tempStr = "%" & Right(CStr(Hex(asc(Mid(strURL, i, 1)))), 2)
            tempStr = "%" & Left(CStr(Hex(asc(Mid(strURL, i, 1)))), Len(CStr(Hex(asc(Mid(strURL, i, 1))))) - 2) & tempStr
            UrlEncode = UrlEncode & tempStr
        ElseIf (asc(Mid(strURL, i, 1)) >= 65 And asc(Mid(strURL, i, 1)) <= 90) Or (asc(Mid(strURL, i, 1)) >= 97 And asc(Mid(strURL, i, 1)) <= 122) Then
            UrlEncode = UrlEncode & Mid(strURL, i, 1)
        Else
            UrlEncode = UrlEncode & "%" & Hex(asc(Mid(strURL, i, 1)))
        End If
    Next
End Function

Function URLDecode(ByVal strURL As String) As String 'GBK URL解码
    Dim i As Long
    If InStr(strURL, "%") = 0 Then URLDecode = strURL: Exit Function
    For i = 1 To Len(strURL)
        If Mid(strURL, i, 1) = "%" Then
            If Val("&H" & Mid(strURL, i + 1, 2)) > 127 Then
                URLDecode = URLDecode & Chr(Val("&H" & Mid(strURL, i + 1, 2) & Mid(strURL, i + 4, 2)))
                i = i + 5
            Else
                URLDecode = URLDecode & Chr(Val("&H" & Mid(strURL, i + 1, 2)))
                i = i + 2
            End If
        Else
            URLDecode = URLDecode & Mid(strURL, i, 1)
        End If
    Next
End Function


