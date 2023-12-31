VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cJSON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'将json的花括号转化为vba的字典，将方括号转化为vba的集合
Option Explicit
Const INVALID_JSON As Long = 1
Const INVALID_OBJECT As Long = 2
Const INVALID_ARRAY As Long = 3
Const INVALID_BOOLEAN As Long = 4
Const INVALID_NULL As Long = 5
Const INVALID_KEY As Long = 6
Const INVALID_RPC_CALL As Long = 7
Private psErrors As String

Friend Function GetParserErrors() As String
    GetParserErrors = psErrors
End Function

Friend Function ClearParserErrors() As String
    psErrors = ""
End Function

Friend Function parse(ByVal str As String) As Object '   parse string and create JSON object
    Dim Index As Long
    Index = 1
    psErrors = ""
    On Error Resume Next
    Call skipChar(str, Index)
    Select Case Mid$(str, Index, 1)
    Case "{"
        Set parse = ParseObject(str, Index)
    Case "["
        Set parse = ParseArray(str, Index)
    Case Else
        psErrors = "Invalid JSON"
    End Select
End Function

Private Function ParseObject(ByRef str As String, ByRef Index As Long) As Object '   parse collection of key/value
    Dim sKey As String
    
    Set ParseObject = CreateObject("Scripting.Dictionary")
    Call skipChar(str, Index)    ' "{"
    If Mid(str, Index, 1) <> "{" Then
        psErrors = psErrors & "Invalid Object at position " & Index & " : " & Mid(str, Index) & vbCrLf
        Exit Function
    End If
    Index = Index + 1
    Do
        Call skipChar(str, Index)
        If "}" = Mid$(str, Index, 1) Then
            Index = Index + 1
            Exit Do
        ElseIf "," = Mid$(str, Index, 1) Then
            Index = Index + 1
            Call skipChar(str, Index)
        ElseIf Index > Len(str) Then
            psErrors = psErrors & "Missing '}': " & Right$(str, 20) & vbCrLf
            Exit Do
        End If
        ' ---------------------------add key/value pair
        sKey = parseKey(str, Index)
        On Error Resume Next
        ParseObject.Add sKey, parseValue(str, Index)
        If Err.Number <> 0 Then
            psErrors = psErrors & Err.Description & ": " & sKey & vbCrLf
            Exit Do
        End If
    Loop
End Function

Private Function ParseArray(ByRef str As String, ByRef Index As Long) As Collection
    Dim k As Long
    
    Set ParseArray = New Collection
    Call skipChar(str, Index)  ' "["
    If Mid(str, Index, 1) <> "[" Then
        psErrors = psErrors & "Invalid Array at position " & Index & " : " + Mid(str, Index, 20) & vbCrLf
        Exit Function
    End If
    Index = Index + 1
    k = Len(str)
    Do
        Call skipChar(str, Index)
        If "]" = Mid(str, Index, 1) Then
            Index = Index + 1
            Exit Do
        ElseIf "," = Mid$(str, Index, 1) Then
            Index = Index + 1
            Call skipChar(str, Index)
        ElseIf Index > k Then
            psErrors = psErrors & "Missing ']': " & Right$(str, 20) & vbCrLf
            Exit Do
        End If
        ' add value
        On Error Resume Next
        ParseArray.Add parseValue(str, Index)
        If Err.Number <> 0 Then
            psErrors = psErrors & Err.Description & ": " & Mid$(str, Index, 20) & vbCrLf
            Exit Do
        End If
    Loop
End Function

Private Function parseValue(ByRef str As String, ByRef Index As Long)  'parse string / number / object / array / true / false / null
    Call skipChar(str, Index)
    Select Case Mid(str, Index, 1)
    Case "{"
        Set parseValue = ParseObject(str, Index)
    Case "["
        Set parseValue = ParseArray(str, Index)
    Case """", "'"
        parseValue = ParseString(str, Index)
    Case "t", "f"
        parseValue = parseBoolean(str, Index)
    Case "n"
        parseValue = parseNull(str, Index)
    Case Else
        parseValue = parseNumber(str, Index)
    End Select
End Function

Private Function ParseString(ByRef str As String, ByRef Index As Long) As String '   parse string
    Dim quote As String
    Dim Char As String
    Dim code As String
    Dim cStb As New cStringBuilder
    Dim k As Long

    Call skipChar(str, Index)
    quote = Mid(str, Index, 1)
    Index = Index + 1
    k = Len(str)
    Do While Index > 0 And Index <= k
        Char = Mid$(str, Index, 1)
        Select Case (Char)
        Case "\"
            Index = Index + 1
            Char = Mid$(str, Index, 1)
            Select Case (Char)
            Case """", "\", "/", "'"
                cStb.Append Char
                Index = Index + 1
            Case "b"
                cStb.Append vbBack
                Index = Index + 1
            Case "f"
                cStb.Append vbFormFeed
                Index = Index + 1
            Case "n"
                cStb.Append vbLf
                Index = Index + 1
            Case "r"
                cStb.Append vbCr
                Index = Index + 1
            Case "t"
                cStb.Append vbTab
                Index = Index + 1
            Case "u"
                Index = Index + 1
                code = Mid$(str, Index, 4)
                cStb.Append ChrW(Val("&h" + code))
                Index = Index + 4
            End Select
        Case quote
            Index = Index + 1
            ParseString = cStb.ToString
            Set cStb = Nothing
            Exit Function
        Case Else
            cStb.Append Char
            Index = Index + 1
        End Select
    Loop
    ParseString = cStb.ToString
    Set cStb = Nothing
End Function

Private Function parseNumber(ByRef str As String, ByRef Index As Long) '   parse number
    Dim Value As String
    Dim Char As String

    Call skipChar(str, Index)
    Do While Index > 0 And Index <= Len(str)
        Char = Mid(str, Index, 1)
        If InStr("+-0123456789.eE", Char) Then
            Value = Value & Char
            Index = Index + 1
        Else
            If IsNumeric(Value) = True Then '防止数据不是数字
                parseNumber = CDec(Value)
            Else
                parseNumber = CDec(Index + 1) '这里还需要处理下
            End If
            Exit Function
        End If
    Loop
End Function

Private Function parseBoolean(ByRef str As String, ByRef Index As Long) As Boolean 'parse true / false
    Call skipChar(str, Index)
    If Mid$(str, Index, 4) = "true" Then
        parseBoolean = True
        Index = Index + 4
    ElseIf Mid$(str, Index, 5) = "false" Then
        parseBoolean = False
        Index = Index + 5
    Else
        psErrors = psErrors & "Invalid Boolean at position " & Index & " : " & Mid$(str, Index) & vbCrLf
    End If
End Function

Private Function parseNull(ByRef str As String, ByRef Index As Long) 'parse null
    Call skipChar(str, Index)
    If Mid$(str, Index, 4) = "null" Then
        parseNull = Null
        Index = Index + 4
    Else
        psErrors = psErrors & "Invalid null value at position " & Index & " : " & Mid(str, Index) & vbCrLf
    End If
End Function

Private Function parseKey(ByRef str As String, ByRef Index As Long) As String
    Dim dquote As Boolean
    Dim squote As Boolean
    Dim Char As String
    Dim k As Long
    
    Call skipChar(str, Index)
    k = Len(str)
    Do While Index > 0 And Index <= k
        Char = Mid$(str, Index, 1)
        Select Case (Char)
        Case """"
            dquote = Not dquote
            Index = Index + 1
            If Not dquote Then
                Call skipChar(str, Index)
                If Mid$(str, Index, 1) <> ":" Then
                    psErrors = psErrors & "Invalid Key at position " & Index & " : " & parseKey & vbCrLf
                    Exit Do
                End If
            End If
        Case "'"
            squote = Not squote
            Index = Index + 1
            If Not squote Then
                Call skipChar(str, Index)
                If Mid$(str, Index, 1) <> ":" Then
                    psErrors = psErrors & "Invalid Key at position " & Index & " : " & parseKey & vbCrLf
                    Exit Do
                End If
            End If
        Case ":"
            Index = Index + 1
            If Not dquote And Not squote Then
                Exit Do
            Else
                parseKey = parseKey & Char
            End If
        Case Else
            If InStr(vbCrLf & vbCr & vbLf & vbTab & " ", Char) Then
            Else
                parseKey = parseKey & Char
            End If
            Index = Index + 1
        End Select
    Loop
End Function

Private Sub skipChar(ByRef str As String, ByRef Index As Long) 'skip special character
    Dim bComment As Boolean
    Dim bStartComment As Boolean
    Dim bLongComment As Boolean
    Dim k As Long
    
    k = Len(str)
    Do While Index > 0 And Index <= k
        Select Case Mid$(str, Index, 1)
        Case vbCr, vbLf
            If Not bLongComment Then
                bStartComment = False
                bComment = False
            End If
        Case vbTab, " ", "(", ")"
        Case "/"
            If Not bLongComment Then
                If bStartComment Then
                    bStartComment = False
                    bComment = True
                Else
                    bStartComment = True
                    bComment = False
                    bLongComment = False
                End If
            Else
                If bStartComment Then
                    bLongComment = False
                    bStartComment = False
                    bComment = False
                End If
            End If
        Case "*"
            If bStartComment Then
                bStartComment = False
                bComment = True
                bLongComment = True
            Else
                bStartComment = True
            End If
        Case Else
            If Not bComment Then
                Exit Do
            End If
        End Select
        Index = Index + 1
    Loop
End Sub

Friend Function ToString(ByRef obj As Variant) As String
    Dim cStb As New cStringBuilder
    Dim bFI As Boolean
    Dim i As Long
    
    Select Case VarType(obj)
    Case vbNull
        cStb.Append "null"
    Case vbDate
        cStb.Append """" & CStr(obj) & """"
    Case vbString
        cStb.Append """" & Encode(obj) & """"
    Case vbObject
        bFI = True
        If TypeName(obj) = "Dictionary" Then
            cStb.Append "{"
            Dim Keys
            Keys = obj.Keys
            For i = 0 To obj.Count - 1
                If bFI Then bFI = False Else cStb.Append ","
                Dim key
                key = Keys(i)
                cStb.Append """" & key & """:" & ToString(obj.item(key))
            Next i
            cStb.Append "}"
        ElseIf TypeName(obj) = "Collection" Then
            cStb.Append "["
            Dim Value
            For Each Value In obj
                If bFI Then bFI = False Else cStb.Append ","
                cStb.Append ToString(Value)
            Next Value
            cStb.Append "]"
        End If
    Case vbBoolean
        If obj Then cStb.Append "true" Else cStb.Append "false"
    Case vbVariant, vbArray, vbArray + vbVariant
        Dim sEB
        cStb.Append multiArray(obj, 1, "", sEB)
    Case Else
        cStb.Append Replace(obj, ",", ".")
    End Select
    ToString = cStb.ToString
    Set cStb = Nothing
End Function

Private Function Encode(ByVal str As String) As String
    Dim cStb As New cStringBuilder
    Dim i As Long
    Dim j As Long
    Dim aL1 As Variant
    Dim aL2 As Variant
    Dim c As String
    Dim p As Boolean
    Dim k As Long
    Dim a As Long
    aL1 = Array(&H22, &H5C, &H2F, &H8, &HC, &HA, &HD, &H9)
    aL2 = Array(&H22, &H5C, &H2F, &H62, &H66, &H6E, &H72, &H74)
    k = Len(str)
    For i = 1 To k
        p = True
        c = Mid$(str, i, 1)
        For j = 0 To 7
            If c = Chr(aL1(j)) Then
                cStb.Append "\" & Chr(aL2(j))
                p = False
                Exit For
            End If
        Next

        If p Then
            a = AscW(c)
            If a > 31 And a < 127 Then
                cStb.Append c
            ElseIf a > -1 Or a < 65535 Then
                cStb.Append "\u" & String(4 - Len(Hex(a)), "0") & Hex(a)
            End If
        End If
    Next
    Encode = cStb.ToString
    Set cStb = Nothing
End Function

Private Function multiArray(aBD, iBC, sPS, ByRef sPT)   ' Array BoDy, Integer BaseCount, String PoSition
    Dim iDU As Long
    Dim iDL As Long
    Dim i As Long, k As Long
    Dim cStb As New cStringBuilder
    Dim sPB1, sPB2  ' String PointBuffer1, String PointBuffer2
    
    On Error Resume Next
    iDL = LBound(aBD, iBC)
    iDU = UBound(aBD, iBC)
    If Err.Number = 9 Then
        sPB1 = sPT & sPS
        k = Len(sPB1)
        For i = 1 To k
            If i <> 1 Then sPB2 = sPB2 & ","
            sPB2 = sPB2 & Mid$(sPB1, i, 1)
        Next
        '        multiArray = multiArray & toString(Eval("aBD(" & sPB2 & ")"))
        cStb.Append ToString(aBD(sPB2))
    Else
        sPT = sPT & sPS
        cStb.Append "["
        For i = iDL To iDU
            cStb.Append multiArray(aBD, iBC + 1, i, sPT)
            If i < iDU Then cStb.Append ","
        Next
        cStb.Append "]"
        sPT = Left$(sPT, iBC - 2)
    End If
    Err.Clear
    multiArray = cStb.ToString
    Set cStb = Nothing
End Function

Friend Function StringToJSON(ByVal st As String) As String ' Miscellaneous JSON functions
    Dim sFlds As String
    Dim sRecs As New cStringBuilder
    Dim lRecCnt As Long
    Dim lFld As Long
    Dim fld As Variant
    Dim rows As Variant
    Dim i As Long, j As Long
    Dim m As Long
    Dim n As Long
    Const FIELD_SEP = "~"
    Const RECORD_SEP = "|"
    
    lRecCnt = 0
    If st = "" Then
        StringToJSON = "null"
    Else
        rows = Split$(st, RECORD_SEP)
        i = LBound(rows)
        j = UBound(rows)
        For lRecCnt = i To j
            sFlds = ""
            fld = Split$(rows(lRecCnt), FIELD_SEP)
            m = UBound(fld)
            n = LBound(fld)
            For lFld = n To m Step 2
                sFlds = (sFlds & IIf(sFlds <> "", ",", "") & """" & fld(lFld) & """:""" & ToUnicode(fld(lFld + 1) & "") & """")
            Next    'fld
            sRecs.Append IIf((Trim(sRecs.ToString) <> ""), "," & vbCrLf, "") & "{" & sFlds & "}"
        Next    'rec
        StringToJSON = ("( {""Records"": [" & vbCrLf & sRecs.ToString & vbCrLf & "], " & """RecordCount"":""" & lRecCnt & """ } )")
    End If
End Function

Friend Function RStoJSON(rs As Object) As String
    On Error GoTo errHandler
    Dim sFlds As String
    Dim sRecs As New cStringBuilder
    Dim lRecCnt As Long
    Dim fld As ADODB.Field 'https://www.w3school.com.cn/ado/ado_ref_field.asp

    lRecCnt = 0
    If rs.State = adStateClosed Then
        RStoJSON = "null"
    Else
        If rs.EOF Or rs.BOF Then
            RStoJSON = "null"
        Else
            Do While Not rs.EOF And Not rs.BOF
                lRecCnt = lRecCnt + 1
                sFlds = ""
                For Each fld In rs.Fields
                    sFlds = (sFlds & IIf(sFlds <> "", ",", "") & """" & fld.Name & """:""" & ToUnicode(fld.Value & "") & """")
                Next    'fld
                sRecs.Append IIf((Trim(sRecs.ToString) <> ""), "," & vbCrLf, "") & "{" & sFlds & "}"
                rs.MoveNext
            Loop
            RStoJSON = ("( {""Records"": [" & vbCrLf & sRecs.ToString & vbCrLf & "], " & """RecordCount"":""" & lRecCnt & """ } )")
        End If
    End If
End Function

Friend Function ToUnicode(ByVal str As String) As String
    Dim x As Long
    Dim uStr As New cStringBuilder
    Dim uChrCode As Integer
    Dim k As Long
    
    k = Len(str)
    For x = 1 To k
        uChrCode = asc$(Mid$(str, x, 1))
        Select Case uChrCode
        Case 8:    ' backspace
            uStr.Append "\b"
        Case 9:    ' tab
            uStr.Append "\t"
        Case 10:    ' line feed
            uStr.Append "\n"
        Case 12:    ' formfeed
            uStr.Append "\f"
        Case 13:    ' carriage return
            uStr.Append "\r"
        Case 34:    ' quote
            uStr.Append "\"""
        Case 39:    ' apostrophe
            uStr.Append "\'"
        Case 92:    ' backslash
            uStr.Append "\\"
        Case 123, 125:    ' "{" and "}"
            uStr.Append ("\u" & Right$("0000" & Hex(uChrCode), 4))
        Case Is < 32, Is > 127:    ' non-ascii characters
            uStr.Append ("\u" & Right$("0000" & Hex(uChrCode), 4))
        Case Else
            uStr.Append Chr$(uChrCode)
        End Select
    Next
    ToUnicode = uStr.ToString
End Function

Private Sub Class_Initialize()
    psErrors = ""
End Sub




