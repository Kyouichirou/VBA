Attribute VB_Name = "杂项"
'vbCr    Chr(13) 回车符。
'vbLf    Chr(10) 换行符。
'vbCrLf  Chr(13) & Chr(10)   回车符和换行符。
'vbNewLine   Chr(13) & Chr(10)或 Chr(10) 平台指定的新行字符，适用于任何平台。
'vbNullChar  Chr(0)  ASCII码为0的字符。
'vbNullString  值为0的字符串，但和""不同。
'vbTab   Chr(9)  水平附签。
'-------------------------------常用常量

'https://docs.microsoft.com/zh-cn/dotnet/visual-basic/language-reference/statements/declare-statement
'
'[ <attributelist> ] [ accessmodifier ] [ Shadows ] [ Overloads ] _
'Declare [ charsetmodifier ] [ Sub ] name Lib "libname" _
'[ Alias "aliasname" ] [ ([ parameterlist ]) ]
'' -or-
'[ <attributelist> ] [ accessmodifier ] [ Shadows ] [ Overloads ] _
'Declare [ charsetmodifier ] [ Function ] name Lib "libname" _
'[ Alias "aliasname" ] [ ([ parameterlist ]) ] [ As returntype ]
'用于声明api接口
'
'attributelist   可选。 请参阅特性列表。
'accessmodifier  可选。 可以是以下其中一个值：
'
'- 公布
'- 避免
'- 友好
'- 专有
'- 受保护的朋友
'- 私有受保护
'
'请参阅Visual Basic 中的访问级别。
'Shadows 可选。 请参阅阴影。
'charsetmodifier 可选。 指定字符集和文件搜索信息。 可以是以下其中一个值：
'
'- Ansi （默认值）
'- Unicode
'- 自动
'Sub 可选，但 Sub Function 必须出现或。 指示外部过程不返回值。
'Function    可选，但 Sub Function 必须出现或。 指示外部过程返回值。
'name    必需。 此外部引用的名称。 有关详细信息，请参阅已声明的元素名称。
'Lib 必需。 引入一个 Lib 子句，该子句标识包含外部过程的外部文件（DLL 或代码资源）。
'libname 必需。 包含已声明过程的文件的名称。
'Alias   可选。 指示无法在其文件中按中指定的名称标识所声明的过程 name 。 在中指定其标识 aliasname 。
'aliasname   如果使用关键字，则为必需 Alias 。 通过以下两种方式之一标识过程的字符串：
'
'过程在其文件中的入口点名称，在引号（ "" ）中
'
'- 或 -
'
'数字符号（ # ）后跟一个整数，该整数指定过程入口点在其文件中的序号
'parameterlist   如果过程使用参数，则为必需。 请参阅参数列表。
'returntype  如果 Function 指定了并为，则 Option Strict 为必需 On 。 过程返回的值的数据类型。
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function CharUpper Lib "user32.dll" Alias "CharUpperA" (ByVal lpsz As String) As String
Private Declare Function CharLower Lib "user32.dll" Alias "CharLowerA" (ByVal lpsz As String) As String


Private Sub dkkdfdd00()
Dim csj As New cString
Dim i As Long
Dim d As String, k As Long
d = "hello, my world,则 Option Strict 为必需 On"
k = Len(d)
csj.BuffSize = 99999
csj.cString_Initial
For i = 1 To 100000
    csj.Append d
Next
csj.Combine_String
Set csj = Nothing
End Sub

'Sub StopTimer() '计时器 /理论上较高精度
'With New Stopwatch
'    .Restart
'    fuckgirl
'    .Pause
'    Debug.Print Format(.Elapsed, "0.000000"); " seconds elapsed"
'End With
'End Sub

Sub Faster_String_Connect()
    Dim i As Long
    Dim s As String, d As String
    Dim k As Long, m As Long
    d = "hello, my world,则 Option Strict 为必需 On"
    m = Len(d)
    i = m * 100000
    s = Space(i)
    k = 1
    For i = 1 To 100000
        Mid$(s, k, m) = d: k = 1 + i * m
    Next
    s = ""
End Sub

Private Sub ooopp()
Dim cr As New cRegex

cr.oReg_Initial
cr.oReg_Text = "2017-06-12" & vbCrLf & "2019-06-12"
cr.oReg_Pattern = "(\d{4})-(\d{2})-(\d{2})"
cr.ReplaceText "$1"

Set cr = Nothing
End Sub


Sub dkkfk()
MsgBox "OK", vbCritical
End Sub

Private Function CheckRname(ByVal cText As String, ByVal iMode As Byte) As String '将文件名中的非法字符替换掉 'C:\Windows\System32\drivers\etc
'    Dim strTemp As String
    Dim Codex As Variant
    Dim rText As String
    Dim i As Byte, k As Byte
    Dim strA As String, strB As String, strTempA As String
    '------------------------------------------------------------------其他的涉及到文件命名的也可以调用这个模块
    k = Len(Trim$(cText))
    strTempA = StrConv(cText, vbFromUnicode)
    strTempA = StrConv(strTempA, vbUnicode)
    For i = 1 To k
        strA = Mid$(cText, i, 1)
        strB = Mid$(strTempA, i, 1)
        If InStr(1, strB, strA, vbBinaryCompare) = 0 Then Mid$(cText, i, 1) = ChrW$(39)
    Next
    If iMode = 1 Then
        Codex = Array(124, 60, 62, 34, 39, 42, 63, 47)
    Else
        Codex = Array(58, 124, 92, 60, 62, 34, 39, 42, 63, 47)
    End If
    k = UBound(Codex)
    For i = 0 To k
        cText = Replace$(cText, ChrW$(Codex(i)), vbNullString, 1, , vbBinaryCompare)
    Next
    CheckRname = cText
End Function



