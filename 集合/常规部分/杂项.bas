Attribute VB_Name = "����"
'vbCr    Chr(13) �س�����
'vbLf    Chr(10) ���з���
'vbCrLf  Chr(13) & Chr(10)   �س����ͻ��з���
'vbNewLine   Chr(13) & Chr(10)�� Chr(10) ƽָ̨���������ַ����������κ�ƽ̨��
'vbNullChar  Chr(0)  ASCII��Ϊ0���ַ���
'vbNullString  ֵΪ0���ַ���������""��ͬ��
'vbTab   Chr(9)  ˮƽ��ǩ��
'-------------------------------���ó���

'https://docs.microsoft.com/zh-cn/dotnet/visual-basic/language-reference/statements/declare-statement
'
'[ <attributelist> ] [ accessmodifier ] [ Shadows ] [ Overloads ] _
'Declare [ charsetmodifier ] [ Sub ] name Lib "libname" _
'[ Alias "aliasname" ] [ ([ parameterlist ]) ]
'' -or-
'[ <attributelist> ] [ accessmodifier ] [ Shadows ] [ Overloads ] _
'Declare [ charsetmodifier ] [ Function ] name Lib "libname" _
'[ Alias "aliasname" ] [ ([ parameterlist ]) ] [ As returntype ]
'��������api�ӿ�
'
'attributelist   ��ѡ�� ����������б�
'accessmodifier  ��ѡ�� ��������������һ��ֵ��
'
'- ����
'- ����
'- �Ѻ�
'- ר��
'- �ܱ���������
'- ˽���ܱ���
'
'�����Visual Basic �еķ��ʼ���
'Shadows ��ѡ�� �������Ӱ��
'charsetmodifier ��ѡ�� ָ���ַ������ļ�������Ϣ�� ��������������һ��ֵ��
'
'- Ansi ��Ĭ��ֵ��
'- Unicode
'- �Զ�
'Sub ��ѡ���� Sub Function ������ֻ� ָʾ�ⲿ���̲�����ֵ��
'Function    ��ѡ���� Sub Function ������ֻ� ָʾ�ⲿ���̷���ֵ��
'name    ���衣 ���ⲿ���õ����ơ� �й���ϸ��Ϣ���������������Ԫ�����ơ�
'Lib ���衣 ����һ�� Lib �Ӿ䣬���Ӿ��ʶ�����ⲿ���̵��ⲿ�ļ���DLL �������Դ����
'libname ���衣 �������������̵��ļ������ơ�
'Alias   ��ѡ�� ָʾ�޷������ļ��а���ָ�������Ʊ�ʶ�������Ĺ��� name �� ����ָ�����ʶ aliasname ��
'aliasname   ���ʹ�ùؼ��֣���Ϊ���� Alias �� ͨ���������ַ�ʽ֮һ��ʶ���̵��ַ�����
'
'���������ļ��е���ڵ����ƣ������ţ� "" ����
'
'- �� -
'
'���ַ��ţ� # �����һ��������������ָ��������ڵ������ļ��е����
'parameterlist   �������ʹ�ò�������Ϊ���衣 ����Ĳ����б�
'returntype  ��� Function ָ���˲�Ϊ���� Option Strict Ϊ���� On �� ���̷��ص�ֵ���������͡�
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function CharUpper Lib "user32.dll" Alias "CharUpperA" (ByVal lpsz As String) As String
Private Declare Function CharLower Lib "user32.dll" Alias "CharLowerA" (ByVal lpsz As String) As String


Private Sub dkkdfdd00()
Dim csj As New cString
Dim i As Long
Dim d As String, k As Long
d = "hello, my world,�� Option Strict Ϊ���� On"
k = Len(d)
csj.BuffSize = 99999
csj.cString_Initial
For i = 1 To 100000
    csj.Append d
Next
csj.Combine_String
Set csj = Nothing
End Sub

'Sub StopTimer() '��ʱ�� /�����Ͻϸ߾���
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
    d = "hello, my world,�� Option Strict Ϊ���� On"
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

Private Function CheckRname(ByVal cText As String, ByVal iMode As Byte) As String '���ļ����еķǷ��ַ��滻�� 'C:\Windows\System32\drivers\etc
'    Dim strTemp As String
    Dim Codex As Variant
    Dim rText As String
    Dim i As Byte, k As Byte
    Dim strA As String, strB As String, strTempA As String
    '------------------------------------------------------------------�������漰���ļ�������Ҳ���Ե������ģ��
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



