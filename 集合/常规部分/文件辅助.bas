Attribute VB_Name = "�ļ�����"
Option Explicit
Option Compare Text
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'ʱ�� -����ѵ��
Private Declare Function timeGetTime Lib "winmm.dll" () As Long 'ʱ�� -����ѵ��
Private Declare Function CharUpper Lib "user32.dll" Alias "CharUpperA" (ByVal lpsz As String) As String                              '���ַ�����ͳһת��Ϊansi(��д)

Public errcodenx As String '��ansi�����ַ��������ַ����ĵ�λ��
Public errcodepx As String
Public Tagfnansi As Boolean, Tagfpansi As Boolean '����ĸ�λ�ó��ַ�ansi������ַ�
Private Declare Function IsTextUnicode Lib "Advapi32.dll" (ByVal intP As Long, ByVal sBuffer As Long) As Long



Sub utetst()
Dim s As String
s = Cells(20, 2).Value
Debug.Print IsTextUnicode(StrPtr(s), LenB(s))
'Dim arr() As Byte
'arr = StrConv(s, vbFromUnicode)
'Cells(21, 2) = CharUpper(s)
End Sub


Function ErrCode(ByVal strFilen As String, ByVal Exccode As Byte, Optional ByVal strFilep As String) As Integer '����ļ����Ƿ�����쳣�ַ�
    Dim strFile As String
    Dim i As Byte, n As Byte, p As Byte, m As Byte   'Exccode�������ֲ�ͬ��Դ������'strfilen,strfilep·��,��Ҫȷ������ansi�����λ�����ļ�������·��
    Dim arrn() As String, arrp() As String, k As Byte
    Dim strx As String, strx1 As String
    
    ErrCode = 1                                                   'If InStr(Mid(strFile, i, 1), CharUpper(Mid(strFile, i, 1))) = 0 Then
    Tagfnansi = False 'ʹ��ǰ���в�������
    Tagfpansi = False
    errcodenx = ""
    errcodepx = ""
    
    n = Len(strFilen)
    p = Len(strFilep)
    
    If p = 0 And n = 0 Then '�������������»�����޷���ȡ·���ļ�������
        ErrCode = -1
        Exit Function
    End If
    
    If n > 0 Then ReDim arrn(1 To n)
    If p > 0 Then ReDim arrp(1 To p)
    
    If Exccode = 1 Then ''�Ƚ��ĸ����ֵ��ַ����ȳ�,�ȼ���϶̵�,��ʡʱ��,��ΪҪ�ֱ��ȡ��ͬ״̬�µķ�ansi,1��ʱ��ֻ��Ҫ���ִ����쳣�ַ�,����Ҫ֪���ַ�������ʲôλ��,0��ʱ����Ҫ��ȡ�ļ�·����ͬ���ֵ��ַ������쳣�ַ���״��
        If p < n Then
            m = 1
        Else
            m = 2
        End If
    End If
    
    If m = 1 Or Exccode = 0 Then
100
        strFile = strFilen
        For i = 1 To n
            strx = Mid$(strFile, i, 1)
            strx1 = strx                  'ע�����ﲻ��ͬʱʹ��strx
            If InStr(strx, CharUpper(strx1)) = 0 Then 'mid/left/rgiht �����$��ʾ��������ݰ���string�������ͽ��д���
                ErrCode = ErrCode + 1
                If Exccode = 1 Then Exit Function '�������Լӿ�������ת�ٶ�,�����޷����쳣�ַ���λ��ȫ����ע����
                If Tagfnansi = False Then Tagfnansi = True
                arrn(i) = i '��ʱ�洢��ǳ��쳣�ַ���·���г��ֵ�λ��
            End If
        Next
        If m = 1 And Tagfpansi = False Then GoTo 101
    End If
    
    m = 0 '���ò���, ��ֹm=2,û���쳣�ַ���ʱ���γ���ѭ��
    
    If m = 2 Or Exccode = 0 Then
101
        strFile = strFilep
        For k = 1 To p
            strx = Mid$(strFile, k, 1)
            strx1 = strx
            If InStr(strx, CharUpper(strx1)) = 0 Then
                ErrCode = ErrCode + 1
                If Exccode = 1 Then Exit Function '�������Լӿ�������ת�ٶ�,�����޷����쳣�ַ���λ��ȫ����ע����
                If Tagfpansi = False Then Tagfpansi = True
                arrp(k) = k '��ʱ�洢��ǳ��쳣�ַ���·���г��ֵ�λ��
            End If
        Next
        If m = 2 And Tagfnansi = False Then GoTo 100 '�Ҳ�����������
    End If
    
    If ErrCode > 0 And Exccode = 0 Then
        If Tagfpansi = True Then errcodepx = Trim(Join(arrp, " ")) '���ݺϲ�����(ע��������Ҫ���ַ�������)
        If Tagfnansi = True Then errcodenx = Trim(Join(arrn, " "))
    End If
End Function

Function OpenBy(ByVal FilePath As String) As String '��ȡ�ļ�����Ĭ�Ϲ�������
    Dim str$, Result$
    
    str = LCase(Right(FilePath, Len(FilePath) - InStrRev(FilePath, ".") + 1)) '��.�ŵ��ļ���׺��
    With CreateObject("wscript.shell")
        On Error Resume Next
        Result = .RegRead("HKEY_CLASSES_ROOT\" & str & "\") '��ȡ��׺����Ӧ��ע������
        If Len(Result) > 0 Then
            Result = .RegRead("HKEY_CLASSES_ROOT\" & Result & "\shell\open\command\") '��ע�������ҵ��򿪵ĳ���·��
            If Result Like """*" Then Result = Split(Result, """")(1) Else Result = Split(Result, " ")(0)
            OpenBy = Result
        End If
    End With
End Function

Function CheckFileFrom(ByVal xpath As String, ByVal cmCode As Byte) As Boolean '����ļ��Ƿ���Դ��ϵͳ�ļ���
    Dim strx As String
    Dim strdsk As String, strdlw As String, strdcm As String, struserfile As String
    
    CheckFileFrom = False
    If cmCode = 1 Then     '��ʾ�ļ�
        strx = fso.GetFile(xpath).ParentFolder & "\"
    ElseIf cmCode = 2 Then
        strx = xpath & "\"     '��ʾ�ļ���
    End If
    
    If fso.GetDriveName(xpath) = Environ("SYSTEMDRIVE") Then '����ļ����ڴ��̺�ϵͳλ��ͬһ����
        struserfile = Environ("UserProfile") '�û��ļ���
        strdsk = struserfile & "\Desktop\"
        strdlw = struserfile & "\Downloads\"
        strdcm = struserfile & "\Documents\" '�����ļ�ֻ������ϵͳ�̵�����λ��-����-����-�ĵ�
        If InStr(strx, strdsk) = 0 And InStr(strx, strdlw) = 0 And InStr(strx, strdcm) = 0 Then CheckFileFrom = True
    End If
End Function
'������word��д������/������,openstream��������ִ���, �޷�ͨ���˷��������word�ļ��Ƿ�����������
'���ֵ�pdf���ܻ����openstream�޷���
Function FileStatus(ByVal filecodex As String, Optional ByVal cmCode As Byte) As Byte '�ж��ļ��Ƿ����/����excel/�Ƿ��ڴ򿪵�״̬     'https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/openastextstream-method
    Dim i As Byte, k As Byte
    Dim address As String, strx As String              '�����ж��Ƿ���Ҫ���ļ�
    Dim fl As File, flop As Object, filex As String
                                                        'Boolean-�ɸĳ�byte������Ե��ڿ��Ը���ͬ�Ľ׶γ��ֵ���������fileexist��ֵ,��1,2,3�ֱ����ͬ�ĺ���
    On Error GoTo 100
    Call SearchFile(filecodex) '�����ļ��Ƿ������Ŀ¼
    If Rng Is Nothing Then
        FileStatus = 1 '�ļ���������Ŀ¼
    Else
        If cmCode = 1 Then FileStatus = 2: Exit Function
        address = Rng.Offset(0, 3).Value '������Ŀ¼
        If fso.fileexists(address) = False Then
            FileStatus = 3
            Call DeleMenu(Rng.Row) '����ļ���������ִ�����Ŀ¼��Ϣ
        Else
            If cmCode = 2 Then FileStatus = 4: Exit Function '��ִ�к�����ж�-ֻ�ж��ļ��Ƿ�����ڱ����Լ�Ŀ¼
            filex = LCase(Rng.Offset(0, 2).Value)
            If filex Like "xl*" Then  '����ļ���������excel,��ô�жϴ򿪵��ļ��Ƿ�����/�����Ǳ��ļ�,��ΪExcel�޷����������ļ�
                k = Workbooks.Count
                strx = Rng.Offset(0, 1).Value
                For i = 1 To k
                    If strx = Workbooks(i).Name Then FileStatus = 5: Exit Function
                Next
            Else
                If Len(Rng.Offset(0, 35).Value) > 0 Then '�ļ���������
                    If WmiCheckFileOpen(address) = False Then
                        FileStatus = 0
                    Else
                        FileStatus = 6
                    End If
                    Set fl = Nothing
                    Set flop = Nothing
                    Exit Function
                End If
                If filex <> "txt" Then '�ж��ļ��Ƿ��ڴ򿪵�״̬,֧�ַ�ansi�ַ�·��
                    Set fl = fso.GetFile(address) '��ȡ�ļ�����
                    Set flop = fl.OpenAsTextStream(ForAppending, TristateUseDefault) 'ע�����ﲻҪѡforwriting����,����᳹�����ļ� ,ForAppending��ʾ�����һ�е�λ��׼��д����Ϣ
                    flop.Close
                End If
            End If
        End If
    End If
    Set fl = Nothing
    Set flop = Nothing
    Exit Function
100
    If Err.Number = 70 Then
       FileStatus = 7     '���ڴ򿪵�״̬
       If WmiCheckFileOpen(address) = False Then FileStatus = 0: Rng.Offset(0, 35) = 1 '����ļ��Ƿ������뱣��
    Else
        FileStatus = 8 '���������Ĵ���
    End If
    Err.Clear
    Set fl = Nothing
    Set flop = Nothing
End Function

Function FileTest(ByVal FilePath As String, ByVal filex As String, ByVal FileName As String) As Byte '���ж��ļ��Ƿ��ڴ򿪵�״̬/���ж��ļ��Ƿ�����ǲ��е�
    Dim fl As File, i As Byte, k As Byte, c As Byte
    Dim flop As Object
    Dim errx As Integer
    '-----------------------OpenAsTextStream���ַ����и��ϴ�ľ���,�Ǿ����޷��ж����뱣�����ļ��Ĵ�״̬,���������뱣����pdf�ļ�, txt�ļ�Ҳ�޷��ж�
    On Error GoTo 100
    FileTest = 0 '��ʼֵ
    If Len(FilePath) = 0 Then FileTest = 1: Exit Function '���ݹ�����ֵΪ�ջ����޷���Ч��ȡ����
  
    If fso.fileexists(FilePath) = False Then FileTest = 2: Exit Function
    
    filex = LCase(filex)
    If filex = "txt" Then FileTest = 3: Exit Function
    If filex Like "xl*" Then
        i = Workbooks.Count
        For k = 1 To i
            If FileName = Workbooks(k).Name Then FileTest = 4
        Next
        c = 1
    End If
    Set fl = fso.GetFile(FilePath)
    Set flop = fl.OpenAsTextStream(ForAppending, TristateUseDefault) 'ͨ���ж��ļ��Ŀɷ���״̬(�Ƿ�����)
    flop.Close
    Set fl = Nothing
    Set flop = Nothing
    Exit Function
100
    errx = Err.Number
    If errx = 70 Then
        FileTest = 5 '�ļ����ڴ򿪵�״̬
        If c <> 1 Then
            If WmiCheckFileOpen(FilePath) = False Then FileTest = 6 '�ļ��������뱣��
        End If
    Else
        FileTest = 7
    End If
    Err.Clear
    Set fl = Nothing
    Set flop = Nothing
End Function
