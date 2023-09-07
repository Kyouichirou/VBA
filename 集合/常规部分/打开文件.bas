Attribute VB_Name = "���ļ�"
Option Explicit
'---------------https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/declare-statement
Private Declare Function aShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function uShellExecute Lib "shell32.dll" Alias "ShellExecuteW" (ByVal hwnd As Long, _
ByVal lpOperation As String, ByVal lpFile As Long, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1
Public Conn As New ADODB.Connection
Public SQL As String
Public rs As ADODB.Recordset
'-----------------------------ado��������
Public Rng As Range  'ע��ʹ�õ�ģ��Ĳ㼶�������ģ��͵��ó���������ͬ��ģ��
'--------------------����Ŀ¼
Public Recentfile As String '����Ķ�
Public Reditx As Byte '����Ķ�
'---------------------�������ݸ���
Public QRfilepath As String '��ά���ַ
Public QRtextCN As String, QRtextEN As String, Barcodex As String '��ά��,������
Public Turlx As String '�������ַ
'-----------------------��������/��ά��/������
Public Statisticsx As Byte '����ͳ�ƴ���
Public UF3Show As Byte 'uf3�����������
Public UF4Show As Byte 'uf4��¼����
Public OpenFilex As Byte

Function OpenFile(ByVal filecode As String, ByVal filenx As String, ByVal filex As String, ByVal FilePath As String, _
ByVal spcode As Byte, ByVal Erransi As String, Optional ByVal cmCode As Byte) As Boolean '���ļ� 'errcode���ڱ���ļ�·���Ƿ���ڿո������ 'spode 0 һ������,1��ʾ���Կ��ư�,2��ʾ�����Ҽ�
    Dim i As Integer, k As Byte
    Dim fl As File, exepath As String, filepathx As String, recenttime As String
    
    On Error GoTo 100
    OpenFile = True '��ִ�нϸ��ӵ�����ʱ����ִ�еĽ��
    If filex = "xlsx" Then 'Excel�޷���ͬ���ļ�
        If spcode = 1 Then           'spcode���ڱ�ʶ�򿪵���Դ�ǿ��ư廹��ֱ�����������������
'            UserForm3.Hide
'            UserForm3.Show 0             '��������
        ' ���ǵ�userform�ڴ򿪶��Excel�����ļ���Ϊ�鷳�Ľ���, �������ñ����������Ϊ���õĹ������
            Unload UserForm3  'ж�ص�����
            k = 1
            Call Rewds                    '���򿪵��ļ�Ϊexcel���ʱ��,���ô���
        End If
        Workbooks.Open (FilePath) '������Ҫע�ⴰ������,�����жϵ�����
    Else
        If Erransi = "ERC" Then   '�ж��ļ���·���Ƿ�����������ַ�
            exepath = OpenBy(FilePath) '��ȡ�����ļ����͵ĳ���     '������Ըĳ�cmd����/powershell�ķ�ʽ���ļ�
            If LenB(exepath) = 0 Then
                MsgBox "���ļ����Ͳ����ڹ�������"
                Exit Function
            Else
                filepathx = """" & FilePath & """"             '���Է�ֹ�ļ�·�����ڿո������/shell����û�����������ַ�������
                exepath = """" & exepath & """"
                Shell exepath & " " & filepathx, vbNormalFocus 'ע�����·������Ҫ�пո�,���������ҵ���������� '��һ�ִ򿪷�ʽ
            End If
        Else
            ShellExecute 0, "open", FilePath, 0, 0, SW_SHOWNORMAL        '����api�ķ�ʽ���ļ�(����Ҫ�����ļ����ո������)
        End If
    End If
    If cmCode = 1 Then OpenFile = True: Exit Function '��ʾ������Դ�ڿ��ư�-�Ƚ�
'    Call LockWorkSheet 'ȷ�������Դ��ڱ༭״̬
'    ----------------------------------------ѡ��򿪵ĳ���
    Rng.Offset(0, 11) = Now                    '���ļ���ʱ��
    Rng.Offset(0, 12) = Rng.Offset(0, 12) + 1 '������ļ��Ĵ���
    
    Set fl = fso.GetFile(FilePath)
    With fl
        If .DateLastModified <> Rng.Offset(0, 6).Value Then '����ļ��޸�ʱ�䷢���仯,�ļ������ݷ����˱仯
            Reditx = 1 '����һ��ִ��ֵ���ڸ��´��������
            Rng.Offset(0, 6) = .DateLastModified '�ļ����޸�ʱ�䷢���ı�,���ܴ�������ļ������ݷ����仯,��md5���ܷ����仯,��С�����仯
            Rng.Offset(0, 5) = .Size
            If .Size < 1048576 Then
                Rng.Offset(0, 7) = Format(.Size / 1024, "0.00") & "KB"
            Else
                Rng.Offset(0, 7) = Format(.Size / 1048576, "0.00") & "MB"
            End If
            filex = UCase(filex)
            If filex Like "EPUB" Or filex Like "MOBI" Or filex Like "PDF" Then Rng.Offset(0, 9) = GetFileHashMD5(FilePath)
        End If
        Set fl = Nothing
    End With
    '---------------����ļ��������Ƿ����仯
    With ThisWorkbook.Sheets("������")         '�������¼����Ϣ
'        If .Range("w27").NumberFormatLocal <> "yyyy/m/d h:mm;@" Then Call DataSwitch '������ʾ�ĸ�ʽ
        If Len(.Range("u27").Value) > 0 Then
            If .Range("u27").Value = filecode Then .Range("w27") = Now: GoTo 1000 '���Ŀ¼�Ѵ�����д��
        End If
        If Len(.Range("u27").Value) = 0 Then             'ȷ����ӽ�����ֵ����һֱ���ڵ�һ��
            .Range("p27") = Rng.Offset(0, 1)
            .Range("u27") = filecode
            .Range("w27") = Now
        Else
            For i = 33 To 28 Step -1
                .Range("p" & i) = .Range("p" & i - 1)
                .Range("u" & i) = .Range("u" & i - 1)
                .Range("w" & i) = .Range("w" & i - 1)
            Next
            .Range("p27") = Rng.Offset(0, 1)
            .Range("u27") = filecode
            recenttime = Now              'ȷ�����ʹ��ڵ�ʱ����ȫһ��'������ֱ��ʹ��now
            .Range("w27") = recenttime
            If spcode = 1 Then Recentfile = recenttime '���ڼ�¼������ļ���ʱ��,�ж��Ƿ���Ҫ���´����е���Ϣ
        End If
    End With
    '----------------------------��¼��
1000
    RecordWrite filecode, Rng.Offset(0, 1).Value, Rng.Offset(0, 13).Value, Rng.Offset(0, 21).Value '���ļ���¼
    If spcode = 2 Or k = 1 Then Set Rng = Nothing '���ư���Ҫ��������rng
    If UF3Show = 3 Then OpenFilex = 1 '-----------�����ô���ִ�����ݸ���
    Exit Function '������
100
    If Err.Number <> 0 Then
        Err.Clear
        Set Rng = Nothing
        OpenFile = False
    End If
End Function

Function SearchFile(ByVal filecode As String) '�����ļ�Ŀ¼-ȫ�ֱ���
    With ThisWorkbook.Sheets("���")
        If Len(filecode) = 0 Then Exit Function: Set Rng = Nothing
        Set Rng = .Range("b6:b" & .[b65536].End(xlUp).Row).Find(filecode, lookat:=xlWhole) '��ȷ����
    End With
End Function

Private Function RecordWrite(ByVal Unicode As String, ByVal filen As String, ByVal mfilen As String, ByVal idcode As String) '��¼�򿪵Ĳ���
    Dim timea As String 'time As Date
    Dim FilePath As String
    
    If RecData = True Then
'        time = Format(Now, "yyyy��mm��dd�� hh:mm:ss")
        timea = Format(Recentfile, "ddd")
        SQL = "Insert into [�򿪼�¼$] (ͳһ����,�ļ���,���ļ���,��ʶ����,ʱ��,����) Values ('" & Unicode & "', '" & filen & "', '" & mfilen & "', '" & idcode & "','" & Recentfile & "', '" & timea & "')" '#�������ڱ�ʾʱ��ı��� '#" & time & "#
        Conn.Execute (SQL)
    End If
End Function
