VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm21 
   Caption         =   "FTP����"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3975
   OleObjectBlob   =   "UserForm21.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "UserForm21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents cf As cFTP    '��Ҫ��Ӧ�¼�
Attribute cf.VB_VarHelpID = -1
Dim strHost As String           '�������ƣ�IP��ַ��ʽ
Dim strUser As String           '�û���
Dim strPassword As String       '�û�����
Dim strFlag As String           '����״̬��������
Dim FilePath As String          '��userform3���ļ�·����ȡ����

'cFTP��ģ�����ػ��ϴ��ļ����¼�������ʹ������¼����ɽ���״̬��
Private Sub cf_FileTransferProgress(lCurrentBytes As Long, lTotalBytes As Long)
    If strFlag = "Downloading" Then     '������ó�Downloading����ʾ���ص����ļ�
        txtStatus.Text = Format(lCurrentBytes / lTotalBytes * 100, "0.00") & "% Done..."
    Else
        '����ļ��ϴ����ڽ��̱���֮ǰ��Ӹ�����Ϣ
        txtStatus.Text = strFlag & " | " & Format(lCurrentBytes / lTotalBytes * 100, "0.00") & "% Done..."
    End If
End Sub

Private Sub cmdConnect_Click() '����
    Dim strContent As String
    Dim lReturn As Long
    Dim lPort As Long
    Dim i As Integer
    Dim arrCont() As String
    
    If cmdConnect.Caption = "����" Then
        '������������
        If Trim(txtHost.Text) = "" Then
            MsgBox "���������������ƣ�", vbInformation, "��ʾ"
            Exit Sub
        Else
            strHost = Trim(txtHost.Text)
        End If
        constatus = 0
        If IsPing(strHost) = False Then MsgBox "�豸δ����": Exit Sub '����豸�Ƿ�������״̬ '����ģ����Լ�ftp�Ƿ����ӷǳ���(��Ҫ�ȴ�10+s)
        lPort = Trim(txtPort.Text)
        If CheckPort(strHost, lPort) = False Then MsgBox "�豸FTPδ��": Exit Sub '����豸��FTP�Ĵ�״̬
        '��Ϊ��ģ���е��Դ��ļ������״̬���豸�Ͽ���ʱ����ַǳ���ʱ��Ŀ���,������Ҫ����Ĺ�����ʵ�ּ��
        '����û���Ϊ�գ����ʾ����
        If Trim(txtUser.Text) = "" And Trim(txtPassword.Text) = "" Then
            strUser = "anonymous"
            strPassword = ""
        Else
            strUser = Trim(txtUser.Text)
            strPassword = Trim(txtPassword.Text)
        End If
        Set cf = New cFTP
        If chkPassive.Value = True Then
            cf.SetModePassive
        Else
            cf.SetModeActive
        End If
        
        '�������ӣ�����False��ʾ����ʧ��
        If lPort = "" Then
            lReturn = cf.OpenConnection(strHost, strUser, strPassword)
        Else
            lReturn = cf.OpenConnection(strHost, strUser, strPassword, lPort)
        End If
        
        If lReturn = False Then
            GoTo ErrHandle
        Else
            '��ȡ��ǰĿ¼���ļ��б������뵽�б���У�������ǰĿ¼�����ı�����
            '�ļ�����β�����"/"�ַ����ļ�����β������ļ���С
            strContent = cf.GetFTPDirectory
            txtPath.Text = strContent
            strContent = cf.GetFTPDirectoryContent
            '��������Ϊ��False��ʱ���ʾû���ļ�
            If strContent = "False" Then
                lstFile.Clear
            Else
                arrCont = Split(strContent, vbCrLf)
                lstFile.Clear
                For i = 0 To UBound(arrCont)
                    If Trim(arrCont(i)) <> "" Then
                        lstFile.AddItem arrCont(i)
                    End If
                Next i
            End If
            cmdConnect.Caption = "�Ͽ�����"
            txtStatus.Text = "�û�" & strUser & "�����ӵ�����" & strHost
        End If
    Else
        cf.CloseConnection
        txtPath.Text = ""
        lstFile.Clear
        cmdConnect.Caption = "����"
        txtStatus.Text = "�û��˳���δ����"
    End If
    Exit Sub
ErrHandle:
    MsgBox cf.GetLastErrorMessage
End Sub

'��FTP�ĵ�ǰĿ¼�´����µ�Ŀ¼
Private Sub cmdCreate_Click()
    Dim strPath As String
    If cmdConnect.Caption = "����" Then Exit Sub
    If Right(strPath, 1) = "/" Then strPath = Mid(strPath, 1, Len(strPath) - 1)
    strPath = Application.InputBox("��������Ҫ�������ļ������ƣ�", "�����ļ���", "Default")
    If Trim(strPath) = "" Then Exit Sub
    If cf.CreateFTPDirectory(strPath) = True Then
        lstFile.AddItem strPath & "/"
        txtStatus.Text = "����Ŀ¼" & strPath & "�ɹ�"
    Else
        MsgBox cf.GetLastErrorMessage
    End If
End Sub
'ɾ��ָ���ļ����ļ��У�����ļ����´��������ļ���������ɾ��
Private Sub cmdDelete_Click()
    Dim strPath As String
    If cmdConnect.Caption = "����" Then Exit Sub
    If lstFile.Text = "" Then Exit Sub
    strPath = lstFile.Text
    If MsgBox("��ȷ����Ҫɾ���ļ���Ŀ¼" & strPath & "��", vbYesNo) = vbNo Then Exit Sub
    If Right(strPath, 1) = "/" Then
        If cf.RemoveFTPDirectory(strPath) = True Then
            txtStatus.Text = "��ɾ��Ŀ¼" & strPath
            lstFile.RemoveItem lstFile.ListIndex
        Else
            MsgBox cf.GetLastErrorMessage
        End If
    Else
        strPath = Mid(strPath, 1, Len(strPath) - 1)
        strPath = Mid(strPath, 1, InStrRev(strPath, "(") - 1)
        If cf.DeleteFTPFile(strPath) = True Then
            txtStatus.Text = "��ɾ���ļ�" & strPath
            lstFile.RemoveItem lstFile.ListIndex
        Else
            MsgBox cf.GetLastErrorMessage
        End If
    End If
End Sub
'�����ļ�
Private Sub cmdDownload_Click()
    Dim strFile As String
    Dim strSource As String
    Dim fd As FileDialog, i As Byte
    'ʹ��FileDialog�����ȡ�ļ��������������������ļ�
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    strSource = ""
    With fd
        fd.AllowMultiSelect = False
        fd.Filters.Clear
        If .Show = -1 Then strSource = .SelectedItems(1)
    End With
    Set fd = Nothing
    If strSource = "" Then Exit Sub
    If cmdConnect.Caption = "����" Then Exit Sub
    strFlag = "Downloading"
    strFile = lstFile.Text
    If Trim(strFile) = "" Or Right(strFile, 1) = "/" Then Exit Sub
    strFile = Mid(strFile, 1, Len(strFile) - 1)
    strFile = Mid(strFile, 1, InStrRev(strFile, "(") - 1)
    '���Դ����¼�
    i = 2 '�Ѷ����Ƶķ�ʽ�����ļ�
    If LCase(Right$(strFile, Len(strFile) - InStrRev(strFile, "."))) Like "txt" Then i = 1 'ֻ���ı����ļ�����
    If cf.FTPDownloadFile(strSource & "\" & strFile, strFile, i) = True Then
        txtStatus.Text = "�����ļ�" & strFile & "�ɹ�"
    Else
        MsgBox cf.GetLastErrorMessage
    End If
End Sub
'�ϴ��ļ������Զ�ѡ
Private Sub cmdUpload_Click()
    Dim arrFile As Variant
    Dim i As Integer, k As Byte, m As Byte, j As Byte
    Dim strBaseFile As String
    Dim strError As String
    Dim strLocalFile As String
    
    If cmdConnect.Caption = "����" Then Exit Sub
    strError = ""
    m = 1
    If Len(FilePath) = 0 Then
        arrFile = Application.GetOpenFilename("�����ļ�(*.*),*.*", , "ѡ���ļ��ϴ�", , True)
        If IsArray(arrFile) = False Then Exit Sub
        m = UBound(arrFile)
    End If
    For i = 1 To m
        strLocalFile = arrFile(i)
        If ErrCode(strLocalFile, 1) > 1 Then GoTo 100 '����ļ���·���Ƿ������ansi, ���ϴ��ļ�ģ�����open�ķ�ʽ,��Ҫ�޸�Ϊado.stream
        j = j + 1
        strBaseFile = Mid(strLocalFile, InStrRev(strLocalFile, "\") + 1)
        strFlag = i & "/" & UBound(arrFile) & "�ϴ�"
        k = 2 '�Ѷ����Ƶķ�ʽ�����ļ�
        If LCase(Right$(strLocalFile, Len(strLocalFile) - InStrRev(strLocalFile, "."))) Like "txt" Then k = 1 'ֻ���ı����ļ�����
        If cf.FTPUploadFile(strLocalFile, strBaseFile, k) = True Then
            lstFile.AddItem strBaseFile & "(" & Format(FileLen(strLocalFile) / 1024, "0.00") & "kB)" '�޸ĵ�ʱ��,ע��Ҳ��Ҫ��filelen���ֺ�����ͬʱ�����
        Else
            strError = strError & vbCrLf & cf.GetLastErrorMessage
        End If
100
    Next i
    If strError <> "" Then MsgBox strError: Exit Sub
    If j < m Then
        If j = 0 Then
            MsgBox "�ϴ��ļ�ʧ��", vbInformation, "Warning"
        Else
            MsgBox "ѡ��" & m & "���ļ�" & ";�ɹ��ϴ�" & "j" & "��"
        End If
        Exit Sub
    End If
    txtStatus.Text = "�ϴ��ļ����"
End Sub
'����ǰĿ¼�ĳ���һ���ļ��У�ͬʱ�����б���е�����
Private Sub cmdUpper_Click()
    Dim strTemp As String
    Dim strContent As String
    Dim arrCont() As String
    Dim i As Integer
    
    If cmdConnect.Caption = "����" Then Exit Sub
    If txtPath.Text = "/" Then Exit Sub
    strTemp = txtPath.Text
    strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
    strTemp = Mid(strTemp, 1, InStrRev(strTemp, "/"))
    txtPath.Text = strTemp
    If cf.SetFTPDirectory(txtPath.Text) = True Then
        strContent = cf.GetFTPDirectoryContent
        If strContent = "False" Then
            lstFile.Clear
        Else
            arrCont = Split(strContent, vbCrLf)
            lstFile.Clear
            For i = 0 To UBound(arrCont)
                lstFile.AddItem arrCont(i)
            Next i
        End If
    End If
End Sub

Private Sub CommandButton1_Click() '����ļ�
    With UserForm3
    If .Label74.Caption = "Y" Then MsgBox "��֧�ִ��ļ����ϴ�": Exit Sub
    FilePath = .Label25.Caption
    End With
End Sub

'˫���б�����ݿ��Դ򿪸��ļ��У��������б������
Private Sub lstFile_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim strContent As String
    Dim arrCont() As String
    Dim i As Integer
    Cancel = True
    If cmdConnect.Caption = "����" Then Exit Sub
    If Right(lstFile.Text, 1) = "/" Then
        strTemp = txtPath.Text
        txtPath.Text = strTemp & lstFile.Text
        If cf.SetFTPDirectory(txtPath.Text) = True Then
            strContent = cf.GetFTPDirectoryContent
            If strContent = "False" Then
                lstFile.Clear
            Else
                arrCont = Split(strContent, vbCrLf)
                lstFile.Clear
                For i = 0 To UBound(arrCont)
                    lstFile.AddItem arrCont(i)
                Next i
            End If
        Else
            txtPath.Text = strTemp
            GoTo errhandle1
        End If
    End If
    Exit Sub
errhandle1:
    MsgBox cf.GetLastErrorMessage
End Sub

Private Sub UserForm_Initialize()
    Dim strx As String, strx1 As String, strx2 As String, strx3 As String

    If Statisticsx = 1 Then Exit Sub
    txtStatus.Text = "δ����"
    With UserForm20
        strx = .Label5.Caption 'ip
        strx1 = .TextBox6.Text '�û���
        strx2 = .TextBox7.Text '����
        strx3 = .TextBox5.Text '�˿�
    End With
    With Me
        .txtHost.Text = strx
        .txtPassword.Text = strx2
        .txtPort.Text = strx3
        .txtUser.Text = strx1
    End With
End Sub

Private Sub UserForm_Terminate()
    If Statisticsx = 1 Then Exit Sub
    If Not cf Is Nothing Then Set cf = Nothing
End Sub
