VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm13 
   Caption         =   "�����ļ�"
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8520
   OleObjectBlob   =   "UserForm13.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "UserForm13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FilePath As String
Dim filecode As String
Dim filez As Long
Dim Folderpath As String
Dim filex As String
Dim filen As String
Dim arrfilepath(1 To 30) As String
'Dim clickindex As Byte
Dim fc As Byte

Private Sub CommandButton1_Click() '�����ļ�
    Dim strfolder As String, strx As String, tagx As String, strx1 As String, strx2 As String, i As Byte, dics As String
    Dim strx3 As String
    
    If LenB(filecode) = 0 Then Exit Sub
    If LenB(Folderpath) > 0 Then
        strx3 = Folderpath & "\" & filecode '���ļ���
        If fso.folderexists(Folderpath) = True Then
            If fso.folderexists(strx3) = False Then '�������ļ���
                fso.CreateFolder strx3
            Else
                If fso.GetFolder(strx3).Files.Count >= 30 Then
                    MsgBox "�ļ���������30", vbOKOnly, "Warning"
                    OpenFileLocation (strx3)
                    If filex Like "xl*" Then Unload Me
                    Exit Sub
                End If
            End If
        Else
            If InStr(Folderpath, "\") > 0 Then
                fso.CreateFolder Folderpath
                fso.CreateFolder strx3
            Else
                GoTo 100
            End If
        End If
    Else
100
        With Application.FileDialog(msoFileDialogFolderPicker) '�ļ���ѡ�񴰿�(�ļ����Դ�ֻ��ѡһ��)
            .Show
            If .SelectedItems.Count = 0 Then Exit Sub 'δѡ���ļ������˳�sub
            strfolder = .SelectedItems(1)
        End With
        If CheckFileFrom(strfolder, 2) = True Then MsgBox "�ļ���λ������", vbCritical, "Warning"
        ThisWorkbook.Sheets("temp").Cells(40, "ab") = strfolder
        Folderpath = strfolder
        strx3 = Folderpath & "\" & filecode '���ļ���
    End If
    '----------------------------------������Ӧ�ı����ļ��кͶ�Ӧ�����ļ���
    dics = Left$(Folderpath, 1)
    If filez > fso.GetDrive(dics).AvailableSpace Then MsgBox "���̿ռ䲻��!", vbCritical, "Warning": Exit Sub '�жϴ����Ƿ����㹻�Ŀռ�
    
    With Me
        tagx = Trim(.TextBox1.Text) '�Զ���-���Ƴ���
        i = Len(tagx)
        If i > 20 Then
            .Label2.Caption = "�Զ��峤�ȳ�����Χ"
            .TextBox1.SetFocus
            Exit Sub
        End If
        .Label2.Caption = "������...����رմ���"
        strx3 = strx3 & "\" '������ļ���·��
        DoEvents
        fso.CopyFile (FilePath), strx3
        FilePath = strx3 & filen '���ƺ���ļ�·��
    
        If i = 0 Then
            If .ComboBox1.Text = "����ļ�" Then
                tagx = "NA"
            Else
                strx2 = Trim(.ComboBox1.Text)
                If Len(strx2) > 0 Then
                    tagx = strx2
                Else
                    tagx = "NA"
                End If
            End If
        End If
        strx1 = CStr(Format(Now, "yyyymmddhhmmss"))
        strx = strx1 & "-" & tagx & "." & filex '�µ��ļ���-���
        fso.GetFile(FilePath).Name = strx          'ע�����ﲻ��ʹ��cmd���ⲿ������ļ�,��Ϊ��Ҫ�ȴ��ļ��������
        FilePath = strx3 & strx           '�µ��ļ�·��
        .CommandButton1.Enabled = False
        .CommandButton2.Enabled = True
        If fc > 0 Then
            With .ListView1.ListItems.Add
                .Text = strx1
                .SubItems(1) = tagx
                .SubItems(2) = Now
                .SubItems(3) = fso.GetFile(FilePath).DateLastModified
            End With
        End If
        fc = fc + 1
        arrfilepath(fc) = FilePath
        .Label2.Caption = "���ݳɹ�"
    End With
End Sub

Private Sub CommandButton2_Click() '��������
Dim i As Byte
    On Error GoTo 100
    With Me
        If fso.fileexists(FilePath) = True Then fso.DeleteFile (FilePath)
        .CommandButton2.Enabled = False
        .CommandButton1.Enabled = True
        fc = fc - 1
        If fc > 0 Then
            With .ListView1
                .ListItems.Remove (.ListItems.Count)
            End With
            If fc = 0 Then Exit Sub
            For i = 1 To fc
                arrfilepath(i) = arrfilepath(i)
            Next
        Else
            arrfilepath(1) = ""
        End If
    End With
    Exit Sub
100
    If Err.Number = 70 Then
        Me.Label2.Caption = "�ļ����ڴ򿪵�״̬"
    Else
        Me.Label2.Caption = "�쳣"
    End If
    Err.Clear
End Sub

Private Sub CommandButton3_Click() '�鿴����
    Dim fd As Folder, i As Byte, litm As Variant, filetemp As Variant, strx As String
    Dim fl As File, strx1 As String, strx2 As String, strx3 As String
    
    With Me.ListView1.ListItems
        If fso.folderexists(Folderpath) = True Then
            strx1 = Folderpath & "\" & filecode '�����ļ���·��
            If fso.folderexists(strx1) = True Then
                Set fd = fso.GetFolder(strx1)
                For Each fl In fd.Files
                    strx = fl.Name
                    If InStr(strx, "-") = 0 Then GoTo 100 '���ϸ�ʽҪ��
                    i = i + 1
                    If i = 31 Then GoTo 1000 '����30
                    arrfilepath(i) = fl.Path
                    filetemp = Split(strx, "-")
                    Set litm = .Add()
                    litm.Text = filetemp(0)
                    strx2 = filetemp(1)
                    strx3 = Left$(strx2, InStrRev(strx2, ".") - 1) '״̬
                    litm.SubItems(1) = strx3
                    litm.SubItems(2) = fl.DateCreated
                    litm.SubItems(3) = fl.DateLastModified
100
                Next
            Else
                Me.Label2.Caption = "��δ��������"
            End If
        Else
            Me.Label2.Caption = "��δ��������"
        End If
    End With
    Me.CommandButton4.Enabled = True
1000
    fc = i
    Set litm = Nothing
    Set fd = Nothing
End Sub

Private Sub CommandButton4_Click() 'ɾ������
    Dim i As Byte, k As Integer, j As Byte
    
    With Me.ListView1
        i = .ListItems.Count
        If i = 0 Then Exit Sub
        For k = i To 1 Step -1  '�ڵ���ɾ����ʱ��,�������Ͳ�����byte����(0-255)
        If .ListItems(k).Selected = True Then
            On Error GoTo 100 '����ļ����ڴ򿪵�״̬
            If fso.fileexists(arrfilepath(k)) = True Then fso.DeleteFile (arrfilepath(k))
            .ListItems.Remove (k)
        Else
            j = j + 1
            arrfilepath(j) = arrfilepath(k)
        End If
100
        Next
        fc = j
    End With
End Sub

Private Sub CommandButton5_Click() '����ļ���
    Dim strx As String
    
    If LenB(Folderpath) = 0 Then Exit Sub
    On Error GoTo 100
    strx = Folderpath & "\" & filecode
    If fso.folderexists(strx) = True Then
        fso.DeleteFolder (strx)
        Me.ListView1.ListItems.Clear
        Me.Label2.Caption = "�ɹ�����"
    End If
    Exit Sub
100
    If Err.Number = 70 Then
        Me.Label2.Caption = "�ļ����е��ļ����ڴ�״̬"
    Else
        Me.Label2.Caption = "�쳣"
    End If
End Sub

Private Sub ListView1_DblClick() '˫�����ļ�
    Dim i As Byte, strx As String
    
    With Me.ListView1
        If .ListItems.Count = 0 Then Exit Sub
        If UCase(filex) Like "XL*" Then Exit Sub
        i = .SelectedItem.Index
        If ErrCode(arrfilepath(i), 1) > 1 Then strx = "ERC"
        Call OpenFile("a", filen, filex, arrfilepath(i), 1, strx, 1)
    End With
End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    With Me
        Select Case KeyAscii
            Case asc("/"), asc("\"), asc(":"), asc("*"), asc("?"), asc("<"), asc(">"), asc("|")
            .Label2.ForeColor = &HFF&
            .Label2.Caption = "��������Ƿ��ַ���""/\ : * ? <> |"
            .TextBox1.Text = ""
            KeyAscii = 0
        Case Else
            .Label2.Caption = ""
        End Select
    End With
End Sub

Private Sub UserForm_Initialize()
    If Statisticsx = 1 Then Exit Sub
    With Me.ListView1 '���ڳ�ʼ��
        .ColumnHeaders.Add , , "���", 94, lvwColumnLeft
        .ColumnHeaders.Add , , "״̬", 94, lvwColumnLeft
        .ColumnHeaders.Add , , "����ʱ��", 94, lvwColumnLeft
        .ColumnHeaders.Add , , "�޸�ʱ��", 94, lvwColumnLeft
        .View = lvwReport                            '�Ա���ĸ�ʽ��ʾ
        .LabelEdit = lvwManual                       'ʹ���ݲ��ɱ༭
        .Gridlines = True
        .MultiSelect = True '֧�ֶ�ѡ
    End With
    Me.TextBox1.SetFocus
    Me.ComboBox1.List = Array("ԭ��", "��ʼ��", "��չA", "��չB", "��չC", "���")
    With UserForm3
        FilePath = Filepathi
        If LenB(FilePath) = 0 Then Exit Sub
        filecode = .Label29.Caption
        filez = .Label76.Caption
        filex = .Label24.Caption
        filen = Filenamei
    End With
    Folderpath = ThisWorkbook.Sheets("temp").Cells(40, "ab").Value
End Sub
