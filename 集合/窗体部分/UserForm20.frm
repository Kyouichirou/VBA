VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm20 
   Caption         =   "FTP�ļ�����"
   ClientHeight    =   4575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9795
   OleObjectBlob   =   "UserForm20.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "UserForm20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ipcount As Byte 'ͳ��ip
Dim FilePath As String
Dim filex As String
Dim filecode As String
Dim control As Byte

Private Sub CommandButton2_Click() '�����豸
    Dim arr() As String, xi As Variant, i As Byte, k As Byte
    
    ReDim arr(1 To 50)
    If ipcount = 1 Then MsgShow "δ��ȡ��������IP", "Warning", 1500: Exit Sub
    arr = ObtainIP
    ipcount = ipcount - 1
    With Me.ListView1.ListItems
        .Clear
        For i = 1 To ipcount
            For k = 0 To 4
                xi = Split(arr(i), ".")
                With .Add
                    .Text = xi(0)
                    .SubItems(1) = xi(1)
                    .SubItems(2) = xi(2)
                    .SubItems(3) = CInt(xi(3)) + k
                End With
            Next
        Next
    End With
    Erase arr
End Sub

Private Sub CommandButton3_Click() 'ȷ��
    Dim i As Byte, k As Byte, textx As Object, m As Byte, strx1 As String
    
    If fso.fileexists(FilePath) = False Then MsgBox "�ļ�������", vbCritical: Exit Sub
    With Me
        For i = 1 To 5
            Set textx = .Controls("textbox" & i) 'Ϊ�˷������ؼ���,������Ի��ܵ�ʱ��,���Խ�����˳����һ��
            strx = Trim(textx.Text)
            If IsNumeric(strx) = False Then GoTo 100 '�ж������Ƿ�Ϊ����
            If i = 5 Then Exit For
            k = Len(strx)
            If k = 0 Or k > 3 Then GoTo 100 '��������ĳ���1-255(1-3)
            m = CInt(strx)
            If m < 1 Or m > 255 Then GoTo 100 '���ƴ�С
            strx1 = strx1 & strx & "."
        Next
        strx1 = Left(strx1, Len(strx1) - 1) 'ȷ��ip��ַ����Ҫ��
        If IsPing(strx1) = False Then MsgBox "�豸����δ����״̬", vbCritical, "Warning": Exit Sub '����豸�Ƿ�������״̬
        If CheckPort(strx1, strx) = False Then MsgBox "�豸FTPδ����", vbCritical, "Warning": Exit Sub '����豸��FTP
    
        .Label5.Caption = strx1
        strx3 = Trim(.TextBox7.Text) '����
        strx2 = Trim(.TextBox6.Text) '�˺�
        strx4 = Trim(.TextBox8.Text) '�ļ���λ��
        
        If Len(strx4) = 0 Then
            strx4 = "/"
        Else
            strx4 = "/" & strx4 & "/"
        End If
        
        strx4 = strx4 & Right$(filecode, Len(filecode) - 4) & "." & filex
        '(�ļ����λ��+)���ļ��� '��ʹ������,��Ϊ�ܶ�FTP�����ĵ�֧�ֺܲ�,Ϊ����Ҳ����ʹ��cmd ������ftp, �ļ��������ϴ���ͳһ�޸�Ϊ��Ӧ�ı��
        
        If UpLoadFile(strx1, strx, strx2, strx3, strx4) = True Then
            .Label6.Caption = "�ϴ��ɹ�"
        Else
            .Label6.Caption "�ϴ�ʧ��"
        End If
    End With
    Set textx = Nothing
    Exit Sub
100
    MsgBox "�������������"
    textx.SetFocus
    Set textx = Nothing
End Sub

Function UpLoadFile(ByVal ipaddress As String, ByVal portx As String, ByVal Username As String, ByVal Password As String, ByVal newname As String) As Boolean '�ϴ��ļ����ֻ�FTP
    Dim FTP As FTP
    Set FTP = New FTP
    
    UpLoadFile = False
    If FTP.Connect(ipaddress, portx, Username, Password) = 1 Then '������ӳɹ�
        If FTP.PutFile(FilePath, newname) = 1 Then UpLoadFile = True
    End If
    Set FTP = Nothing
End Function

Private Sub CommandButton4_Click() '����FTP
    UserForm21.Show
End Sub

Private Sub CommandButton5_Click() '��������
    With Me
        If .CommandButton5.Caption = "�޸�����" Then
            .TextBox6.Enabled = True
            .TextBox7.Enabled = True
            .TextBox8.Enabled = True
            .CommandButton5.Caption = "��������"
        Else
            With ThisWorkbook.Sheets("temp")
                .Cells(47, "ab").Value = Me.TextBox6.Text
                .Cells(48, "ab").Value = Me.TextBox7.Text
                .Cells(49, "ab").Value = Me.TextBox8.Text
            End With
            .TextBox6.Enabled = False
            .TextBox7.Enabled = False
            .TextBox8.Enabled = False
            .CommandButton5.Caption = "�޸�����"
        End If
    End With
End Sub

Private Sub CommandButton6_Click() '��ftp
    Dim strx As String, strx1 As String
    Dim textx As Object, i As Byte, k As Byte, m As Integer
    
    On Error GoTo ErrHandle
    With Me
        For i = 1 To 5
            Set textx = .Controls("textbox" & i) 'Ϊ�˷������ؼ���,������Ի��ܵ�ʱ��,���Խ�����˳����һ��
            strx = Trim(textx.Text)
            If IsNumeric(strx) = False Then GoTo ErrHandle
            If i = 5 Then Exit For
            k = Len(strx)
            If k = 0 Or k > 3 Then GoTo ErrHandle
            m = CInt(strx)
            If m < 1 Or m > 255 Then GoTo ErrHandle
            strx1 = strx1 & strx & "."
        Next
    End With
    strx1 = Left(strx1, Len(strx1) - 1) 'ȷ��ip��ַ����Ҫ��
    If CheckPort(strx1, strx) = False Then MsgBox "�豸FTPδ����", vbCritical, "Warning": Set textx = Nothing: Exit Sub '����豸��FTP
    strx1 = "ftp://" & strx1 & ":" & strx
    Shell "explorer.exe " & strx1, vbNormalFocus '����Դ�������ϴ�ftp
    Set textx = Nothing
    Me.CommandButton7.Visible = True
    Exit Sub
ErrHandle:
    textx.SetFocus: Set textx = Nothing: MsgBox "�����������", vbCritical, "Waring"
End Sub

Private Sub CommandButton7_Click()
If Len(FilePath) > 0 Then
    If fso.fileexists(FilePath) = True Then
        CutOrCopyFiles FilePath
    Else
        MsgShow "�ļ�������", "Warning", 1200
    End If
End If
End Sub

Private Sub ListView1_Click() '����б�
    With Me.ListView1
        If .ListItems.Count = 0 Then Exit Sub
        Me.TextBox1.Text = .SelectedItem.Text
        Me.TextBox2.Text = .SelectedItem.SubItems(1)
        Me.TextBox3.Text = .SelectedItem.SubItems(2)
        Me.TextBox4.Text = .SelectedItem.SubItems(3)
    End With
    Me.TextBox5.SetFocus
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim i As Byte
    i = Len(Trim(Me.TextBox1.Text))
    If KeyCode = 13 Then
        If i > 0 And i < 4 Then Me.TextBox2.SetFocus
    End If
End Sub

Private Sub TextBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim i As Byte
    i = Len(Trim(Me.TextBox2.Text))
    If KeyCode = 13 Then
        If i > 0 And i < 4 Then Me.TextBox3.SetFocus
    End If
End Sub

Private Sub TextBox3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer) '�ⲿ��ģ���������дip
    Dim i As Byte
    i = Len(Trim(Me.TextBox3.Text))
    If KeyCode = 13 Then
        If i > 0 And i < 4 Then Me.TextBox4.SetFocus
    End If
End Sub
Private Sub TextBox4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim i As Byte
    i = Len(Trim(Me.TextBox4.Text))
    If KeyCode = 13 Then
        If i > 0 And i < 4 Then Me.TextBox5.SetFocus
    End If
End Sub

Private Sub TextBox6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then Me.TextBox7.SetFocus
End Sub

Private Sub TextBox7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then Me.TextBox8.SetFocus
End Sub

Private Sub UserForm_Initialize()
    Dim strx1 As String, strx2 As String, strx3 As String, textx As Object
    
    If Statisticsx = 1 Then Exit Sub
    Me.TextBox6.SetFocus
    With Me.ListView1
        .ColumnHeaders.Add , , "IP A", 91, lvwColumnLeft
        .ColumnHeaders.Add , , "IP B", 90, lvwColumnLeft
        .ColumnHeaders.Add , , "IP C", 90, lvwColumnLeft
        .ColumnHeaders.Add , , "IP D", 91, lvwColumnLeft
        .View = lvwReport                            '�Ա���ĸ�ʽ��ʾ
        .LabelEdit = lvwManual                       'ʹ���ݲ��ɱ༭
        .Gridlines = True
    End With
    With ThisWorkbook.Sheets("temp")
        strx1 = .Cells(47, "ab")
        strx2 = .Cells(48, "ab")
        strx3 = .Cells(49, "ab")
    End With
    With Me
        .TextBox6.Text = strx1
        .TextBox7.Text = strx2
        .TextBox8.Text = strx3
        If Len(strx1) > 0 And Len(strx2) > 0 And Len(strx3) > 0 Then
        .TextBox6.Enabled = False
        .TextBox7.Enabled = False
        .TextBox8.Enabled = False
        .CommandButton5.Caption = "�޸�����"
        End If
    End With
    If UF3Show > 0 Then
        With UserForm3
            FilePath = Filepathi
            filecode = .Label29.Caption
            filex = .Label24.Caption
        End With
    End If
End Sub

Private Function ObtainIP() As String() '��ȡ������ip��ַ
    Dim obj As Object, ipobj As Object, ip As Variant, pcname As String, arrTemp() As String
    
    ReDim ObtainIP(1 To 50)
    ReDim arrTemp(1 To 50)
    pcname = "localhost"
    Set obj = GetObject("winmgmts:{impersonationLevel=impersonate}//" & pcname).ExecQuery("SELECT index, IPAddress FROM Win32_NetworkAdapterConfiguration")
    ipcount = 1
    For Each ipobj In obj
      If TypeName(ipobj.ipaddress) <> "Null" Then
         For Each ip In ipobj.ipaddress
            If InStr(ip, ".") > 0 Then
               If IsNumeric(Left(ip, InStr(ip, ".") - 1)) = True Then arrTemp(ipcount) = ip: ipcount = ipcount + 1 'ֻ��ȡipv4�ĵ�ַ
            End If
         Next
      End If
    Next
    ObtainIP = arrTemp
End Function


