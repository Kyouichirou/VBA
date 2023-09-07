VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm25 
   Caption         =   "�ı�ת����"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17880
   OleObjectBlob   =   "UserForm25.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "UserForm25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------https://wutils.com/com-dll/constants/constants-SpeechLib.htm
Private Const SAFT48kHz16BitStereo = 39
Private Const SSFMCreateForWrite As Byte = 3 ' Creates file even if file exists and so destroys or overwrites the existing file
Private Const SSFMOpenForRead As Byte = 0
Private Const SSFMOpenReadWrite As Byte = 1
Private Const SSFMCreate As Byte = 2
Private Const FlagsAsync As Byte = 1
Dim oVoice As Object, oFileOpen As Object
Dim Flapx As Boolean
Dim voi As Byte, Vox As Byte, vor As Integer
Dim FilePath As String
Private Declare Function timeGetTime Lib "winmm.dll" () As Long 'ʱ��api���Ծ�ȷ������
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim oStream As Object '�ļ���
Dim IsClose As Boolean
Dim Lbox1 As Integer

Private Function ObtainText(ByVal Urlx As String, ByVal FilePath As String) As Boolean 'ץȡҳ��С˵����
    Dim fl As Object, FilePath As String
    Dim bookt As Object
    Dim strHtml As String, textl As Object, HtmF As Object
    Dim XmlH As Object
    Dim rt As Long
    Dim i As Integer, strTemp As String
    Dim k As Integer, j As Integer, p As Integer
    
    ObtainText = True
    Set HtmF = CreateObject("htmlfile")
    Set XmlH = CreateObject("Msxml2.ServerXMLHTTP") 'https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms766431(v=vs.85)
    HtmF.DesignMode = "on" ' �����༭ģʽ
    With XmlH
        .Open "GET", Urlx, True 'https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms757849(v=vs.85)
        .Send
        rt = timeGetTime
        Do While .readyState <> 4 And timeGetTime - rt < 6000 '�ȴ���Ϣ����
            Sleep 25
            DoEvents
        Loop
        If .readyState <> 4 Then 'û�гɹ���ȡ����Ϣ
            XmlH.Close
            ObtainText = False
            Set XmlH = Nothing
            Exit Function
        End If
        strHtml = .responseText
        XmlH.Close
    End With
    HtmF.Write strHtml ' д������
    Set bookt = HtmF.getElementById("BookText") '����id��ȡ����
    Set fl = fso.OpenTextFile(FilePath, ForWriting, True, TristateUseDefault)
    k = bookt.Children.Length - 1
    j = bookt.ChildNodes.Length - 1 '���������������Ϊ׼,������ֱ��ʹ��elementcount
    For Each textl In bookt.Children 'ÿ���ڵ�
        strTemp = Trim(textl.innertext)
        If Len(strTemp) > 0 Then
            If p < k Then
                fl.WriteLine strTemp '������д��txt
                p = p + 1
            End If
        Else
            i = i + 1 '���ڼ���ж���Ϣ�洢��λ��
            If i = 3 Then ObtainText = False: Exit For
        End If
    Next
    If i = 3 Then
        strTemp = ""
        i = 0
        p = 0
        ObtainText = True
        For Each textl In bookt.ChildNodes
            If textl.NodeType = 3 Then
                strTemp = Trim(textl.Data)
                If Len(strTemp) > 0 Then
                    If p < j Then   '-------����һ���ڵ����޹���Ϣ
                        fl.WriteLine strTemp
                        p = p + 1
                    End If
                Else
                    i = i + 1
                    If i = 3 Then ObtainText = False: Exit For
                End If
            End If
        Next
    End If
    fl.Close
    Set fl = Nothing
    Set XmlH = Nothing
    Set HtmF = Nothing
End Function

Private Function TextSpeak(ByVal strText As String) '���ı�ת��Ϊ����
    If Vox = 0 Then Set oVoice = CreateObject("SAPI.SpVoice"): Vox = Vox + 1
    oVoice.Speak strText, FlagsAsync '�첽����
End Function

Private Sub CommandButton1_Click()
    If IsClose = True Then Exit Sub
    With Me.CommandButton1
        If .Caption = "��ͣ" Then
            .Caption = "����"
            oVoice.Pause
        Else
            .Caption = "��ͣ"
            oVoice.Resume
        End If
    End With
End Sub

Private Sub CommandButton10_Click() '��ת����λ��
If IsClose = True Then Exit Sub

End Sub

Private Sub CommandButton11_Click() '�Ƴ���Ƶ�ļ�
    Dim strx As String
    Dim FilePath As String
    
    strx = Trim(Me.TextBox2.Text)
    If Len(strx) = 0 Then Exit Sub
    FilePath = ThisWorkbook.Path & "\" & Format(Now, "yyyymmddhhmmss") & ".wav"
    If TextToVoice(FilePath, 0, strText:=strx) = True Then
        strx = "ת���ɹ�"
    Else
        strx = "ת��ʧ��"
    End If
    MsgShow strx, "Tips", 1200
End Sub

Private Sub CommandButton12_Click() '��ȡ�ı�
    If IsClose = True Then Exit Sub
    If Not oStream Is Nothing Then Me.TextBox2.Text = GetString(sPosition:=0)
End Sub

Private Sub CommandButton13_Click() '����д���ı�

End Sub

Private Sub CommandButton14_Click() '�����ı�


End Sub

Private Sub CommandButton2_Click() '�ر�
    If IsClose = True Then Exit Sub
    oVoice.Pause
    Set oVoice = Nothing
    fl.Close
    Set fl = Nothing
    IsClose = True
End Sub
'-----------------------------------------------Values for the Volume property range from 0 to 100
Private Sub CommandButton3_Click() '����+
    If IsClose = True Then Exit Sub
    If voi = 100 Then
        Me.CommandButton3.Enabled = False
    Else
        voi = voi + 10
        oVoice.Volume = voi
    End If
End Sub

Private Sub CommandButton4_Click() '����-
    If IsClose = True Then Exit Sub
    If voi = 0 Then
        Me.CommandButton4.Enabled = False
    Else
        voi = voi - 10
        oVoice.Volume = voi
    End If
    Me.CommandButton3.Enabled = True
End Sub

Private Sub ListBox1_Click() '���������������Ľ��
    If Lbox1 = -1 Then Exit Sub
    With Me
        If .ListBox1.ListCount = 0 Then Exit Sub
        If Lbox1 <> .ListBox1.ListIndex Then
            If .CommandButton5.Caption <> "��" Then .CommandButton5.Caption = "��"
        Else
            .CommandButton5.Caption = "ǰ��"
        End If
    End With
End Sub

Private Sub CommandButton5_Click() '������
    Dim strx As String, yesno As Variant
    Dim arr() As String
    Dim i As Integer, j As Byte
    Dim strTemp As String
    
    With Me
        strTemp = .CommandButton5.Caption
        If strTemp = "��" Then
            If .ListBox1.ListCount = 0 Then Exit Sub
            i = .ListBox1.ListIndex
            If i = -1 Then Exit Sub '-1��ʾδѡ��
            Lbox1 = i
            arr = ObtainPage_Info(.ListBox1.List(i, 1), 0)
            If Pagexs > 0 Then
                .ComboBox1.Text = 1
                For j = 1 To Pagexs
                    .ComboBox1.AddItem j
                Next
            End If
            j = UBound(arr)
            With .ListBox2
                For i = 0 To j
                    .AddItem
                    .List(i, 0) = arr(i, 0)
                    .List(i, 1) = arr(i, 1)
                Next
            End With
            .ListBox1.Visible = False
            .ListBox2.Visible = True
            .CommandButton5.Caption = "����"
            .CommandButton9.Visible = True
        ElseIf strTemp = "����" Then
            .CommandButton9.Visible = False
            .ListBox1.Visible = True
            .ListBox2.Visible = False
            .CommandButton5.Caption = "ǰ��"
        ElseIf strTemp = "ǰ��" Then
            .ListBox1.Visible = False
            .ListBox2.Visible = True
            .CommandButton9.Visible = True
            .CommandButton5.Caption = "����"
        End If
    End With
End Sub

Private Sub CommandButton9_Click() '����
    Dim i As Integer
    Dim FilePath As String
    Dim folderx As Folder, fl As File
    Dim strx As String
    Dim strTemp As String
    '----------�ȼ�鱾���ļ��Ƿ����, ������������Ȳ��ű����ļ�
    With Me
        If .ListBox2.ListCount = 0 Then Exit Sub
        i = .ListBox2.ListIndex
        If i = -1 Then Exit Sub
        strx = .ListBox2.List(i, 0)
        strTemp = ThisWorkbook.Path & "\" & .ListBox1.List(Lbox1, 0)
        FilePath = strTemp
        If fso.folderexists(FilePath) = True Then '�鿴�Ƿ���ڱ����ļ���
            Set folderx = fso.GetFolder(FilePath)
            FilePath = ""
            For Each fl In folderx.Files
            If strx = fl.Name Then
                If fl.Size > 0 Then FilePath = fl.Path: Exit For
            End If
            Next
            Set folderx = Nothing
            If Len(FilePath) = 0 Then
                If IsNetConnectOnline = False Then
                    MsgBox "���粻����", vbInformation, "Tips": Exit Sub
                Else
                    FilePath = strTemp & "\" & strx & ".txt"
                    ObtainText strx, FilePath
                End If
            Else
            FilePath = strTemp & "\" & strx & ".txt"
            TextSpeak (GetString(FilePath))
            End If
        Else
        fso.CreateFolder FilePath '���������ļ���
        End If
    End With
IsClose = False
End Sub

Private Sub CommandButton6_Click() '����
    Dim strx As String
    Dim arr() As String
    Dim i As Integer
    Dim k As Integer
    
    strx = Trim(Me.TextBox1.Text)
    If Len(strx) < 2 Then Exit Sub
    If IsNetConnectOnline = False Then MsgShow "����δ����", "Tips", 1200: Exit Sub
    arr = ObtainPage_Info(strx, 1)
    i = UBound(arr)
    With Me.ListBox1 '���������������
        For k = 0 To i
            .AddItem
            .List(k, 0) = arr(k, 0)
            .List(k, 1) = arr(k, 1)
        Next
    End With
    Me.CommandButton5.Caption = "��"
End Sub
'-----------------------------------------------Values for the Rate property range from -10 to 10
Private Sub CommandButton7_Click() '�ٶ�+
    If IsClose = True Then Exit Sub
    If vor = 10 Then
        Me.CommandButton7.Enabled = False
    Else
        vor = vor + 1
        oVoice.Volume = vor
    End If
    Me.CommandButton8.Enabled = True
End Sub

Private Sub CommandButton8_Click() '�ٶ�-
    If IsClose = True Then Exit Sub
    If vor = -10 Then
        Me.CommandButton8.Enabled = False
    Else
        vor = vor - 1
        oVoice.Volume = vor
    End If
    Me.CommandButton7.Enabled = True
End Sub

Private Sub UserForm_Initialize()
    If Statisticsx = 1 Then Exit Sub
    voi = 50
    vor = 0
    Flapx = False
    IsClose = True
    Lbox1 = -1
    With Me
        .ComboBox2.List = Array("10%", "20%", "30%", "40%", "50%", "60%", "70%", "80%", "90%")
    End With
End Sub

Private Function GetString(Optional ByVal FilePath As String, Optional ByVal sCharset As String = "gb2312", Optional ByVal sPosition As Long = 0) As String
    Const adTypeBinary As Byte = 1
    Const adTypeText As Byte = 2
    Const adModeRead As Byte = 1
    Const adModeWrite As Byte = 2
    Const adModeReadWrite As Byte = 3
    
    '----------------------��ָ����λ�ö�ȡ��Ϣ
    If oStream Is Nothing Then
        If fso.fileexists(FilePath) = False Then Exit Function
        Set oStream = CreateObject("ADODB.Stream")
        With oStream
            .Mode = 3
            .type = adTypeText
            .CharSet = sCharset
            .Open
            .LoadFromFile (FilePath)
        End With
    Else
        With oStream
            .Position = sPosition
            GetString = .ReadText()
        End With
    End If
End Function

Private Sub UserForm_Terminate()
    If Statisticsx = 1 Then Exit Sub
    If IsClose = True Then Exit Sub
    If Not oStream Is Nothing Then
        oStream.Close
        Set oStream = Nothing
    End If
End Sub
