VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PDFDetail 
   Caption         =   "PDF�ļ�����"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7515
   OleObjectBlob   =   "PDFDetail.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "PDFDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'��ҪAdobe acrobat�����֧��
Dim EditSave As Boolean
Dim objAcrobatPDDoc As Object
Dim objJSO As Object
Dim FilePath As String
'--------------------------PDF�������Ϣ��Ҫ������XML

Private Sub CommandButton1_Click() '�޸�
    Dim textbox As Object
    
    If IsEditable = True Then Exit Sub
    Dim i As Byte
    For i = 1 To 4
        Set textbox = Me.Controls("TB" & i)
        textbox.Enabled = True
    Next
    Me.CommandButton1.Enabled = False
    Me.CommandButton2.Enabled = True
    EditSave = False
End Sub

Private Sub CommandButton2_Click() '����
    Edit
    Me.CommandButton1.Enabled = True
    Me.CommandButton2.Enabled = False
End Sub

Private Sub Edit() '����д��
    Dim arr(1 To 4) As String
    Dim textbox As Object
    For i = 1 To 4
        Set textbox = Me.Controls("TB" & i)
        arr(i) = Trim(textbox.Text)
        textbox.Enabled = False
    Next
    With objJSO.info
        .Title = arr(1)
        .Author = arr(2)
        .Subject = arr(3)
        .Keywords = arr(4)
    End With
    Set textbox = Nothing
    EditSave = True
    objAcrobatPDDoc.Save PDSaveFull + PDSaveLinearized + PDSaveCollectGarbage, FilePath
End Sub

Private Sub UserForm_Initialize()
    Dim Titlex As Variant
    Dim lb As Object
    Dim k As Byte
    Dim h As Integer
    Dim textbox As Object
    Dim arr(1 To 11) As String
    Dim strx As String
    Dim IsEditable As Boolean
    
    If Statisticsx = 1 Then Exit Sub
    On Error Resume Next
    If Len(FilePath) = 0 Then Exit Sub
    Set objAcrobatPDDoc = CreateObject("AcroExch.PDDoc")
    objAcrobatPDDoc.Open (FilePath)
    With objAcrobatPDDoc '--------------ֱ�Ӵ�pddoc��ȡ��Ϣ,��objJSO.info��ʽ��ȡ����Ϣ��Ϊ����,����ʱ��������ڵ�ʱ��
          arr(1) = .GetInfo("Title")
          arr(2) = .GetInfo("Author")
          arr(3) = .GetInfo("Subject")
          arr(4) = .GetInfo("Keywords")
          arr(5) = .GetNumPages
          strx = .GetInfo("Creator")
          strx = StrConv(strx, vbFromUnicode) 'ȥ����ansi�ַ�
          strx = StrConv(strx, vbUnicode)
          strx = Replace(strx, "?", " ")
          arr(7) = strx
          strx = .GetInfo("Producer")
          strx = StrConv(strx, vbFromUnicode)
          strx = StrConv(strx, vbUnicode)
          strx = Replace(strx, "?", " ")
          arr(8) = strx
          arr(9) = .GetInfo("CreationDate")
          arr(10) = .GetInfo("ModDate")
          arr(11) = .GetInfo("Rights")
          Set objJSO = .GetJSObject
          If Err.Number > 0 Then Err.Clear '------------------------------------'���ļ����ڷǼ��ܵ�״̬ʱ�����
          If IsNull(objJSO.securityHandler) = False Then arr(6) = "Encrypted": IsEditable = False '�ж��ļ��Ƿ��ڼ��ܵ�״̬
          If Err.Number > 0 Then Err.Clear: arr(6) = "UnEncrypted": IsEditable = True
    End With
    Titlex = Array("Title:", "Author:", "Subject:", "Keywords:", "Pages:", "Encrypted:", "Creator:", "Producer:", "CreationDate:", "ModDate:", "CopyRight:")
    h = 18
    For k = 1 To 11
        Set lb = Me.Controls.Add("Forms.Label.1", "LB" & k, True) '��̬������ǩ
        Set textbox = Me.Controls.Add("Forms.TextBox.1", "TB" & k, True) '��̬�����ı���
        With lb
            .Caption = Titlex(k - 1)
            .TextAlign = fmTextAlignRight
            .Height = 16
            .Left = 8
            .Width = 72
            .Top = h + 2 '����λ��
        End With
        With textbox
            .Text = arr(k)
            .Height = 16
            .Left = 84
            .Width = 240
            .Top = h
            .Enabled = False
        End With
        h = h + 18
    Next
    EditSave = True
    If IsEditable = False Then Me.CommandButton1.Enabled = False: Me.Caption = "�ļ����ڼ���״̬(���ɱ༭)" '���������м����ļ������ɱ༭��Ϊ�������ͳһ�����ܵ��ļ�����Ϊ���ɱ༭��״̬
    Set lb = Nothing
    Set textbox = Nothing
End Sub

Private Sub UserForm_Terminate()
    Dim yesno As Variant
    
    If Statisticsx = 1 Then Exit Sub
    If EditSave = False Then
        yesno = MsgBox("�Ƿ񱣴��޸�?", vbYesNo, "Tips")
        If yesno = vbYes Then Edit
    End If
    objAcrobatPDDoc.Close
    Set objAcrobatPDDoc = Nothing
    Set objJSO = Nothing
End Sub
