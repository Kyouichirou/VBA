VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm12 
   Caption         =   "�ļ�ɾ��"
   ClientHeight    =   2265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9210
   OleObjectBlob   =   "UserForm12.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "UserForm12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim filepathx As String, filext As String, includx As Byte

Private Sub CheckBox1_Click()
    If Len(filepathx) = 0 Then Me.CheckBox1.Value = False
    If Me.CheckBox1.Value = True Then MsgShow "����md5��Ҫһ����ʱ��", "Tips", 1500
End Sub

Private Sub CommandButton1_Click()
    Dim strx As String, strx1 As String, i As Byte
    
    If Len(filepathx) = 0 Then Exit Sub
    With Me
        strx = Trim(.TextBox1.Text)
        strx1 = Trim(.ComboBox1.Text)
        i = Len(strx1)
        If i > 0 Then Reasona = strx1
        If strx1 = "ɾ��ԭ��" Or i = 0 Then Reasona = "һ��ɾ��"
        If Len(strx) > 0 Then Reasonb = strx
        If .CheckBox1.Visible = False Then
            If includx = 0 Then
                filext = UCase(filext)
                If filext <> "EPUB" And filext <> "MOBI" And filext <> "PDF" Then Filehashx = GetFileHashMD5(filepathx) '��ɾ���ļ�ǰ,�Զ�����md5
            Else
                Filehashx = GetFileHashMD5(filepathx)
            End If
        Else
            Filehashx = GetFileHashMD5(filepathx)
        End If
        filepathx = ""
        filext = ""
        includx = 0
        Unload Me
    End With
End Sub

Private Sub UserForm_Initialize()
    Dim strx As String, strx1 As String, strx2 As String

    If Statisticsx = 1 Then Exit Sub
    With ThisWorkbook.Sheets("temp")
        strx = .Range("ab35").Value
        strx1 = .Range("ab36").Value
    End With
    Me.ComboBox1.List = Array("һ��ɾ��", "�¾��ļ�", "�ص��ļ�", "���ݵ���", "�ļ�����")
    If Len(strx) > 0 Then Me.CheckBox1.Visible = False
    If Len(strx1) > 0 Then includx = 1
    With UserForm3
        filepathx = .Label25.Caption
        filext = .Label24.Caption
    End With
End Sub
