VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "setup"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7350
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Private Sub CommandButton1_Click()
'
'Unload Me
'
'End Sub
'
'Private Sub CommandButton2_Click() '�޸�����
'Dim exepatha1 As String
'Dim exepatha2 As String
'Dim exepath As String
'
'If Me.TextBox1.Text = "" And Me.TextBox2 = "" Then '���߶�Ϊ�յ�ʱ��
'MsgBox "��������Ҫ���õĳ���"
'Me.TextBox1.SetFocus
'Exit Sub
'End If
'
'
'If Me.TextBox1.Text <> "" And Me.TextBox2 <> "" Then                                                                               '�������ö�����ʱ
'   If fso.FileExists(repalce(Me.TextBox1.Text, " ", "")) = False Or fso.FileExists(repalce(Me.TextBox2.Text, " ", "")) = False Then
'   MsgBox "��������,����ϸ���"
'   Me.TextBox1.SetFocus
'   Exit Sub
'   Else
'   exepatha1 = Me.TextBox1.Text
'   exepatha2 = Me.TextBox2.Text
'   ThisWorkbook.VBProject.VBComponents.Item(8).CodeModule.ReplaceLine 29, "exepath = " & "" & exepatha1 & ""          'ע����������λ���ں����޸��п��ܳ��ֵ�λ�ñ仯
'   ThisWorkbook.VBProject.VBComponents.Item(9).CodeModule.ReplaceLine 437, "exepath = " & "" & exepatha2 & ""
'End If
'
'If Me.TextBox1.Text <> "" And Me.TextBox2.Text = "" Then
'   If fso.FileExists(repalce(Me.TextBox1.Text, " ", "")) = False Then
'   MsgBox "��������,����ϸ���"
'   Me.TextBox1.SetFocus
'   Exit Sub
'   Else
'   exepatha1 = Me.TextBox1.Text
'   ThisWorkbook.VBProject.VBComponents.Item(8).CodeModule.ReplaceLine 29, "exepath = " & "" & exepatha1 & ""
'End If
'
'If Me.TextBox2.Text <> "" And Me.TextBox1.Text = "" Then
'   If fso.FileExists(repalce(Me.TextBox2.Text, " ", "")) = False Then
'   MsgBox "��������,����ϸ���"
'   Me.TextBox2.SetFocus
'   Exit Sub
'   Else
'   exepatha2 = Me.TextBox2.Text
'   ThisWorkbook.VBProject.VBComponents.Item(9).CodeModule.ReplaceLine 437, "exepath = " & "" & exepatha2 & ""
'End If
'
'Unload Me
'
'End Sub
