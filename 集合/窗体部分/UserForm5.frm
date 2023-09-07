VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "setup"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7350
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   1  '所有者中心
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
'Private Sub CommandButton2_Click() '修改设置
'Dim exepatha1 As String
'Dim exepatha2 As String
'Dim exepath As String
'
'If Me.TextBox1.Text = "" And Me.TextBox2 = "" Then '两者都为空的时候
'MsgBox "请设置需要调用的程序"
'Me.TextBox1.SetFocus
'Exit Sub
'End If
'
'
'If Me.TextBox1.Text <> "" And Me.TextBox2 <> "" Then                                                                               '两个设置都设置时
'   If fso.FileExists(repalce(Me.TextBox1.Text, " ", "")) = False Or fso.FileExists(repalce(Me.TextBox2.Text, " ", "")) = False Then
'   MsgBox "设置有误,请仔细检查"
'   Me.TextBox1.SetFocus
'   Exit Sub
'   Else
'   exepatha1 = Me.TextBox1.Text
'   exepatha2 = Me.TextBox2.Text
'   ThisWorkbook.VBProject.VBComponents.Item(8).CodeModule.ReplaceLine 29, "exepath = " & "" & exepatha1 & ""          '注意这里代码的位置在后续修改中可能出现的位置变化
'   ThisWorkbook.VBProject.VBComponents.Item(9).CodeModule.ReplaceLine 437, "exepath = " & "" & exepatha2 & ""
'End If
'
'If Me.TextBox1.Text <> "" And Me.TextBox2.Text = "" Then
'   If fso.FileExists(repalce(Me.TextBox1.Text, " ", "")) = False Then
'   MsgBox "设置有误,请仔细检查"
'   Me.TextBox1.SetFocus
'   Exit Sub
'   Else
'   exepatha1 = Me.TextBox1.Text
'   ThisWorkbook.VBProject.VBComponents.Item(8).CodeModule.ReplaceLine 29, "exepath = " & "" & exepatha1 & ""
'End If
'
'If Me.TextBox2.Text <> "" And Me.TextBox1.Text = "" Then
'   If fso.FileExists(repalce(Me.TextBox2.Text, " ", "")) = False Then
'   MsgBox "设置有误,请仔细检查"
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
