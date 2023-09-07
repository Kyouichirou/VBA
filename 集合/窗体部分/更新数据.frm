VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 更新数据 
   Caption         =   "更新数据"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8760
   OleObjectBlob   =   "更新数据.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "更新数据"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'
'Private Sub CommandButton1_Click() '部分更新
'
'If Me.ListBox1.ListCount = 0 Or Me.ListBox1.ListIndex = -1 Then Exit Sub
'
'Dim arrt()
'
'ReDim arrt(0 To ListBox1.ListCount - 1)
'
'For i = 0 To ListBox1.ListCount - 1
'
'If ListBox1.Selected(i) = True Then
'   If fso.FolderExists(Me.ListBox1.List(i)) = True Then arrt(i) = Me.ListBox1.List(i) '添加进数组前,检查文件夹是否还存在
'End If
'Next
'
'Unload UserForm1
'
'
'Call fleshdata(arrt)
'
'End Sub
'
'Private Sub CommandButton2_Click() '全部更新
'
'If Me.ListBox1.ListCount = 0 Or Me.ListBox1.ListIndex = -1 Then Exit Sub
'
'Dim arrt()
'
'ReDim arrt(0 To ListBox1.ListCount - 1)
'
'For i = 0 To ListBox1.ListCount - 1
'
'   If fso.FolderExists(Me.ListBox1.List(i)) = True Then arrt(i) = Me.ListBox1.List(i) '添加进数组前,检查文件夹是否还存在
'
'Next
'
'Unload UserForm1
'
'Call fleshdata(arrt)
'
'End Sub
'
'Private Sub UserForm_Initialize()
'
'Dim arr()
'Dim k As Integer, m As Integer
'
'With ThisWorkbook.Sheets("主界面")
'If .Range("e37") = "" Then
'    MsgBox "尚未添加文件夹"
'    Exit Sub
'End If
'
'For k = 38 To 100
'    If .Range("e" & k) = "" Then
'    Exit For
'    End If
'Next
'
'ReDim arr(1 To k - 37)
'm = 37
'For i = 1 To k - 37
'arr(i) = .Range("e" & m).Value
'm = m + 1
'Next
'
'End With
'
'With Me.ListBox1
'
'.MultiSelect = fmMultiSelectMulti
'.ListStyle = fmListStyleOption
'
'End With
'
'Me.ListBox1.List = arr
'
'End Sub
'
'
