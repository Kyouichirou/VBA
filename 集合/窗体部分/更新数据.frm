VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �������� 
   Caption         =   "��������"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8760
   OleObjectBlob   =   "��������.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "��������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'
'Private Sub CommandButton1_Click() '���ָ���
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
'   If fso.FolderExists(Me.ListBox1.List(i)) = True Then arrt(i) = Me.ListBox1.List(i) '��ӽ�����ǰ,����ļ����Ƿ񻹴���
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
'Private Sub CommandButton2_Click() 'ȫ������
'
'If Me.ListBox1.ListCount = 0 Or Me.ListBox1.ListIndex = -1 Then Exit Sub
'
'Dim arrt()
'
'ReDim arrt(0 To ListBox1.ListCount - 1)
'
'For i = 0 To ListBox1.ListCount - 1
'
'   If fso.FolderExists(Me.ListBox1.List(i)) = True Then arrt(i) = Me.ListBox1.List(i) '��ӽ�����ǰ,����ļ����Ƿ񻹴���
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
'With ThisWorkbook.Sheets("������")
'If .Range("e37") = "" Then
'    MsgBox "��δ����ļ���"
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
