VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm8 
   Caption         =   "�����ļ���"
   ClientHeight    =   6900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7800
   OleObjectBlob   =   "UserForm8.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "UserForm8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim arrlx(1 To 50) '��ʱ�洢key
'Dim arrch(1 To 50) '�洢ѡ���ļ�������Ŀ��key
'Dim ich As Integer
'Dim s As Integer
'
'Private Sub CommandButton1_Click() '��������
'Dim rnglistx As Range
'Dim filexpath As String
'Dim k As Integer, addcodex As Integer
'
'With Me
'If .TreeView1.SelectedItem.Index = 1 Then Exit Sub
'If .CheckBox1.Value = True Then '��ѡ
'addcodex = 1
'Else
'addcodex = 2
'End If
'k = .TreeView1.SelectedItem.Index
'filexpath = .TreeView1.Nodes(k).key
'Call Listallfiles(addcodex, filexpath)
'
'If InStr(.TreeView1.Nodes(k).text, "�����仯") > 0 Then
'.TreeView1.Nodes(k).text = Left(.TreeView1.Nodes(k).text, Len(.TreeView1.Nodes(k).text) - 6)
'ElseIf InStr(.TreeView1.Nodes(k).text, "�����") = 0 Then .TreeView1.Nodes(k).text = .TreeView1.Nodes(k).text & "(�����)" '�������
'End If
'
'If .TreeView1.Nodes(k).Children > 0 Then '���������ļ���
'If addcodex = 1 Then Call CheckTreelists(.TreeView1, .TreeView1.Nodes(k))
'    For i = 1 To ich - 1                                '��������һ��ֵ����ѡ����
'        If InStr(.TreeView1.Nodes(arrch(i)).text, "�����仯") > 0 Then
'        .TreeView1.Nodes(arrch(i)).text = Left(.TreeView1.Nodes(arrch(arrch(i))).text, Len(.TreeView1.Nodes(arrch(i)).text) - 6) 'ȥ�������е�(�����仯)
'        ElseIf InStr(.TreeView1.Nodes(arrch(i)).text, "�����") = 0 Then .TreeView1.Nodes(arrch(i)).text = .TreeView1.Nodes(arrch(i)).text & "�����" '�������,���������
'        End If
'    Next
'End If
'End With
'
'End Sub
'
'Private Sub CommandButton2_Click() 'չ�����еĽڵ�
'Dim i As Integer
'With Me.TreeView1
'For i = 1 To .Nodes.Count
'.Nodes(i).Expanded = True
'Next
'End With
'End Sub
'
'Private Sub CommandButton3_Click() '�۵����нڵ�
'Dim i As Integer
'With Me.TreeView1
'For i = 2 To .Nodes.Count
'.Nodes(i).Expanded = False
'Next
'End With
'End Sub
'
'Private Sub TreeView1_DblClick() '˫��չ��
'With Me
'.TreeView1.Nodes(.TreeView1.SelectedItem.Index).Expanded = True
'End With
'End Sub
'
''Private Sub TreeView1_NodeCheck(ByVal node As MSComctlLib.node)
''Dim nd As node
''With Me
'' .TreeView1.Nodes(.TreeView1.SelectedItem.Index).Child.Checked = True '���ѡ, ���ѡ
''For Each nd.Child In .TreeView1.Nodes(.TreeView1.SelectedItem.Index).Children
''nd.Child.Checked = True
''Next
''.TreeView1.Nodes(.TreeView1.Nodes(.TreeView1.SelectedItem.Index).Child.Index).Checked = True
''MsgBox .TreeView1.Nodes(.TreeView1.SelectedItem.Index).Children
'
''End With
'
''Call CheckTreelists(Me.TreeView1, Me.TreeView1.Nodes(Me.TreeView1.SelectedItem.Index), True)
''Me.TreeView1.Nodes(1).Checked = False
''
''End Sub
'
'Private Sub TreeView1_NodeClick(ByVal node As MSComctlLib.node)
'
'With Me
'.Label1.Caption = .TreeView1.Nodes(.TreeView1.SelectedItem.Index).key 'ѡ����ʾ·��
'.TreeView1.SelectedItem.Bold = True
'
'For i = 2 To .TreeView1.Nodes.Count
'If i = .TreeView1.SelectedItem.Index Then GoTo 100
'.TreeView1.Nodes(i).Bold = False
'100
'Next
'
'End With
'
'End Sub
'
'Private Sub UserForm_Initialize()
'
'Dim dic As New Dictionary '�洢ȥ�ص����ļ���
'Dim arr()
'Dim arrx()
'Dim i As Integer, l As Integer, k As Integer
'
'With Me.TreeView1
'.Appearance = cc3D
'.HotTracking = True
'.Nodes.Add , , "Menus", "Menus" '��Ŀ¼
'.Nodes(1).Expanded = True
'.Nodes(1).Bold = True
'End With
'
'With ThisWorkbook.Sheets("������")
'If .Range("e37") = "" Or InStr(.Range("e37").Value, "\") = 0 Then Exit Sub
'ReDim arr(1 To .[e65536].End(xlUp).Row - 36)
'For i = 37 To .[e65536].End(xlUp).Row
'arr(i - 36) = Split(.Range("e" & i), "\")(0) & "\" & Split(.Range("e" & i), "\")(1) '���ϲ��Ŀ¼
'Next
'End With
'
'For k = 1 To UBound(arr)
'dic(arr(k)) = ""
'Next
'ReDim arrx(0 To UBound(dic.Keys))
'For l = 0 To UBound(dic.Keys)
'arrx(l) = dic.Keys(l)
'Next
'listfolderx arrx
'
'End Sub
'
'Function listfolderx(arrt())
'Dim fd As Folder
'Dim i As Integer
'Dim showname As String
'Dim rnglistx As Range
'Dim strp As String
'
'For i = 0 To UBound(arrt())
's = 1                         'ע���������¹�1����
'Set fd = fso.GetFolder(arrt(i))
'arrlx(s) = fd.Path
'With ThisWorkbook.Sheets("Ŀ¼")
'strp = fd.Path & "\"
'Set rnglistx = .Cells(4, 3).Resize(.[b65536].End(xlUp).Row, .Cells.SpecialCells(xlCellTypeLastCell).Column).Find(strp, lookat:=xlWhole)
'If rnglistx Is Nothing Then
'showname = fd.Name
'Else
'    If fd.DateLastModified <> rnglistx.Offset(0, 2) Then '�ļ��е��޸�ʱ�䷢���仯(��ζ���ļ���(���������е����ļ���)��һ�㷢���仯,�޸�/ɾ��/�½��ļ���/�޸��ļ���)
'    showname = fd.Name & "(�����)(�����仯)"
'    Else
'    showname = fd.Name & "(�����)"
'    End If
'End If
'End With
'
'With UserForm8.TreeView1.Nodes
'    .Add "Menus", 4, arrlx(s), showname
'End With
'listfolderxs fd
'Next
'
'Set fd = Nothing
'
'End Function
'
'Function listfolderxs(ByVal fd As Folder)
'Dim sfd As Folder
'Dim showname As String
'Dim rnglistx As Range
'Dim strp As String
'If fd.SubFolders.Count = 0 Then Exit Function '���ļ�����ĿΪ�����˳�sub
'
'For Each sfd In fd.SubFolders
'    With ThisWorkbook.Sheets("Ŀ¼")
'    strp = sfd.Path & "\"
'    Set rnglistx = .Cells(4, 3).Resize(.[b65536].End(xlUp).Row, .Cells.SpecialCells(xlCellTypeLastCell).Column).Find(strp, lookat:=xlWhole)
'        If rnglistx Is Nothing Then
'        showname = sfd.Name
'        Else
'        If sfd.DateLastModified <> rnglistx.Offset(0, 2) Then '�ļ��е��޸�ʱ�䷢���仯(��ζ���ļ���(���������е����ļ���)��һ�㷢���仯,�޸�/ɾ��/�½��ļ���/�޸��ļ���)
'        showname = sfd.Name & "(�����)(�����仯)"
'        Else
'        showname = sfd.Name & "(�����)"
'        End If
'        End If
'    End With
'    With UserForm8.TreeView1.Nodes
'        If arrlx(s + 1) <> sfd.Path Then
'        arrlx(s + 1) = sfd.Path
'        .Add arrlx(s), 4, arrlx(s + 1), showname
'        End If
'    End With
'    If sfd.SubFolders.Count > 0 Then s = s + 1
'    listfolderxs sfd
'Next
's = s - 1 '����
'End Function
'Function CheckTreelists(ByRef treevw As TreeView, ByRef nodThis As node) '�ӽڵ�,��ѡ
'Dim lngIndex As Integer
'If nodThis.Children > 0 Then
'lngIndex = nodThis.Child.Index
'Call CheckTreelists(treevw, treevw.Nodes(lngIndex))
'
'While lngIndex <> nodThis.Child.LastSibling.Index
'  lngIndex = treevw.Nodes(lngIndex).Next.Index
'  Call CheckTreelists(treevw, treevw.Nodes(lngIndex))
'Wend
'End If
'ich = ich + 1
'arrch(ich) = nodThis.Index
'End Function
'
''Private Sub TreeView1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS) '��ֹ��ѡ����
'''    Dim myNode As node
'''    For Each myNode In Me.TreeView1.Nodes
'''        If myNode.Children > 0 And myNode.Checked = True Then myNode.Checked = False
'''    Next
''End Sub
