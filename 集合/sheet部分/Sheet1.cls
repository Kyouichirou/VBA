VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub CommandButton1_Click() '查询
'Set Rng = Sheet2.Range("b6:b" & Sheet2.[b65536].End(xlUp).Row).Find(Me.TextBox1.Value)
'If Rng Is Nothing Then
'MsgBox "未找到文件"
'End If
'With Sheet1
'     .Range("i14:i24").ClearContents
'     .Range("i14") = Me.TextBox1.Value
'     .Range("i19") = Rng.Offset(0, 1)
'End With
End Sub

Private Sub CommandButton10_Click()
ThisWorkbook.Sheets(3).Activate
End Sub

Private Sub CommandButton11_Click() '百度搜索
Call baidu
End Sub

Private Sub CommandButton12_Click() '豆瓣搜素
Call douban
End Sub

Private Sub CommandButton13_Click() 'wikipedia搜索
Call wikipedia
End Sub

Private Sub CommandButton14_Click() '打开控制板
ThisWorkbook.Windows(1).WindowState = xlMinimized
UserForm3.Show
End Sub

Private Sub CommandButton4_Click() '转到书库的位置
'Set Rng = Sheet2.Range("b6" & ":" & "b" & Sheet2.[b65536].End(xlUp).Row).Find(Me.TextBox1.Value)
'Sheet2.Select
'Rng.Select
End Sub

Private Sub CommandButton9_Click() '转到书库
ThisWorkbook.Sheets("书库").Select
End Sub

Private Sub Worksheet_Activate()
ThisWorkbook.Sheets("主界面").ScrollArea = "a1:z50"
End Sub

Private Sub ListBox1_Click() '点击赋值
'Me.TextBox1.Value = Me.ListBox1.Value
'Me.ListBox1.Clear
'Me.ListBox1.Visible = False
End Sub

Private Sub TextBox1_Change() '联想
'Dim arr()
'With ThisWorkbook              '避免打开其他的工作簿时出现崩溃的问题
'arr = .Sheets("书库").Range("b6:c" & .Sheets("书库").[c65536].End(xlUp).Row).Value
'
'If Len(Me.TextBox1.Value) >= 2 Then
'    Me.ListBox1.Clear
'    For i = 1 To .Sheets("书库").[c65536].End(xlUp).Row - 5
'    If InStr(arr(i, 1), Me.TextBox1.Value) > 0 Then
'       Me.ListBox1.AddItem
'       Me.ListBox1.List(Me.ListBox1.ListCount - 1, 0) = arr(i, 1) '分列显示
'       Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = arr(i, 2)
'    End If
'    If Me.ListBox1.ListCount > 0 Then
'    Me.ListBox1.Visible = True
'    Else
'    Me.ListBox1.Visible = False
'    End If
'    Next
'Else
'Me.ListBox1.Visible = False
'End If
'End With
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal target As Range, Cancel As Boolean)
'Dim targetcode As String
'Dim slr As Integer, slc As Integer
'
'slc = target.Column
'slr = target.Row
'
'If target.Locked Then Cancel = True '双击文件时不发出警告
'
'If slr > 26 And slr < 34 And slc > 3 And slc < 11 Then '限制双击的执行范围
'    targetcode = Sheets("主界面").Range("i" & slr).Value
'    ElseIf slr > 26 And slr < 34 And slc > 15 And slc < 23 Then
'    targetcode = Sheets("主界面").Range("u" & slr).Value
'Else
'    Exit Sub
'End If
'
'If FileExist(targetcode) = False Then
'MsgBox "文件已删除或处于打开的状态"
'Set Rng = Nothing
'Exit Sub
'End If
'
'Call OpenFile(targetcode, Rng.Offset(0, 1), Rng.Offset(0, 2), Rng.Offset(0, 3), 0)

End Sub
