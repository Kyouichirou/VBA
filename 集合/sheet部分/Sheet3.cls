VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long '判断数组是否完成初始化
Dim arra() As Variant
Dim arrB() As Variant '表格的模块级别的变量在发生表格切换时会被自动重置
Dim blow As Integer
Private Sub CommandButton1_Click()
    Sheet3.Range("a1").Activate
    ThisWorkbook.Sheets("书库").Activate
End Sub

Private Sub CommandButton2_Click()
    Sheet3.Range("a1").Activate
    ThisWorkbook.Sheets("主界面").Activate
End Sub

Private Sub CommandButton3_Click()
    Sheet3.Range("a1").Activate
    Application.ActiveWindow.WindowState = xlMinimized
    UserForm3.Show
End Sub

Private Sub CommandButton4_Click() '精简模式
    Sheet3.Range("a1").Activate
    UserForm3.Hide
    If Workbooks.Count > 1 Then
        Application.ActiveWindow.WindowState = xlMinimized
    Else
        Application.Visible = False
        With UserForm4
            .Hide
            .Caption = "锁定"
            .Show 1
        End With
    End If
End Sub

Private Sub CommandButton5_Click()
    Sheet3.Range("a1").Activate
    ThisWorkbook.Sheets("temp").Activate
End Sub

Private Sub CommandButton6_Click()
    Sheet3.Range("a1").Activate
    ThisWorkbook.Sheets("目录").Activate
End Sub

Private Sub CommandButton7_Click() '详情
    Dim strx As String, strx1 As String
    Dim i As Byte, k As Byte
    
    With Me.ListBox1
        i = .ListCount
        If i = 0 Then Exit Sub
        i = i - 1
        For k = 0 To i
            If .Selected(k) = True Then strx = .List(k, 0): Exit For
        Next
    End With
    Sheet3.Range("a1").Activate
    SearchFile strx
    If Rng Is Nothing Then MsgBox "文件不存在", vbCritical, "Warning": Exit Sub
    ThisWorkbook.Application.ScreenUpdating = False
    If UF3Show = 0 Then
        Load UserForm3
        Call ShowDetail(strx)
    ElseIf UF3Show = 1 Or UF3Show = 3 Then
        strx1 = UserForm3.Label29.Caption
        If UF4Show = 1 Then Unload UserForm4
        If Len(strx1) > 0 Then
            If strx1 <> strx Then Call ShowDetail(strx)
        End If
    End If
    Set Rng = Nothing
    ThisWorkbook.Application.ScreenUpdating = True
    UserForm3.Show
End Sub

Private Sub CommandButton8_Click() 'vbe
    Sheet3.Range("a1").Activate
    ThisWorkbook.Application.SendKeys ("%{F11}")
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Integer, strx As String
    
    With Me.ListBox1
        If .ListCount = 0 Then Exit Sub
        i = .ListIndex
        strx = .List(i, 0)
        SearchFile strx
        If Rng Is Nothing Then MsgBox "文件不存在", vbCritical, "Warning": Exit Sub
    End With
End Sub

Private Sub TextBox1_Change() '搜索框
    Dim i As Integer, j As Integer, k As Integer
    Dim strx As String, mi As Byte 'ltb As Object

    If SafeArrayGetDim(arra) = 0 Then ArrayLoadx
    If blow < 8 Then
        Me.TextBox1.Text = "数据库尚未存储数据"
        Exit Sub
    End If
    strx = Me.TextBox1.Value
    strx = Replace(strx, "/", " ") '替换掉"/"符号
    With Me.ListBox1
        If Len(strx) >= 2 Then
            .Clear
            blow = blow - 5
            mi = 0
            For k = 1 To blow
                If InStr(1, arrB(k, 1) & "/" & arra(k, 1), strx, vbTextCompare) > 0 Then
'                    If mi = 0 Then Set ltb = Me.OLEObjects.Add(ClassType:="Forms.ListBox.1", Link:=False, DisplayAsIcon:=False, Left:=356.4, Top:=154.2, Width:=429, Height:=92)
                    If mi = 0 Then .Height = 96 '注意这里的高度只能设置一次否则会出现,无法看到添加的列表值的情况
                    .AddItem
                    .List(mi, 0) = arra(k, 1)
                    .List(mi, 1) = arra(k, 2)
                    mi = mi + 1
                    If mi > 10 Then Exit For
                Else
                   If mi = 0 Then .Clear: .Height = 1 '这里需要注意, 如果直接使用visible属性,会造成屏幕刷新的问题,无法通过application.screen来解决
                End If                                '这里使用控制listbox的宽度来假装处于隐藏的状态
            Next
        Else
            .Clear: .Height = 1
        End If
    End With
End Sub

Private Sub Worksheet_Activate()
    With ThisWorkbook.Sheets("首页")
        .Range("a1").Activate '让按钮保持透明
        .ScrollArea = "a1:t36" '控制滑动的区域
    End With
End Sub

Private Sub ArrayLoadx()
    With ThisWorkbook.Sheets("书库") '注意这里的sheet被隐藏的问题
        blow = .[d65536].End(xlUp).Row
        arra = .Range("b6:c" & blow).Value
        arrB = .Range("e6:e" & blow).Value
    End With
End Sub
