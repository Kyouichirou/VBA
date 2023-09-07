VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub CommandButton1_Click() '启用高亮提示
    Call EnColorHigh
    CommandButton2.Enabled = True '启用与禁用为互斥功能
    CommandButton1.Enabled = False '功能需要相互唤醒
End Sub

Private Sub CommandButton2_Click() '禁用高亮
    Call DisColorHigh
    Selection.EntireRow.Interior.Pattern = xlPatternNone '禁用前将颜色去掉
    Me.CommandButton1.Enabled = True
    Me.CommandButton2.Enabled = False
End Sub

Private Sub CommandButton10_Click() '清理数据
    UserForm2.Show
End Sub

Private Sub CommandButton11_Click() '启用右键菜单 '防止程序崩溃造成右键菜单无法启用
 Call EnableRm
End Sub

Private Sub CommandButton12_Click() '返回主界面
    ThisWorkbook.Sheets("主界面").Activate
End Sub

Private Sub CommandButton13_Click() '控制板
    'ThisWorkbook.Windows(1).WindowState = xlMinimized
    UserForm3.Show
End Sub

Private Sub CommandButton14_Click() '首页
    ThisWorkbook.Sheets(3).Activate
End Sub

Private Sub CommandButton15_Click() '顶部
    Application.GoTo Cells(7, Selection.Column) '这个执行后将出现无法使用键盘方向键控制光标移动,需要用鼠标点击后重新激活光标
    Cells(6, Selection.Column).Select
End Sub

Private Sub CommandButton16_Click() '底部
    Dim slr As Integer
    
    slr = [c65536].End(xlUp).Row
    If slr = 6 Then Exit Sub
    Application.GoTo Cells(slr, Selection.Column)
    Cells(slr + 1, Selection.Column).Select
End Sub

Private Sub CommandButton17_Click() '右侧
    Application.GoTo Cells(Selection.Row, 29)
    Cells(Selection.Row, 30).Select
End Sub

Private Sub CommandButton18_Click() '左侧
    Application.GoTo Cells(Selection.Row, 5)
    Cells(Selection.Row, 4).Select
End Sub

Private Sub CommandButton19_Click() '上翻
    Dim slr As Integer, slrx As Integer
    
    slr = Selection.Row
    slrx = slr - 25
    If slrx < 5 Then Exit Sub
    Application.GoTo Cells(slrx, 3)
    Cells(slrx + 1, 3).Select
End Sub

Private Sub CommandButton21_Click() '调试模式
    Dim strx As String
    
    With Me.CommandButton21
        strx = .Caption
        If strx = "调试模式" Then
            ThisWorkbook.Application.EnableEvents = False
            .Caption = "退出调试"
            .ForeColor = &HFF& '红色
            With Me
                .CommandButton1.Enabled = False
                .CommandButton11.Enabled = False
            End With
        Else
            ThisWorkbook.Application.EnableEvents = True
            .Caption = "调试模式"
            .ForeColor = &H80000012 '恢复
            With Me
                .CommandButton1.Enabled = True
                .CommandButton11.Enabled = True
            End With
        End If
    End With
End Sub

Private Sub CommandButton9_Click() '下翻
    Dim slr As Integer, slrx As Integer
    
    slr = Selection.Row
    slrx = slr + 30
    If slrx < 31 Then Exit Sub
    Application.GoTo Cells(slrx, 3)
    Cells(slrx - 1, 3).Select
End Sub

Private Sub CommandButton3_Click() '添加文件
    Call AddFile
End Sub

Private Sub CommandButton4_Click() '添加文件夹
    If ListAllFiles(0, "NU") = True Then Me.Label1.Caption = "添加成功"
End Sub

Private Sub CommandButton5_Click() '目录
    ThisWorkbook.Sheets("目录").Activate
End Sub

Private Sub CommandButton6_Click() '检查文件存在
    Call CheckAllFile
End Sub

Private Sub CommandButton7_Click() '批量删除文件
    Call MFilesDele
End Sub

Private Sub CommandButton8_Click() '打开VBE
    ThisWorkbook.Application.SendKeys ("%{F11}") '组合快捷键 alt+f11
End Sub

Private Sub Worksheet_Activate() '启用selection事件
    Call EnableRm
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal target As Range, Cancel As Boolean) '双击打开文件
    Dim i As Byte
    
    If target.Locked Then Cancel = True '双击文件时不发出警告
    i = target.Column
    If i = 2 Or i = 3 Or i = 5 Then
        Call OpenFileTemp
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal target As Range) '保证没有exit sub的情况下维持17行代码 '代码的行数的计算从横线往下一直算,并不是sub和end sub之间的行数(包括换行符形成的空行)
        With ThisWorkbook.Sheets("temp")
        If target.Row < 6 Or target.Column = 1 Then '限定右键菜单的使用范围 'Selection
            If LenB(.Range("ab9").Value) = 0 Or .Range("ab9").Value <> 0 Then '重置菜单
                Call ResetMenu
                .Range("ab9") = 0
            End If
        Else                                                                        '0/1用于控制右键菜单的创建或者重置，避免不必要的重复动作 '必须要有这个参数,否则很容易触发这个事件
            If LenB(.Range("ab9").Value) = 0 Or .Range("ab9").Value <> 1 Then '创建菜单
                Call MyNewMenu
                .Range("ab9") = 1
            End If
        End If
    End With
    Call ColorHigh '高亮提示
End Sub
