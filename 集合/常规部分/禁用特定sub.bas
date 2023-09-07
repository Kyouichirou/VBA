Attribute VB_Name = "禁用特定sub"
Option Explicit
Sub PauseSub(ByRef SubName As String, ByRef ModulName As String, ByRef Commandx As String, ByVal Contnum As Integer)
    Dim i As Integer
    
    i = ThisWorkbook.VBProject.VBComponents(ModulName).CodeModule.ProcBodyLine(SubName, vbext_pk_Proc)
    If Err.Number = 35 Then
        Err.Clear
        MsgBox "未找到sub"
        Exit Sub
    Else
        If Commandx = "Disable" And Contnum = 0 Then
            ThisWorkbook.VBProject.VBComponents(ModulName).CodeModule.InsertLines i + 1, "Exit Sub"
        ElseIf Commandx = "Disable" And Contnum = 1 Then
            ThisWorkbook.VBProject.VBComponents(ModulName).CodeModule.InsertLines i + 2, "Exit Sub"
        ElseIf Commandx = "Enable" And Contnum = 0 Then
            ThisWorkbook.VBProject.VBComponents(ModulName).CodeModule.DeleteLines i + 1, 1
        ElseIf Commandx = "Enable" And Contnum = 1 Then
            ThisWorkbook.VBProject.VBComponents(ModulName).CodeModule.DeleteLines i + 2, 1
        End If
    End If
End Sub

Sub DisColorHigh() '颜色高亮
    If ThisWorkbook.VBProject.VBComponents("sheet2").CodeModule.ProcCountLines("Colorhigh", vbext_pk_Proc) = 7 Then Call PauseSub("Colorhigh", "其他", "Disable", 0) '代码的行数的计算从横线往下一直算,并不是sub和end sub之间的行数(包括换行符形成的空行)
End Sub

Sub EnColorHigh()
    With ThisWorkbook.VBProject.VBComponents("sheet2").CodeModule
       If .ProcCountLines("Worksheet_SelectionChange", vbext_pk_Proc) = 17 Then Call EnableRm '如果处于禁用状态,就启用
       If .ProcCountLines("Colorhigh", vbext_pk_Proc) = 8 Then Call PauseSub("Colorhigh", "其他", "Enable", 0)
    End With
End Sub

Sub PauseRm() '禁用整个select事件,由于selection事件非常容易触发
    If ThisWorkbook.VBProject.VBComponents("sheet2").CodeModule.ProcCountLines("Worksheet_SelectionChange", vbext_pk_Proc) = 17 Then Call PauseSub("Worksheet_SelectionChange", "sheet2", "Disable", 0) '禁用sheet2(书库)表的selection事件,在很多的执行当中很多都涉及这个操作
End Sub

Sub EnableRm() '启用事件
    If ThisWorkbook.VBProject.VBComponents("sheet2").CodeModule.ProcCountLines("Worksheet_SelectionChange", vbext_pk_Proc) = 18 Then Call PauseSub("Worksheet_SelectionChange", "sheet2", "Enable", 0) '启用
End Sub

Sub ColorHigh() '高亮提示
    Exit Sub
    If Selection.Row >= 6 And Selection.Column > 1 Then '限制执行的区域
        Cells.Interior.Pattern = xlPatternNone '将表格的样式设置为空
        Selection.EntireRow.Interior.Color = 65535 '改变选择的列的颜色
    End If
End Sub
