Attribute VB_Name = "�����ض�sub"
Option Explicit
Sub PauseSub(ByRef SubName As String, ByRef ModulName As String, ByRef Commandx As String, ByVal Contnum As Integer)
    Dim i As Integer
    
    i = ThisWorkbook.VBProject.VBComponents(ModulName).CodeModule.ProcBodyLine(SubName, vbext_pk_Proc)
    If Err.Number = 35 Then
        Err.Clear
        MsgBox "δ�ҵ�sub"
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

Sub DisColorHigh() '��ɫ����
    If ThisWorkbook.VBProject.VBComponents("sheet2").CodeModule.ProcCountLines("Colorhigh", vbext_pk_Proc) = 7 Then Call PauseSub("Colorhigh", "����", "Disable", 0) '����������ļ���Ӻ�������һֱ��,������sub��end sub֮�������(�������з��γɵĿ���)
End Sub

Sub EnColorHigh()
    With ThisWorkbook.VBProject.VBComponents("sheet2").CodeModule
       If .ProcCountLines("Worksheet_SelectionChange", vbext_pk_Proc) = 17 Then Call EnableRm '������ڽ���״̬,������
       If .ProcCountLines("Colorhigh", vbext_pk_Proc) = 8 Then Call PauseSub("Colorhigh", "����", "Enable", 0)
    End With
End Sub

Sub PauseRm() '��������select�¼�,����selection�¼��ǳ����״���
    If ThisWorkbook.VBProject.VBComponents("sheet2").CodeModule.ProcCountLines("Worksheet_SelectionChange", vbext_pk_Proc) = 17 Then Call PauseSub("Worksheet_SelectionChange", "sheet2", "Disable", 0) '����sheet2(���)���selection�¼�,�ںܶ��ִ�е��кܶ඼�漰�������
End Sub

Sub EnableRm() '�����¼�
    If ThisWorkbook.VBProject.VBComponents("sheet2").CodeModule.ProcCountLines("Worksheet_SelectionChange", vbext_pk_Proc) = 18 Then Call PauseSub("Worksheet_SelectionChange", "sheet2", "Enable", 0) '����
End Sub

Sub ColorHigh() '������ʾ
    Exit Sub
    If Selection.Row >= 6 And Selection.Column > 1 Then '����ִ�е�����
        Cells.Interior.Pattern = xlPatternNone '��������ʽ����Ϊ��
        Selection.EntireRow.Interior.Color = 65535 '�ı�ѡ����е���ɫ
    End If
End Sub
