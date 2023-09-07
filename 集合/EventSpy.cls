VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Eventspy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents Napp As Excel.Application
Attribute Napp.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set Napp = Excel.Application
End Sub

Private Sub Class_Terminate()
    Set Napp = Nothing
End Sub

Private Sub Napp_WindowResize(ByVal wb As Workbook, ByVal Wn As Window) '任意的工作簿最小化时，窗口4显示,注意在锁定的状态下,在任务栏下依然能打开新的窗体
    Dim i As Byte, k As Byte
    Dim strx As String
    
    strx = ThisWorkbook.Name
    k = Workbooks.Count
    If k > 1 Then
        For i = 1 To k
            If Workbooks(i).Windows(1).WindowState = xlMinimized And Workbooks(i).Name <> strx Then UserForm4.Show
        Next
    End If
End Sub
