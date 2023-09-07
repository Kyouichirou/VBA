VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "Mini"
   ClientHeight    =   495
   ClientLeft      =   26115
   ClientTop       =   14970
   ClientWidth     =   2835
   OleObjectBlob   =   "UserForm4.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CommandButton1_Click() '返回主窗口
    Me.Hide
    If Workbooks.Count > 1 Then        '1/ 0 代表不同的交互模式showmodal, 0 标识可以和excel进行交互
        UserForm3.Show 0
    Else
        UserForm3.Show 1
    End If
End Sub

Private Sub CommandButton2_Click() '解除隐藏状态
    Call Rewds
End Sub

Private Sub CommandButton3_Click()
    UserForm17.Show
End Sub

Private Sub UserForm_Activate() '窗口出现的位置
    Dim i As Integer, k As Integer
    If Statisticsx = 1 Then Exit Sub
    i = Int(ActiveWindow.UsableHeight)
    k = Int(ActiveWindow.UsableWidth)
    With Me
        .StartUpPosition = 0
'        .Left = 1300           '根据屏幕的不同进行调整
'        .Top = 729
        '.Left = 998
        '.Top = 572
        .Left = k - 138
        .Top = i + 130
    End With
End Sub

Private Sub UserForm_Initialize()
    If Statisticsx = 1 Then Exit Sub
    UF4Show = 1
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Statisticsx = 1 Then Exit Sub
    If Workbooks.Count = 1 And Application.Visible = False Then '当进入精简模式的时候,禁止窗口被关闭
        If CloseMode = vbFormControlMenu Then Cancel = True
    End If
End Sub

Private Sub UserForm_Terminate()
    If Statisticsx = 1 Then Exit Sub
    UF4Show = 0
End Sub
