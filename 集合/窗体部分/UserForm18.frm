VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm18 
   Caption         =   "二维码C"
   ClientHeight    =   3435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4515
   OleObjectBlob   =   "UserForm18.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    If Len(QRtextEN) = 0 Then Exit Sub
    With Me.BarCodeCtrl1
        .Style = 11      'https://docs.microsoft.com/ja-jp/previous-versions/office-development/cc427149(v=msdn.10)?redirectedfrom=MSDN
        .Validation = 2
        .Width = 156
        .Height = 156
        .Value = QRtextEN 'vba自带的空间不支持中文
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    QRtextEN = ""
End Sub
