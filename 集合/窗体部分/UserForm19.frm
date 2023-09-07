VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm19 
   Caption         =   "条形码"
   ClientHeight    =   2910
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7335
   OleObjectBlob   =   "UserForm19.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    If Len(Barcodex) = 0 Then Exit Sub
    With Me.BarCodeCtrl1
        .Style = 7      'https://docs.microsoft.com/ja-jp/previous-versions/office-development/cc427149(v=msdn.10)?redirectedfrom=MSDN '一些旧的的微软的资料很多都可以在日文站点找到
        .Validation = 0
        .Value = Barcodex
        .Height = 90
        .Width = 306
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Barcodex = ""
End Sub
