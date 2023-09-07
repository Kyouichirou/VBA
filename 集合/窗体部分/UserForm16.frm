VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm16 
   Caption         =   "二维码-Offline"
   ClientHeight    =   3435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4515
   OleObjectBlob   =   "UserForm16.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    If Statisticsx = 1 Then Exit Sub
    If Len(QRfilepath) = 0 Then Exit Sub
    With Me.Image1 '----------------------https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/image-control
        .Picture = LoadPicture(QRfilepath)
        .PictureAlignment = fmPictureAlignmentCenter '居中
    End With
End Sub
