VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm26 
   Caption         =   "进度"
   ClientHeight    =   2850
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11145
   OleObjectBlob   =   "UserForm26.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Byte



Private Sub CommandButton1_Click()
i = i + 1
Me.pgbDisplay.Value = i
End Sub


Private Sub UserForm_Initialize()
Me.pgbDisplay.Max = 10
Me.pgbDisplay.Min = 0
End Sub
