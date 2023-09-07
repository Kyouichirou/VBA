VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm27 
   Caption         =   "音乐"
   ClientHeight    =   7740
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15840
   OleObjectBlob   =   "UserForm27.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton2_Click()
 Me.CommandButton3.Enabled = True
 Me.CommandButton2.Enabled = False
End Sub

Private Sub CommandButton3_Click()
Me.CommandButton3.Enabled = False
Me.CommandButton2.Enabled = True
End Sub
