VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "DataClean"
   ClientHeight    =   2490
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6090
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CommandButton1_Click() '�������
    Call ClearAll
    ThisWorkbook.Save
    Unload Me     '��Ҫ��ж�ش���,����u3�޷�ж��
    If UF3Show > 0 Then Unload UserForm3 'u3���巢������
    If UF4Show > 0 Then Unload UserForm4
End Sub

Private Sub CommandButton2_Click() '�رմ���
    Unload Me
End Sub
