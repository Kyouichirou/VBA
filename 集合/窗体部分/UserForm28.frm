VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm28 
   Caption         =   "UserForm28"
   ClientHeight    =   8970
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16830
   OleObjectBlob   =   "UserForm28.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm28"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Byte

Private Sub CommandButton1_Click()
Dim url As String
    url = "https://www.qcc.com/search?key="
    url = url & Application.EncodeUrl(Cells(i, "d").Value)
    i = i + 1
    Webbrowserx url
End Sub

Private Function Webbrowserx(ByVal url As String) '执行浏览器动作 'webbrowser - https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa752043%28v%3dvs.85%29
    With Me.WebBrowser1
        .MenuBar = False
        .Silent = True
        .Navigate (url)
    End With
End Function

Private Sub CommandButton2_Click() '搜索选中
    Dim selectobj As Object, rngobj As Object, strx As String
    
    Set selectobj = Me.WebBrowser1.Document.Selection
    If selectobj Is Nothing Then Exit Sub
    Set rngobj = selectobj.createrange
    If rngobj Is Nothing Then Exit Sub
    strx = rngobj.HTMLText
    Cells(i - 1, "e") = strx
    Set selectobj = Nothing
    Set rngobj = Nothing
End Sub

Private Sub UserForm_Initialize()
i = 18
Webbrowserx "https://www.qcc.com/"
End Sub

Private Sub UserForm_Terminate()
i = 0
End Sub


