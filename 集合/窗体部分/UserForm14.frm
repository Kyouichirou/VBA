VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm14 
   Caption         =   "二维码-Online"
   ClientHeight    =   3435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4515
   OleObjectBlob   =   "UserForm14.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize() '生成豆瓣链接的二维码
    Dim Urlx As String, strx As String, Filenamex As String, filecodex As String
    
    If Statisticsx = 1 Then Exit Sub
    With UserForm3
        Urlx = .Label106.Caption
        filecodex = .Label29.Caption
    End With
    If Len(Urlx) = 0 Then Exit Sub
    strx = Urlx
    strx = encodeURI(strx) '转码 '不同二维码生成站点的参数不一样, 有一些使用的md5 hash值作为参数
    Urlx = "https://tool.oschina.net/action/qrcode/generate?data=" & strx & "&output=image%2Fgif&error=L&type=0&margin=0&size=4.jpeg" '注意这里的"gif"参数
    ''---------------------------不能下载png格式的文件, 因为loadpicture不支持png格式
    With WebBrowser1
        .Width = 156
        .Height = 156
        .Navigate (Urlx)
    End With
    Filenamex = ThisWorkbook.Sheets("temp").Cells(44, "ab").Value & "\" & filecodex & ".png" '下载二维码图片
    If DownloadFilex(Urlx, Filenamex) = False Then Exit Sub
    SearchFile filecodex
    If Rng Is Nothing Then Exit Sub
    Rng.Offset(0, 33) = Filenamex '二维码的位置
    Set Rng = Nothing
End Sub
