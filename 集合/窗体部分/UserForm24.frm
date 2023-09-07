VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm24 
   Caption         =   "豆瓣封面"
   ClientHeight    =   6135
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   4725
   OleObjectBlob   =   "UserForm24.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim arr() As Long
    Dim pw As Integer, ph As Integer, i As Byte
    
    If Statisticsx = 1 Then Exit Sub
    If Len(Imgurl) > 0 Then
        If LCase$(Right$(Imgurl, 3)) <> "jpg" Then MsgShow "图片不符号要求", "Tips", 1200: Unload Me '或者使用图片格式 转换,将图片转成image控件支持的格式
        ReDim arr(1)
        With Me.Image1
            arr = ObtainPicWH(Imgurl)
            pw = arr(0)
            i = 1
            If pw > 700 Then i = 2 '防止图片过大
            If pw <> 297 Then 'x1.5 '从豆瓣上抓取的图片的大小不一,webbrowser/IE抓取的图片不一样,webbrowser较小,IE较大
                pw = pw / 1.5
                ph = arr(1) / 1.5
                With Me    '主窗体尺寸调整
                    .Height = ph * 1.12 / i
                    .Width = pw * 1.13 / i
                End With
                .Width = pw / i '图片控件尺寸调整
                .Height = ph / i
            End If
            .Picture = LoadPicture(Imgurl)
            .PictureSizeMode = fmPictureSizeModeStretch
        End With
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) '需要修改
    If Statisticsx = 1 Then Exit Sub
    If UF3Show > 0 Then UserForm3.TextBox1.SetFocus
End Sub
