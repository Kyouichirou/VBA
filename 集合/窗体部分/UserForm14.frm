VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm14 
   Caption         =   "��ά��-Online"
   ClientHeight    =   3435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4515
   OleObjectBlob   =   "UserForm14.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "UserForm14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize() '���ɶ������ӵĶ�ά��
    Dim Urlx As String, strx As String, Filenamex As String, filecodex As String
    
    If Statisticsx = 1 Then Exit Sub
    With UserForm3
        Urlx = .Label106.Caption
        filecodex = .Label29.Caption
    End With
    If Len(Urlx) = 0 Then Exit Sub
    strx = Urlx
    strx = encodeURI(strx) 'ת�� '��ͬ��ά������վ��Ĳ�����һ��, ��һЩʹ�õ�md5 hashֵ��Ϊ����
    Urlx = "https://tool.oschina.net/action/qrcode/generate?data=" & strx & "&output=image%2Fgif&error=L&type=0&margin=0&size=4.jpeg" 'ע�������"gif"����
    ''---------------------------��������png��ʽ���ļ�, ��Ϊloadpicture��֧��png��ʽ
    With WebBrowser1
        .Width = 156
        .Height = 156
        .Navigate (Urlx)
    End With
    Filenamex = ThisWorkbook.Sheets("temp").Cells(44, "ab").Value & "\" & filecodex & ".png" '���ض�ά��ͼƬ
    If DownloadFilex(Urlx, Filenamex) = False Then Exit Sub
    SearchFile filecodex
    If Rng Is Nothing Then Exit Sub
    Rng.Offset(0, 33) = Filenamex '��ά���λ��
    Set Rng = Nothing
End Sub
