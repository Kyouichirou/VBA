VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm24 
   Caption         =   "�������"
   ClientHeight    =   6135
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   4725
   OleObjectBlob   =   "UserForm24.frx":0000
   StartUpPosition =   1  '����������
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
        If LCase$(Right$(Imgurl, 3)) <> "jpg" Then MsgShow "ͼƬ������Ҫ��", "Tips", 1200: Unload Me '����ʹ��ͼƬ��ʽ ת��,��ͼƬת��image�ؼ�֧�ֵĸ�ʽ
        ReDim arr(1)
        With Me.Image1
            arr = ObtainPicWH(Imgurl)
            pw = arr(0)
            i = 1
            If pw > 700 Then i = 2 '��ֹͼƬ����
            If pw <> 297 Then 'x1.5 '�Ӷ�����ץȡ��ͼƬ�Ĵ�С��һ,webbrowser/IEץȡ��ͼƬ��һ��,webbrowser��С,IE�ϴ�
                pw = pw / 1.5
                ph = arr(1) / 1.5
                With Me    '������ߴ����
                    .Height = ph * 1.12 / i
                    .Width = pw * 1.13 / i
                End With
                .Width = pw / i 'ͼƬ�ؼ��ߴ����
                .Height = ph / i
            End If
            .Picture = LoadPicture(Imgurl)
            .PictureSizeMode = fmPictureSizeModeStretch
        End With
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) '��Ҫ�޸�
    If Statisticsx = 1 Then Exit Sub
    If UF3Show > 0 Then UserForm3.TextBox1.SetFocus
End Sub
