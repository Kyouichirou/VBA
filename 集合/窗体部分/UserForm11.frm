VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm11 
   Caption         =   "��ʼ�����"
   ClientHeight    =   3915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9075
   OleObjectBlob   =   "UserForm11.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "UserForm11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Dim strx As String
    strx = ThisWorkbook.Sheets("temp").Range("ab2").Value
    Call OpenFileLocation(strx) '�򿪳������ڵ���λ��
    ThisWorkbook.Close savechanges:=True
End Sub

Private Sub UserForm_Initialize()
    Dim arr() As Variant, i As Byte
    
    If Statisticsx = 1 Then Exit Sub
    With ThisWorkbook.Sheets("temp")
        i = .[aa65536].End(xlUp).Row
        arr = .Range("ab1:ac" & i).Value
    End With
    With Me
        .Label10.Caption = arr(23, 1)
        .Label11.Caption = arr(24, 1)
        
        If Len(arr(4, 1)) > 0 Then      'Powershell
            If Len(arr(4, 2)) > 0 Then
                .Label12.Caption = "֧��"
            Else
                .Label12.Caption = "�汾̫��"
            End If
        Else
            .Label12.Caption = "��֧��"
        End If
        
        If Len(arr(5, 1)) > 0 Then
            .Label13.Caption = "֧��"
        Else
            .Label13.Caption = "��֧��"
        End If
        
        If Len(arr(6, 1)) > 0 Then         'IE
            If Len(arr(6, 2)) > 0 Then
                .Label14.Caption = "֧��"
            Else
                .Label14.Caption = "�汾̫��"
            End If
        Else
            .Label14.Caption = "��֧��"
        End If
        
        If Len(arr(8, 1)) > 0 Then
            .Label15.Caption = "����"
        Else
            .Label15.Caption = "������"
        End If
        
        If Len(arr(7, 1)) > 0 Then
            .Label16.Caption = "����"
        Else
            .Label16.Caption = "������"
        End If
        
        If Len(arr(17, 1)) > 0 Then
            .Label17.Caption = "����"
        Else
            .Label17.Caption = "������"
        End If
        
        If Len(arr(25, 1)) > 0 Then
            .Label18.Caption = "����"
        Else
            .Label18.Caption = "������"
        End If
        
        If Len(arr(32, 1)) > 0 Then 'zip
            .Label21.Caption = "֧��"
        Else
            .Label21.Caption = "��֧��"
        End If
        
        If Len(arr(10, 1)) > 0 Then 'chrome
            .Label22.Caption = "֧��"
        Else
            .Label22.Caption = "��֧��"
        End If
        
        .TextBox1.Text = "1. ����������ڽ���ѧϰʹ��,����������ҵ��;." & vbCr & "2. �����򲻰����κζ������." & vbCr & _
        "3. ������ȷ���û���ʹ�ù�������ɵ�������ʧ,���ܳ��򾭹��ϸ����,���޷���֤�����Bug���û������Σ��(��Md5�㷨��õ�Hashֵ,�������ڱȽ��ļ�,��ͬ���ļ�����ɾ��), ʹ��ǰ����ϸ��������." & vbCr _
        & "4. �������õĴ������Դ̫���ڹ㷺,�޷�һһ��עԭ����, �ٴ˶����еĿ�Դ��������߱�ʾ���и�л." & vbCr _
        & "5. ת�ػ����Ƕ����޸�,ϣ���ܱ�������."
    End With
    
    Erase arr
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Exit Sub
    Cancel = False
End Sub
