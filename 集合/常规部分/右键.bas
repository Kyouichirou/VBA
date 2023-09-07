Attribute VB_Name = "�Ҽ�"
Option Explicit

Sub MyNewMenu()              '�����˵�
Dim mymenu As Object
Dim CM1 As Variant, CM2 As Variant, CM3 As Variant, CM4 As Variant, CM5 As Variant, CM6 As Variant, CM7 As Variant, CM8 As Variant

    With ThisWorkbook.Application.CommandBars("cell")
        .Reset
        Set mymenu = .Controls.Add(type:=msoControlPopup, Before:=1)
        .Controls(2).BeginGroup = True
    End With
    
    With mymenu
        .Caption = "�ļ�����"          '������Ҫ�ɸı���ʾ����
        Set CM1 = .Controls.Add(type:=msoControlButton)
        Set CM2 = .Controls.Add(type:=msoControlButton)
        Set CM3 = .Controls.Add(type:=msoControlButton)
        Set CM4 = .Controls.Add(type:=msoControlButton)
        Set CM5 = .Controls.Add(type:=msoControlButton)
        Set CM6 = .Controls.Add(type:=msoControlButton)
        Set CM7 = .Controls.Add(type:=msoControlButton)
        Set CM8 = .Controls.Add(type:=msoControlButton)
    End With
    
    With CM1
        .Caption = "����"        '������Ҫ�ɸı���ʾ����
        .OnAction = "Rfilecopy"          '������Ҫ�ɸı�ִ�к�
        .FaceId = 266             '������Ҫ�ɸı���ʾͼ��
    End With
    
    With CM2
        .Caption = "�ƶ�"
        .OnAction = "RfileMove"
        .FaceId = 49
    End With
    
     With CM3
        .Caption = "������"
        .OnAction = "Filerename"
        .FaceId = 59 'ͼ��
    End With
    
    With CM4
        .Caption = "ɾ���ļ�"
        .OnAction = "Rdelefile"
        .FaceId = 79
    End With
    
    With CM5
        .Caption = "�������ļ���"
        .OnAction = "Ropenfilelocation"
        .FaceId = 89
    End With
    
     With CM6
        .Caption = "�ļ�����"
        .OnAction = "Filedata"
        .FaceId = 89
    End With
    
    With CM7
        .Caption = "��ӵ������Ķ�"
        .OnAction = "Raddplist"
        .FaceId = 89
    End With
    
    With CM8
        .Caption = "���ļ�"
        .OnAction = "Openfiletemp"
        .FaceId = 89
    End With
    
    Set mymenu = Nothing
End Sub

Sub ResetMenu()                      '���ò˵�
    ThisWorkbook.Application.CommandBars("cell").Reset
End Sub

Sub Rfilecopy()  '�ļ�����
    Dim addrowx As Integer
    addrowx = Selection.Row
    With ThisWorkbook.Sheets("���")
        If FileCopy(.Range("e" & addrowx).Value, .Range("c" & addrowx).Value, addrowx) = True Then
            .Label1.Caption = "�ļ����Ƴɹ�"
        Else
            .Label1.Caption = "�ļ�����ʧ��"
        End If
    End With
End Sub

Sub RfileMove() '�ļ��ƶ�
    FileMove
End Sub

Sub RdeleFile()                  'ɾ���ļ�
    Dim addrowx As Integer, excodex As Byte, p As Byte
    Dim yesno As Variant
    
    With ThisWorkbook.Sheets("���")
        If Len(Selection) = 0 Then Exit Sub
        addrowx = Selection.Row
        yesno = MsgBox("�Ƿ�ɾ�������ļ�?_", vbYesNo) '����޷��������Ӵ洢�ļ��Ĵ���
        If yesno = vbNo Then DeleMenu (addrowx): Exit Sub
        If Len(.Cells(addrowx, "ab").Value) > 0 Then p = 1 '����Ƿ���ڷ�ansi
        Call FileDeleExc(.Cells(addrowx, "e").Value, addrowx, p, 0, Cells(addrowx, "c").Value)
    End With
End Sub

Sub FileData() '�ļ�����
    Dim addrowx As Integer, strx As String, strx1 As String
    ThisWorkbook.Application.ScreenUpdating = False
    If Len(Selection) = 0 Then Exit Sub
    addrowx = Selection.Row
    With ThisWorkbook.Sheets("���")
'    If fso.FileExists(.Cells(addrowx, "e").Value) = False Then
'        DeleMenu addrowx
'        MsgBox "�ļ�������", vbCritical, "Warning"
'        Exit Sub
'    End If
        strx = .Cells(addrowx, 2)
        Set Rng = .Range("b" & addrowx)
    End With
    If UF3Show = 0 Then
        Load UserForm3
        Call ShowDetail(strx)
    ElseIf UF3Show = 1 Or UF3Show = 3 Then
        strx1 = UserForm3.Label29.Caption
        If UF4Show = 1 Then Unload UserForm4
        If Len(strx1) > 0 Then
            If strx1 <> strx Then Call ShowDetail(strx)
        End If
    End If
    ThisWorkbook.Application.ScreenUpdating = True
    UserForm3.Show
End Sub

Sub FileReName()                     '�ļ�������
    If Len(Selection) = 0 Then Exit Sub
    If UF3Show = 1 Then Unload UserForm3
    If UF3Show = 3 Then Unload UserForm4
    UserForm9.Show
End Sub

Sub RopenFileLocation()  '���ļ�����λ��
    Call OpenFileLocation(Range("f" & Selection.Row))
End Sub

Sub RAddPlist() '��ӵ������Ķ�
    Dim addrowx As Integer, strx As String
    
    addrowx = Selection.Row
    With ThisWorkbook.Sheets("���")
        strx = .Range("b" & addrowx).Value
        If Len(strx) = 0 Then Exit Sub
        Call AddPList(strx, .Range("c" & addrowx).Value, 0)
    End With
End Sub

Sub OpenFileTemp() '���ļ�-������д��ļ�
    Dim slr As Integer, i As Byte, strx As String
    Dim addressx As String
    Dim filext As String
    
    With ThisWorkbook.Sheets("���")
        If Len(Selection) = 0 Then Exit Sub '�޸�
        slr = Selection.Row
        If slr < 6 Then Exit Sub
        addressx = .Range("e" & slr).Value '·��
        filext = LCase(.Range("d" & slr).Value) '��չ��
        If filext Like "xl*" Then
            If filext <> "xls" And filext <> "xlsx" Then MsgBox "��ֹ�򿪴����ļ�", vbCritical, "Warning": Exit Sub
        End If
        If Len(.Cells(slr, "ak").Value) = 0 Then
            i = FileTest(addressx, filext, .Cells(slr, "c").Value)
            Select Case i
                Case 1: strx = "δ��ȡ����Чֵ"
                Case 2: strx = "�ļ�������": Call DeleMenu(slr) 'ɾ�����Ŀ¼
                Case 3: strx = "�ļ���txt�ļ�"
                Case 4: strx = "�ļ����ڴ򿪵�״̬"
                Case 5: strx = "�ļ����ڴ򿪵�״̬"
                Case 7: strx = "�����쳣"
            End Select
        Else
            If Not filext Like "xls*" Then
                If WmiCheckFileOpen(addressx) = True Then .Label1.Caption = "�ļ��Ѵ�": Exit Sub
            End If
        End If
        If i = 0 Or i = 3 Or i = 6 Then
            If i = 6 Then .Cells(slr, "ak") = 1 '����ļ��������뱣��
            Set Rng = .Range("b" & slr)
            OpenFile .Range("b" & slr).Value, filext, .Range("d" & slr).Value, addressx, 2, .Range("ab" & slr).Value: Exit Sub
        Else
            .Label1.Caption = strx
        End If
    End With
End Sub

Function CheckRightMenu() As Boolean
    If ActiveWorkbook.Name <> ThisWorkbook.Name Then
        CheckRightMenu = False
        Exit Function
    Else
        If ThisWorkbook.ActiveSheet.Name <> "���" Then
            CheckRightMenu = False
            Exit Function
        Else
            If Selection.Row < 6 Or Selection.Column < 1 Then
                CheckRightMenu = False
            Else
                CheckRightMenu = True
            End If
        End If
    End If
End Function
