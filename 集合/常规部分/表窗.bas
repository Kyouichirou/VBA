Attribute VB_Name = "��"
Option Explicit
Public Text3Ch As String, Text4Ch As String '������ļ�/�����Ƿ񾭹��޸�
Dim imgx As Byte
Public Fileptc As Byte '�ļ�����״̬
Public Imgurl As String '����ͼƬ��·��
Public Filenamei As String, Filepathi As String, Folderpathi As String '������ʱ�洢������ansi�ַ����ַ���

Sub DisEvents() '���ø�����
    '------------------------ http://www.360doc.com/content/15/0401/06/7835172_459703611.shtml
    '------------------------ https://docs.microsoft.com/zh-cn/office/vba/api/excel.application.interactive
    With ThisWorkbook.Application
        .ScreenUpdating = False '��ֹ��Ļˢ��
        .EnableEvents = False '�����¼�
        .Calculation = xlCalculationManual '�����Զ�����
        .Interactive = False '��ֹ����(��ִ�к�ʱ,����ڱ���������ݻ���ɺ���ֹ)
    End With
End Sub

Sub EnEvents() '����
    With ThisWorkbook
        With .Application
            .ScreenUpdating = True
            .EnableEvents = True
            .Calculation = xlCalculationAutomatic
            .Interactive = True
        End With
'        .Save
    End With
End Sub

Function ClearAll(Optional ByVal cmCode As Byte) '������е�����
    With ThisWorkbook
        If cmCode = 0 Then
            .Sheets("���").Range("b6:ah10000").ClearContents
            .Sheets("Ŀ¼").Range("a4:q1000").ClearContents
            With .Sheets("������")
                .Range("e37:j100").ClearContents
                .Range("d27:l33").ClearContents
                .Range("p23:x33").ClearContents
            End With
        Else
            .Sheets("Ŀ¼").Range("a4:q1000").ClearContents
        End If
    End With
End Function

'----------------------------------------------------------��ʾ�������
Sub ShowDetail(ByVal filecodex As String) '��ʾ�������
    Dim strx As String
    With UserForm3
        If Len(.Label56.Caption) > 0 Then
            If filecodex <> .Label56.Caption Then .Label55.Visible = False 'ɾ�����ļ���ʾ
        End If
        .Label29.Caption = Rng.Offset(0, 0) 'ͳһ����
        Filenamei = Rng.Offset(0, 1) '�ļ���
        .Label23.Caption = Filenamei
        .Label24.Caption = Rng.Offset(0, 2) '�ļ�����
        Filepathi = Rng.Offset(0, 3) '�ļ�·��
        .Label25.Caption = Filepathi
        Folderpathi = Rng.Offset(0, 4) '�ļ�λ��
        .Label26.Caption = Folderpathi
        Call FileChange
        .Label28.Caption = Rng.Offset(0, 8) '�ļ�����ʱ��
        .Label30.Caption = Rng.Offset(0, 10) '�ļ����
        .Label33.Caption = Rng.Offset(0, 21) '��ʶ����
        .TextBox5.Text = Rng.Offset(0, 19) '��ǩ1
        .TextBox6.Text = Rng.Offset(0, 20) '��ǩ2
        .TextBox3.Text = Rng.Offset(0, 12) '���ļ���
        .TextBox4.Text = Rng.Offset(0, 37) & Rng.Offset(0, 14) '���� & ����
        Text4Ch = .TextBox4.Text
        .ComboBox3.Text = Rng.Offset(0, 15) 'pdf������
        .ComboBox4.Text = Rng.Offset(0, 16) '�ı�����
        .ComboBox5.Text = Rng.Offset(0, 17) '��������
        .ComboBox2.Text = Rng.Offset(0, 18) '�Ƽ�����
        .ComboBox12.Text = Rng.Offset(0, 31) '��������
        .Label69.Caption = Rng.Offset(0, 24) '��������
        .ComboBox14.Text = Rng.Offset(0, 42) '�Ķ�״̬
        With ThisWorkbook.Sheets("temp") '�Ƿ����ñ༭
            If Len(.Cells(31, "ab").Value) = 0 Then
                Call DisablEdit      '��ֹ�༭
            Else                    '������ñ༭,��ô����Ҫ��һ������ļ���״̬
                Call EnablEdit
            End If
        End With
        
        If Rng.Offset(0, 25) = "ERC" Then
            .Label74.Caption = "Y"
        Else
            .Label74.Caption = "N"
        End If
        
        If Len(Rng.Offset(0, 13).Value) = 0 Then '��ȡ�����ļ�������Ϣ
            Call AtoName
        Else
            .TextBox3.Text = Rng.Offset(0, 13) '���ļ���
        End If
        
        .CheckBox13.Value = False
        If Len(Rng.Offset(0, 32)) > 0 Then Fileptc = 1: .CheckBox13.Value = True '�ļ�����״̬ 'checkbox��ֵ���ᴥ��click�¼�
        
        .Label236.Caption = ""
        If Len(Rng.Offset(0, 35)) > 0 Then
            .Label236.Caption = "Y"  '�ļ��Ƿ��������
        Else
            .Label236.Caption = "N"
        End If
        
        Call Text2a '��ȡ�ļ���ժҪ��Ϣ
        .Label106.Caption = Rng.Offset(0, 25) '�������� '���������Ϣ���ھ���ʾ�༭������Ϣ
        If Len(Rng.Offset(0, 25).Value) > 0 Then
            .CommandButton53.Enabled = True
            .CommandButton53.Caption = "�༭������Ϣ"
        End If
        strx = Rng.Offset(0, 36)
        BookCoverShow strx
        If .MultiPage1.Value <> 1 Then .MultiPage1.Value = 1 'ת��ҳ��
    End With
    Set Rng = Nothing
End Sub

Sub BookCoverShow(ByVal FilePath As String)
    If Len(FilePath) > 0 Then
        If fso.fileexists(FilePath) = True Then '����
            Imgurl = strx
            .Frame2.Width = 158
            .TextBox2.Width = 143
            .CommandButton134.Left = 610
            .CommandButton125.Left = 654
            With .Label239
                .Visible = True
                .Left = 728
                .Top = 94
                .Caption = "�������"
            End With
            With .Image1
                .Left = 708
                .Top = 108
                .Width = 84
                .Height = 122
                .Visible = True
                .Picture = LoadPicture(FilePath)
                .PictureSizeMode = fmPictureSizeModeStretch '����ͼƬ
            End With
            imgx = 1
        End If
    Else
        If imgx = 1 Then
            .Image1.Visible = False
            .Label239.Visible = False
            .Frame2.Width = 246
            .TextBox2.Width = 231
            .CommandButton134.Left = 698
            .CommandButton125.Left = 742
            imgx = 0
            Imgurl = ""
        End If
    End If
End Sub

Sub FileChange() '���׷����仯����Ϣ
    With UserForm3
        .Label76.Caption = Rng.Offset(0, 5) '�ļ���ʼ��С
        .Label112.Caption = Rng.Offset(0, 6) '�ļ��޸�ʱ��
        .Label27.Caption = Rng.Offset(0, 7) '�ļ���С
        .Label32.Caption = Rng.Offset(0, 12) '�򿪴���
        .Label31.Caption = Rng.Offset(0, 11) '�����ʱ��
        .Label71.Caption = Rng.Offset(0, 9) 'MD5
    End With
End Sub

Function Text2a() '��ȡ�ļ���ժҪ��Ϣ
    Dim TableName As String, str As String
    
    TableName = "ժҪ��¼"
    With UserForm3
        str = .Label29.Caption
        .TextBox2.Text = ""
        If Len(str) = 0 Then GoTo 100
        If RecData = True Then
            SQL = "select * from [" & TableName & "$] where ͳһ����='" & str & "'" '��ѯ����,�����,�����д������,����оͿ�����ʾ����
            Set rs = New ADODB.Recordset
            rs.Open SQL, Conn, adOpenKeyset, adLockOptimistic
            If rs.BOF And rs.EOF Then             '�����ж������ҵ�����
                rs.Close
                Set rs = Nothing
            Else
                If IsNull(rs(5)) = False Then .TextBox2.Text = rs(5) 'ע�����������Ϊ�յ�����(null), ע����������ʾ�յ����� ��empty�ĵ�
                rs.Close
                Set rs = Nothing
            End If
        End If
    End With
100
End Function

Function AtoName() '�����������ļ���
    Dim i As Byte
    Dim strx As String, strx1 As String
    Dim strLen As Byte
    
    With UserForm3
        strx = .Label23.Caption
        strLen = Len(strx)
        If strLen = 0 Or strLen < 5 Then Exit Function
        If strLen > 16 Then
            i = 10
        ElseIf strLen > 12 And strLen < 17 Then
            i = 7
        ElseIf strLen > 8 And strLen < 13 Then
            i = 5
        Else
            i = 3
        End If
        strx1 = Left$(strx, i)
        .TextBox3.Text = strx1
        Text3Ch = .TextBox3.Text 'ע�����ﲻҪֱ��ʹ��strx1��ֵ,��ֹ��ansi�����ַ������ĸ���
    End With
End Function

Sub EnablEdit() '���ñ༭
  Dim strx As String
  With UserForm3
        If Len(.Label29.Caption) = 0 Then Exit Sub '��ҳ������Ϣ��ʱ����ܱ༭,ɾ���ļ����ɲ���
        .TextBox2.Enabled = True
        .TextBox3.Enabled = True
        .TextBox3.SetFocus
        .TextBox4.Enabled = True
        .TextBox5.Enabled = True
        .TextBox6.Enabled = True
        .ComboBox2.Enabled = True
        .ComboBox5.Enabled = True
        .ComboBox12.Enabled = True
        .CommandButton8.Enabled = True
        .ComboBox14.Enabled = True
        .ComboBox4.Enabled = True '�ı�����
        strx = UCase(.Label24.Caption) '��չ��
        If strx = "EPUB" Or strx = "MOBI" Or strx = "TXT" Then .CommandButton53.Enabled = True '����
        If strx = "PDF" Then
            .ComboBox3.Enabled = True
            .CommandButton53.Enabled = True
        End If
    End With
End Sub

Sub DisablEdit() '����-���������ֹ�༭
    With UserForm3
        .TextBox2.Enabled = False 'ժҪ��Ϣ
        .TextBox3.Enabled = False '���ļ���
        .TextBox4.Enabled = False '����
        .TextBox5.Enabled = False '��ǩ
        .TextBox6.Enabled = False '��ǩ
        .ComboBox2.Enabled = False '�Ƽ�ָ��
        .ComboBox3.Enabled = False 'PDF������
        .ComboBox4.Enabled = False '�ı�����
        .ComboBox5.Enabled = False '��������
        .ComboBox12.Enabled = False '��������
        .ComboBox14.Enabled = False '�Ķ�״̬
        .CommandButton53.Enabled = False '��������
        .CommandButton54.Visible = False '��Ӷ�����Ϣ
        .TextBox16.Visible = False '������Ϣ�༭
        .TextBox17.Visible = False '������Ϣ�༭
        .TextBox15.Visible = False '������Ϣ�༭
        .CommandButton8.Enabled = False '�����Ϣ
        .CommandButton56.Enabled = False 'md5����
    End With
End Sub
'----------------------------------------------------------��ʾ�������
Sub Rewds() '���ô���
    If ThisWorkbook.Application.Visible = False Then
        ThisWorkbook.Application.Visible = True
        UserForm4.Hide
        UserForm4.Show
        UserForm4.Caption = "Mini"
    End If
    ThisWorkbook.Windows(1).WindowState = xlMinimized
End Sub

Sub HideOption() '������ҳ�Ľ���
    With ThisWorkbook.Windows(1)
        .DisplayFormulas = False
        .DisplayHeadings = False
        .DisplayHorizontalScrollBar = False
        .DisplayVerticalScrollBar = False
        .DisplayWorkbookTabs = False
    End With
    ThisWorkbook.Application.DisplayFormulaBar = False
End Sub

Sub Showoption() '�ָ�ԭ���Ľ���
ThisWorkbook.Application.DisplayFormulaBar = True
    With ThisWorkbook.Windows(1)
        .DisplayFormulas = True
        .DisplayHeadings = True
        .DisplayHorizontalScrollBar = True
        .DisplayVerticalScrollBar = True
        .DisplayWorkbookTabs = True
    End With
End Sub

Function DataSwitch()      '��Ԫ��תΪ�ض���ʱ���ʽ
    With ThisWorkbook.Sheets("������")
        .Range("w27:w33").NumberFormatLocal = "yyyy/m/d h:mm;@"
    End With
End Function

Function TextSwitch() '�ı���ʽת��
    ThisWorkbook.Sheets("���").Columns("c:c").NumberFormatLocal = "@"  'ǿ��ת��Ϊ�ı�, ���ڱ���Ϊ�����ֵ��ļ�,�����������ʾ���ܳ��ָ�ʽ�쳣
End Function
