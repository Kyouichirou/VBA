VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm9 
   Caption         =   "�޸��ļ���"
   ClientHeight    =   3480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8730
   OleObjectBlob   =   "UserForm9.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "UserForm9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text  '�����ִ�Сд
Option Explicit
Dim addrowx As Integer '����
Dim standfn As String
Dim standfp As String, standfp1 As String
Dim standx1 As String
Dim standx2 As String
Dim standx3 As String
Dim standx4 As String
Dim standx5 As String
Dim standx6 As String

Private Sub CommandButton1_Click() '�޸��ļ���
    Dim newname As String
    Dim str1 As String
    
    On Error GoTo 100 '������
    If Len(Trim(Me.TextBox1.Text)) = 0 Or Len(standfn) = 0 Then Exit Sub 'δ��������
    
    With ThisWorkbook.Sheets("���")
        If fso.fileexists(standfp) = False Then '�ļ�Ϊ��
            Me.Label1.Caption = "�ļ�������"
            Exit Sub
        End If
        
        newname = Me.TextBox1.Text
        str1 = newname & "." & .Range("d" & addrowx)  '������չ�������ļ���
        
        If fso.fileexists(.Range("f" & addrowx) & "\" & str1) = True Then '�ж��ļ��Ƿ�����
            Me.Label1.Caption = "�ļ�����"
            Me.TextBox1.Text = ""
            Exit Sub
        End If
        
        If ErrCode(newname, 1) > 1 Then '���ļ���������ȥ����Ƿ�����쳣�ַ�
            Me.Label1.Caption = "������ļ��������쳣�ַ�" '�޸��ļ������Ƿ�����쳣�ַ�''�����������Ƴ�������/���ڿո�,������ִ������-���ѭ��(inputbox)
            Me.TextBox1.Text = ""
            Exit Sub
        Else
            fso.GetFile(standfp).Name = str1 '�����ļ�����ʹ��fso��ȫ��ʹ��fso,������ַ�ansi�ַ�������
        End If
        .Range("c" & addrowx) = str1 '�µ��ļ���
        str1 = .Range("f" & addrowx) & "\" & str1
        .Range("e" & addrowx) = str1 '�µ�·��
        standfp1 = str1
        .Range("ae" & addrowx) = "" '���ԭ�������ִ����쳣�ַ���ô��ȥ��������
        
        If .Range("ac" & addrowx) = "EDC" Then '�ļ���/�ļ������ڷ�ansi����
            .Range("ac" & addrowx) = "EPC" 'ֻ���ļ��в��ִ��������ַ�
        Else
            .Range("ac" & addrowx) = ""
            .Range("ab" & addrowx) = ""
        End If
        If Len(standx5) > 0 Then ThisWorkbook.Sheets("Ŀ¼").Cells(standx5, standx6) = Now
        Me.Label1.Caption = "�޸ĳɹ�!"
        Me.CommandButton2.Enabled = True '
        If .Range("d" & addrowx) = "txt" Then Me.Label2.Caption = "txt�ĵ������ڴ򿪵�״̬���ƶ�\������\ɾ��"      '�����û�,txt�ĵ������ڴ򿪵�״̬���ƶ����������Ȳ���
        End With
    Exit Sub
100
    If Err.Number = 70 Then '70�ļ���,76�ļ�Ŀ¼������
        Me.Label1.Caption = "�ļ����ڴ�״̬,����ʧ��"
    Else
        Me.Label1.Caption = "�����쳣,������"
    End If
    Err.Clear
End Sub

Private Sub CommandButton2_Click() '�ָ�ԭ�е��ļ���
    With ThisWorkbook.Sheets("���")
    If fso.fileexists(standfp1) = True Then
        fso.GetFile(standfp1).Name = standfn '�ָ������ļ�����
    Else
        DeleMenu (addrowx)
        Me.Label1.Caption = "�ļ��Ѿ���ʧ"
        Exit Sub
    End If
        .Range("c" & addrowx) = standfn
        .Range("e" & addrowx) = standfp
        .Range("ae" & addrowx) = standx1
        .Range("ab" & addrowx) = standx2
        .Range("ac" & addrowx) = standx3
        .Range("af" & addrowx) = standx4
        If Len(standx5) > 0 Then ThisWorkbook.Sheets("Ŀ¼").Cells(standx5, standx6) = Now '�ļ����±��޸ĵ�ʱ��
    End With
End Sub

Private Sub CommandButton3_Click() '���ļ���
    With Me
        .TextBox1 = .Label3.Caption
    End With
End Sub

Private Sub CommandButton4_Click() '�����ļ���
    With Me
        .TextBox1 = .Label4.Caption
    End With
End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) '�����ļ���������Щϵͳ��ֹ���ַ�
    With Me
        Select Case KeyAscii
            Case asc("/"), asc("\"), asc(":"), asc("*"), asc("?"), asc("<"), asc(">"), asc("|") 'MsgBox "��������Ƿ��ַ���""/\ : * ? <> |""", vbInformation, "����"
            .Label1.Caption = "��������Ƿ��ַ���""/\ : * ? <> |"
            .TextBox1.Text = ""
            KeyAscii = 0
        Case Else
           .Label1.Caption = ""
        End Select
    End With
End Sub

Private Sub UserForm_Initialize() '��ȡ��ʼֵ,���ڻָ�
    Dim rngtime As Range
    Dim str As String, strx1 As String, strx2 As String
    
    If Statisticsx = 1 Then Exit Sub
    addrowx = Selection.Row
    If addrowx < 6 Then Exit Sub
    With ThisWorkbook.Sheets("���")
        standfn = .Range("c" & addrowx).Value '�ļ���
        Me.Label2.Caption = standfn
        standfp = .Range("e" & addrowx).Value '�ļ�·��
        standx1 = .Range("ae" & addrowx).Value '�쳣�ַ����
        standx2 = .Range("ab" & addrowx).Value '�쳣�ַ����
        standx3 = .Range("ac" & addrowx).Value
        standx4 = .Range("af" & addrowx).Value
        str = .Range("f" & addrowx).Value & "\"
        strx1 = .Range("o" & addrowx).Value '���ļ���
        strx2 = .Range("y" & addrowx).Value '��������
    End With
    
    With Me
        If Len(strx1) > 0 Then .Label3.Caption = strx1: .CommandButton3.Visible = True
        If Len(strx2) > 0 Then .Label4.Caption = strx2: .CommandButton4.Visible = True
    End With
    
    With ThisWorkbook.Sheets("Ŀ¼")
        Set rngtime = .Cells(4, 3).Resize(.[b65536].End(xlUp).Row, .Cells.SpecialCells(xlCellTypeLastCell).Column).Find(str, lookat:=xlWhole)
        If Not rngtime Is Nothing Then
            standx5 = rngtime.Row
            standx6 = rngtime.Column
        End If
    End With
End Sub
