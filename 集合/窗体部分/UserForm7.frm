VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm7 
   Caption         =   "About Me"
   ClientHeight    =   5460
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9000
   OleObjectBlob   =   "UserForm7.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "UserForm7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const ChCode As String = "HLAstaticx"

Private Sub CommandButton1_Click() 'support me
    UserForm10.Show
End Sub

Private Sub Label19_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim textb As Object, strx As String
    
    With Me
        Set textb = .Controls.Add("Forms.TextBox.1", "Text1", False) '�Դ�����ʱtextbox�ķ�ʽʵ�ָ�������
        strx = .Label19.Caption
        With textb
            .Text = strx
            .SelStart = 0
            .SelLength = Len(.Text)
            .Copy
        End With
        .Label18.Caption = "�ʼ���ַ�Ѹ���"
    End With
    Set textb = Nothing
End Sub

Private Sub UserForm_Initialize()
    Dim k As Integer, i As Integer, j As Integer, m As Integer, lc As Integer, bl As Integer, n As Integer, p As Integer
    Dim strx As String, a As Byte, b As Byte, c As Integer
    Dim obj As Object, strx1 As String, objdata As Object
    Dim UF As Object, wd As Object, wdc As Integer, strx2 As String
    
    On Error Resume Next
    Statisticsx = 1 '���������еĴ����ʾ,��ʼ������������ͳ��,���ٴ����ʼ�������м���/ж��
    With ThisWorkbook
        .Application.ScreenUpdating = False
        lc = .VBProject.VBComponents.Count
        For Each obj In .VBProject.VBComponents '����userform�ؼ�������
            If obj.Name <> "UserForm7" Then
                If obj.type = 3 Then
                    Set UF = UserForms.Add(obj.Name) '������̻��ʼ������,ע��ĳЩ�����ڳ�ʼ�����߼���ʱִ�е��¼�, �Ƿ��������,�������ɴ���
                    m = m + UF.Controls.Count
                    Unload UF
                End If
            End If
        Next
        For i = 1 To lc
            With .VBProject.VBComponents.item(i).CodeModule
                p = .CountOfLines '������
                k = k + p
                For n = p To 1 Step -1
                    strx = .Lines(n, 1)
                    strx1 = Trim(strx)
                    If strx1 = vbNullString Then
                        bl = bl + 1 '�������
                    Else
                        If InStr(1, strx1, Chr(39), vbBinaryCompare) > 0 Then j = j + 1 'chr(39)'������
                    End If
                Next
            End With
        Next
         '-----------------�������еĿؼ�����
        a = .Sheets.Count
        For b = 1 To a
            c = c + .Sheets(b).Shapes.Count
        Next
        '-----------------------------------------��ȡword���ֵ����/Ҳ����ͨ������docm�ڵĺ���ͳ��
        strx2 = ThisWorkbook.Path & "\LB.docm"
        If fso.fileexists(strx2) = True Then '-------���ֱ�Ӵ���word����,ִ�е��ٶȽ���
            SetClipboard ChCode '����һ��ֵ�����а�,���ڿ���word�ļ���ִ��(open�¼�)
            If Err.Number > 0 Then Err.Clear
            Set wd = GetObject(strx2) '-�б��ڻ�ȡword����, �������ֱ�ӻ�ȡ���ĵ��Ķ���(�ĵ����ڹرյ�״̬),������ִ���
            If Err.Number > 0 Then Set wd = CreateObject(strx2): Err.Clear
            'wd.Application.Run "VBprojectStatic" '-----ֱ������word�ڵ�sub����
            '--------------------------------https://docs.microsoft.com/zh-cn/office/vba/api/word.application.run
            With wd.VBProject.VBComponents
                wdc = .Count
                lc = lc + wdc
                For i = 1 To wdc
                    With .item(i).CodeModule
                        p = .CountOfLines '������
                        k = k + p
                        For n = p To 1 Step -1
                            strx = .Lines(n, 1)
                            strx1 = Trim(strx)
                            If strx1 = vbNullString Then
                                bl = bl + 1 '�������
                            Else
                                If InStr(strx1, Chr(39)) > 0 Then j = j + 1 'chr(39)'������
                            End If
                        Next
                    End With
                Next
            End With
            If GetClipboard <> ChCode Then  '���ԭ����word�ļ����ڹرյ�״̬,�͹ر�word,��word������open�¼�
            '------------------------------------------------����Ӽ��а��ȡ��"HLAstaticx"ֵ, ��ô������ռ��а�,����docm�ļ����ڴ򿪵�״̬,�����ᴥ��open�¼�
            '-------https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/methods-microsoft-forms
                wd.Close savechanges:=False
                Set objdata = New DataObject
                objdata.SetText "" '---------�����а�����Ϊ��,��ֹ������word��������ɸ���
                objdata.PutInClipboard
                Set objdata = Nothing
            End If
            Set wd = Nothing
        End If
        With Me
            .Label15 = Format(Now, "yyyy/mm/dd")
            .Label1.Caption = k
            .Label4.Caption = lc
            .Label6.Caption = m + c + 76
            .Label8.Caption = j
            .Label17.Caption = bl
        End With
        .Application.ScreenUpdating = True
    End With
    Statisticsx = 0
    Set obj = Nothing
    Set UF = Nothing
End Sub
