Attribute VB_Name = "�ļ�����"
Option Compare Text                       '�����ִ�Сд
Option Explicit
Dim filedyn As Boolean '�жϱ����ļ���ɾ���Ƿ���
Public Reasona As String, Reasonb As String, Filehashx As String
Public DeleFilex As Byte, MDeleFilex As Byte, AddPlistx As Byte '��ɾ��

Function FileDeleExc(ByVal FilePath As String, ByVal addrow As Integer, ByVal Px As Byte, ByVal cmCode As Byte, Optional ByVal filex As String, _
Optional ByVal FileName As String) As Boolean 'ִ��ɾ�� '�ļ�·��,�ļ���չ��,�����к�,�Ƿ��з�ansi,ִ���������Դ
    Dim strx As String, cmCodex As Byte, k As Byte
    
    FileDeleExc = True
    filedyn = False 'ʹ��ǰ��ʼ��
    If cmCode = 0 Then           'ɾ��������Դ�ڱ��
        k = FileTest(FilePath, filex, FileName)
        If k >= 3 Then MsgShow "�ļ����ڴ�״̬", "Warning", 1500: FileDeleExc = False: Exit Function
    End If
    If k = 0 Or cmCode = 1 Then
        On Error GoTo 100
        If Px > 0 Then
            fso.DeleteFile (FilePath)
        Else
            DeleteFiles (FilePath)
        End If
        filedyn = True
    End If
    Call DeleMenu(addrow) 'ɾ��Ŀ¼
    Exit Function
100
    If Err.Number = 70 Then
        MsgBox "�ļ����ڴ򿪵�״̬"
    Else
        MsgBox "�쳣"
    End If
    FileDeleExc = False
    Err.Clear
End Function

Sub DeleMenu(ByVal addrow As Integer) 'ɾ��Ŀ¼ 'optional��ʾ������д���߲�д,λ��Ҫ���������,��֪���沿�ֵĲ���ȫ����Ҫдoptional
    Dim addloc As String, i As Byte
    Dim arrback(1 To 34) As String
                                      'ɾ���ļ���ɾ��Ŀ¼�ֿ�
    On Error GoTo 100
    With ThisWorkbook.Sheets("���")
        For i = 1 To 32
            arrback(i) = .Cells(addrow, i + 1).Value '��Ҫɾ������Ϣ�洢������
        Next
        'Ĩ��ժҪ����Ϣ/ɾ������
        arrback(33) = Reasona
        arrback(34) = Reasonb
        If Len(Filehashx) > 0 Then arrback(10) = Filehashx
       Call DeleOverBack(arrback(1), arrback(2), arrback(3), arrback(4), arrback(5), arrback(6), arrback(7), arrback(8), arrback(9), arrback(10), arrback(11), _
            arrback(12), arrback(13), arrback(14), arrback(15), arrback(16), arrback(17), arrback(18), arrback(19), arrback(20), arrback(21), arrback(22), arrback(23), _
            arrback(24), arrback(25), arrback(26), arrback(27), arrback(28), arrback(29), arrback(30), arrback(31), arrback(32), arrback(33), arrback(34))

        addloc = arrback(5)
        .rows(addrow).Delete Shift:=xlShiftUp
        Call DeleFileOver(addloc) 'ɾ����ִ�м��Ŀ¼���ļ��Ĵ���״��
    End With
    If UF3Show = 3 Then DeleFilex = 1
100
    Reasona = ""
    Filehashx = ""
    Reasonb = "" '����֮������
End Sub

Function DeleFileOver(ByVal tfolder As String, Optional ByVal cmCode As Byte)
'-------------------------------------------------------------------------���������Ƴ�Ŀ¼�ķ���,���ļ����κι����ļ��е��ļ�����������,Ŀ¼�е����ݲŻ�ȫ���Ƴ� 'ɾ���ļ��ĺ�������
    Dim rngad As Range, rngadx As Range
    Dim tfolderp As String
    Dim c As Byte, alow As Integer, dlow As Integer, i As Integer, flow As Integer, csc As Byte, blow As Integer
    Dim filec As Integer
    
    With ThisWorkbook
        .Application.ScreenUpdating = False
        .Application.Calculation = xlCalculationManual
        With .Sheets("���")
            tfolderp = tfolder & "\"     '��Ҫ�Ż�
            flow = .[f65536].End(xlUp).Row
            If cmCode = 0 Then
            Set rngadx = .Range("e6:e" & flow).Find(tfolderp, lookat:=xlPart) '����ļ����Ƿ������������ļ����ļ�����Ŀ¼
            End If
            c = UBound(Split(tfolder, "\")) '�ļ��㼶
            With ThisWorkbook.Sheets("Ŀ¼")
                blow = .[b65536].End(xlUp).Row
                csc = .Cells.SpecialCells(xlCellTypeLastCell).Column
                If cmCode = 1 Then GoTo 401
                If Not rngadx Is Nothing Then
                    If filedyn = True Then                          '����ļ�ɾ�����������ִ��
                        Set rngad = .Cells(4, 3).Resize(blow, csc).Find(tfolderp, lookat:=xlWhole) '��ȷ����
                        If Not rngad Is Nothing And filedyn = True Then
                            rngad.Offset(0, 2) = Now '���������ļ��е��޸�ʱ��
                            filec = rngad.Offset(0, 4).Value '�����ļ�������
                            If IsNumeric(filec) = True Then
                                If filec > 1 Then
                                    filec = filec - 1
                                    rngad.Offset(0, 4) = filec
                                Else
                                    rngad.Offset(0, 4) = 0
                                End If
                            End If
                        End If
                    End If
                    Exit Function
                End If
                Set rngadx = Nothing
                Set rngad = Nothing
            End With
            '�ļ������ļ��Ѳ�����Ŀ¼
            With ThisWorkbook.Sheets("������")
                dlow = .[e65536].End(xlUp).Row
                If dlow < 37 Then GoTo 401
                Set rngad = .Range("e37:d" & dlow).Find(tfolder, lookat:=xlWhole) ''������ļ���
                If Not rngad Is Nothing Then
                    i = rngad.Row
                    If UF3Show = 3 Or UF3Show = 1 Then MDeleFilex = 1
                    If i = dlow Then
                        .Range("e" & i & ":" & "j" & i).ClearContents '��������һ��,��ֱ�Ӵ���
                    Else
                        rngad.Delete Shift:=xlUp 'ɾ����Ԫ��(�����������)
                        rngad.Offset(0, 4).Delete Shift:=xlUp 'ɾ�����ʱ��
                    End If
                End If
                Do
                    Set rngad = .Range("d37:d" & dlow).Find(tfolderp, lookat:=xlPart) '���������ļ�Ŀ¼ȫ���Ƴ�,������ļ���
                    If Not rngad Is Nothing Then rngad.Delete Shift:=xlUp: rngad.Offset(0, 4).Delete Shift:=xlUp '����ֱ��ɾ���ķ�ʽ,���Բ�Ҫ�ϲ�����ĵ�Ԫ��,�����������־��浯��
                Loop Until rngad Is Nothing '����ѭ��
            End With
401
            Set rngad = Nothing
            If c = 1 Then
                GoTo 100         '1���ļ���ֱ����������ɸѡ
            ElseIf c > 1 Then
                tfolderp = Split(tfolder, "\")(0) & "\" & Split(tfolder, "\")(1) & "\" 'һ��Ŀ¼
            End If
            Set rngad = .Range("f6:f" & flow).Find(tfolderp, lookat:=xlPart) '�ж��Ƿ񻹴��ڴ��ļ������������ļ��е���Ϣ
        End With
100
        With .Sheets("Ŀ¼")
            If rngad Is Nothing Then '�����������ļ��еĹ����ļ���
                Set rngad = .Cells(4, 3).Resize(blow, csc).Find(tfolderp, lookat:=xlPart) 'ģ������                        '���һ���ļ����������е��ļ�
                If Not rngad Is Nothing Then
                    If .AutoFilterMode = True Then .AutoFilterMode = False 'ɸѡ������ڿ���״̬��ر�
                    .Range("a3:a" & blow).AutoFilter Field:=1, Criteria1:=.Cells(rngad.Row, 1).Value
                    .Range("a3").Offset(1).Resize(blow - 3).SpecialCells(xlCellTypeVisible).Delete Shift:=xlShiftUp 'ɾ����ɸѡ�����Ľ��
                    .Range("a3").AutoFilter
                End If
            Else  '���������ļ����Լ����ļ���
                tfolderp = tfolder & "\"  '���¸�ֵ
                Do
                    Set rngad = .Cells(4, 3).Resize(blow, csc).Find(tfolderp, lookat:=xlPart) '��ȷ���� '��������ض����ļ�
                    If Not rngad Is Nothing Then .rows(rngad.Row).Delete Shift:=xlShiftUp
                Loop Until rngad Is Nothing
            End If
        End With
        .Application.ScreenUpdating = True
        .Application.Calculation = xlCalculationAutomatic
    End With
    Set rngad = Nothing
    Set rngadx = Nothing
End Function

Function OpenFileLocation(ByVal address As String)  '���ļ�����λ��
    If Len(address) = 0 Then Exit Function
    If fso.folderexists(address) = False Then Exit Function
    Shell "explorer.exe " & address, vbNormalFocus
End Function

Function FileCopy(ByVal addressx As String, ByVal FileName As String, ByVal adrowx As Integer, Optional ByVal cmCode As Byte) As Boolean '�ļ�����
    Dim mynewpath As String
    Dim rngad As Range
    Dim Filesize As Long
    Dim strfolder As String, strx As String
    
    On Error GoTo 100
    FileCopy = False
    If Len(addressx) = 0 Then Exit Function '��ֵ�˳�
    If cmCode = 0 Then
        If fso.fileexists(addressx) = False Then '�ļ�������,��ִ��ɾ������,������ļ��Ĵ��ڵ�����ֽ��ÿһ�εĲ���
            Call DeleMenu(adrowx)
            Exit Function
        End If
    End If
    With Application.FileDialog(msoFileDialogFolderPicker) '�ļ���ѡ�񴰿�
        .Show
        If .SelectedItems.Count = 0 Then Exit Function 'δѡ���ļ������˳�sub
        strfolder = .SelectedItems(1)
    End With
    If CheckFileFrom(strfolder, 2) = True Then MsgShow "�ļ�λ������", "Warning", 1500: Exit Function '�����ӵ����ļ����Ƿ�Ϊ����λ��
    mynewpath = strfolder & "\"
    strx = Left(strfolder, 1) '�̷�
    Filesize = fso.GetFile(addressx).Size
    If Filesize > fso.GetDrive(strx).AvailableSpace Then MsgBox "���̿ռ䲻��!", vbCritical, "Warning": Exit Function '�жϴ����Ƿ����㹻�Ŀռ�
    
    If fso.fileexists(mynewpath & FileName) = True Then
        MsgBox "�ļ��Ѵ���"
        Exit Function
    Else
        If Filesize <= 52428800 Then '����50M
            fso.CopyFile (addressx), mynewpath
        Else                         '������Χ,����cmd����ȥִ��
            Shell ("cmd /c" & "copy " & addressx & Chr(32) & strfolder), vbHide
        End If
        With ThisWorkbook.Sheets("Ŀ¼")
            Set rngad = .Cells(4, 3).Resize(.[b65536].End(xlUp).Row, .Cells.SpecialCells(xlCellTypeLastCell).Column).Find(mynewpath, lookat:=xlWhole) '��ȷ����'������Ƶ��ļ���λ����Ŀ¼�Ͼ͸���ʱ��
            If Not rngad Is Nothing Then rngad.Offset(0, 2) = Now
        End With
        FileCopy = True
    End If
100
    Set rngad = Nothing
End Function

Function DeleOverBack(ByVal str1 As String, ByVal str2 As String, ByVal str3 As String, ByVal str4 As String, ByVal str5 As String, _
ByVal str6 As String, ByVal str7 As String, ByVal str8 As String, ByVal str9 As String, _
ByVal str10 As String, ByVal str11 As String, ByVal str12 As String, ByVal str13 As String, ByVal str14 As String, _
ByVal str15 As String, ByVal str16 As String, ByVal str17 As String, ByVal str18 As String, ByVal str19 As String, _
ByVal str20 As String, ByVal str21 As String, ByVal str22 As String, ByVal str23 As String, ByVal str24 As String, _
ByVal str25 As String, ByVal str26 As String, ByVal str27 As String, ByVal str28 As String, ByVal str29 As String, _
ByVal str30 As String, ByVal str31 As String, ByVal str32 As String, ByVal str33 As String, ByVal str34 As String) 'ɾ��֮ǰ,����-�޸�ժҪ�ϵ���Ϣ

    Dim TableName As String
    Dim strx1 As String
    
    If Len(str1) = 0 Then Exit Function
    If RecData = True Then
        TableName = "ժҪ��¼"
        strx1 = "DL-" & Mid$(str1, 5, Len(str1) - 4) '�޸�ժҪ��Ϣ
        SQL = "select * from [" & TableName & "$] where ͳһ����='" & str1 & "'"                                       '��ѯ����
        Set rs = New ADODB.Recordset
        rs.Open SQL, Conn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount <> 0 Then
            SQL = "UPDATE [" & TableName & "$] SET ͳһ����='" & strx1 & "' WHERE ͳһ����='" & str1 & "'"          '����ժҪ�ϵ���Ϣ
            Conn.Execute (SQL)
        End If
        rs.Close
        Set rs = Nothing
        TableName = "ɾ������"  'ɾ������
        SQL = "Insert into [" & TableName & "$] (ͳһ����, �ļ���, �ļ�����, �ļ�·��, �ļ�����λ��, �ļ���ʼ��С, �ļ��޸�ʱ��, �ļ���С, �ļ�����ʱ��, �ļ�Hash, �ļ����, �����ʱ��, �ۼƴ򿪴���, ���ļ���, ����, PDF������, �ı�����, ��������, �Ƽ�ָ��, ��ǩ1, ��ǩ2, ��ʶ���, ���ʱ��, ����, ����, ����, �쳣�ַ����, ��ע, ��Դ, �ļ����쳣�ַ�, �ļ�λ���쳣�ַ�, ��������, ɾ��ԭ��, ɾ����ע) Values ('" & str1 & "','" & str2 & "','" & str3 & "','" & str4 & "','" & str5 & "'," & str6 & ",'" & str7 & "','" & str8 & "','" & str9 & "','" & str10 & "','" & str11 & "','" & str12 & "','" & str13 & "','" & str14 & "','" & str15 & "','" & str16 & "','" & str17 & "','" & str18 & "','" & str19 & "','" & str20 & "','" & str21 & "','" & str22 & "','" & str23 & "','" & str24 & "','" & str25 & "','" & str26 & "','" & str27 & "','" & str28 & "','" & str29 & "','" & str30 & "','" & str31 & "','" & str32 & "','" & str33 & "','" & str34 & "')"
        Conn.Execute (SQL)
    End If
End Function

Function AddPList(ByVal filecode As String, ByVal filen As String, ByVal cmfrom As Byte) As Boolean '��ӵ������Ķ�
    Dim i As Integer, k As Byte
    Dim strx As String, strx1 As String

    With ThisWorkbook.Sheets("������")        '�������¼����Ϣ
'        If .Range("k27").NumberFormatLocal <> "yyyy/m/d h:mm;@" Then .Range("k27:k33").NumberFormatLocal = "yyyy/m/d h:mm;@" '��ʽ����
        strx1 = LCase(Right$(filen, Len(filen) - InStrRev(filen, ".")))
        If strx1 Like "xl*" Then
            If strx1 <> "xls" And strx1 <> "xlsx" Then AddPList = False: MsgBox "�������ļ����������", vbCritical, "Warning": Exit Function
        End If
        For k = 27 To 33
            strx = .Range("i" & k).Value
            If filecode = strx Then AddPList = False: Exit Function '����Ѿ�����,��д��
            If Len(strx) = 0 Then Exit For       '����ֵʱ,�˳�ѭ��
        Next
        ThisWorkbook.Application.ScreenUpdating = False
        If Len(.Range("i27").Value) = 0 Then             'ȷ����ӽ�����ֵ����һֱ���ڵ�һ��
           .Range("i27") = filecode
           .Range("d27") = filen
           .Range("k27") = Now
        Else
            For i = 33 To 28 Step -1   'i������byte����?                               '���һ�е����ݲ��ϱ�����д��
                .Range("d" & i) = .Range("d" & i - 1)
                .Range("i" & i) = .Range("i" & i - 1)
                .Range("k" & i) = .Range("k" & i - 1)
            Next
            .Range("i27") = filecode
            .Range("d27") = filen
            .Range("k27") = Now
        End If
    End With
    AddPList = True
    If UF3Show = 3 Then AddPlistx = 1 '��������ʱ
    ThisWorkbook.Application.ScreenUpdating = True
    If cmfrom = 0 Then
        ThisWorkbook.Sheets("���").Label1.Caption = "�����ɹ�"
    Else
        UserForm3.Label57.Caption = "�����ɹ�"
    End If
End Function

Sub FileMove()  '�ļ��ƶ�
    Dim addrowx As Integer, i As Byte, Filesize As Long
    Dim addressx As String, tfolderx As String
    Dim rngad As Range, filext As String
    Dim strfolder As String, mynewpath As String, strx As String
    
    On Error GoTo 100
    addrowx = Selection.Row
    With ThisWorkbook.Sheets("���")
        addressx = .Range("e" & addrowx).Value
        If addrowx < 6 Or Len(addressx) = 0 Then Exit Sub
        filext = .Range("d" & addrowx).Value
        i = FileTest(addressx, filext, .Cells(addrowx, "c").Value)
        Select Case i                                 ''����ļ��Ĵ���/�Ƿ��ڴ򿪵�״̬
            Case 1: strx = "δ��ȡ����Чֵ"
            Case 2: strx = "�ļ�������": Call DeleMenu(addrowx) 'ɾ�����Ŀ¼
            Case 3: strx = "�ļ���txt�ļ�"
            Case 4: strx = "�ļ����ڴ򿪵�״̬"
            Case 5: strx = "�����쳣"
        End Select
        .Label1.Caption = strx
        If i <> 0 And i <> 3 And i <> 6 Then Exit Sub
    
        With Application.FileDialog(msoFileDialogFolderPicker) '�ļ���ѡ�񴰿�
            .Show
            If .SelectedItems.Count = 0 Then Exit Sub      'δѡ���ļ������˳�sub
            strfolder = .SelectedItems(1)
        End With
        
        tfolderx = .Range("f" & addrowx).Value
        If tfolderx = strfolder Then Exit Sub '��ͬ���ļ���
        If CheckFileFrom(strfolder, 2) = True Then MsgShow "�ļ�λ������", "Warning", 1500: Exit Sub '�����ӵ����ļ����Ƿ�Ϊ����λ��
        strx = Left(strfolder, 1)
        Filesize = fso.GetFile(addressx).Size '��ȡ�ļ���ʵ�ʴ�С
        If Filesize > fso.GetDrive(strx).AvailableSpace Then MsgBox "���̿ؼ�����!", vbCritical, "Warning": Exit Sub '�жϴ����Ƿ����㹻�Ŀռ�
        
        mynewpath = strfolder & "\" 'ע��,Ǩ���ļ���ʱ��,Ŀ���ļ���Ҫ��"\"
        If fso.fileexists(mynewpath & filext) = True Then   '�ж��ļ��Ƿ��Ѵ���
            MsgBox "�ļ��Ѵ���"
            Exit Sub
        '---------------------------------------------------------------------------------------------------------------------------------ǰ��׼��
        Else
            If Filesize <= 52428800 Then '����50M
                fso.MoveFile (addressx), mynewpath
            Else                         '������Χ,����cmd����ȥִ��
                Shell ("cmd /c" & "move " & addressx & Chr(32) & mynewpath), vbHide
            End If
        End If
        DisEvents
        If ErrCode(mynewpath, 1) > 1 Then '����µ��ļ��е�·���Ƿ�����쳣�ַ�'����ȡ�ļ�·�����쳣�ַ���λ��
            .Range("af" & addrowx) = errcodepx
            .Range("ab" & addrowx) = "ERC"
            If .Range("ae" & addrowx) = "ENC" Then .Range("ac" & addrowx) = "EDC"
        Else
            .Range("af" & addrowx) = ""
            If .Range("ae" & addrowx) = "EDC" Then
                .Range("ac" & addrowx) = "ENC"
            Else
                .Range("ab" & addrowx) = ""
            End If
        End If
        
        DeleFileOver (tfolderx) '���ļ����ƶ���,��������ͬɾ��
        
        With ThisWorkbook.Sheets("Ŀ¼") '�ļ�Ǩ�ƺ�,ԭ����/���ڵ��ļ����޸�ʱ�䷢���仯
            Set rngad = .Cells(4, 3).Resize(.[b65536].End(xlUp).Row, .Cells.SpecialCells(xlCellTypeLastCell).Column).Find(mynewpath, lookat:=xlWhole)
            If Not rngad Is Nothing Then rngad.Offset(0, 2) = Now: rngad.Offset(0, 4) = rngad.Offset(0, 4) + 1
        End With
        .Range("e" & addrowx) = mynewpath & .Range("c" & addrowx) '����Ŀ¼�ϵ���Ϣ
        .Range("f" & addrowx) = strfolder
        EnEvents
        If i = 3 Then
            MsgBox "�޸ĳɹ�!" & Chr(13) & "txt�ĵ������ڴ򿪵�״̬���ƶ�\������\ɾ��"  'txt�ĵ������ڴ򿪵�״̬���ƶ����������Ȳ���
        Else
            MsgBox "�޸ĳɹ�!"
        End If
    End With
    Set rngad = Nothing
Exit Sub
100
    If Err.Number = 70 Then '�ļ����ڴ򿪵�״̬
        MsgBox "�ļ����ڴ򿪵�״̬"
    Else
        MsgBox "�����쳣����"
    End If
    EnEvents
    Err.Clear
End Sub
