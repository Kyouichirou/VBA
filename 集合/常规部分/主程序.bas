Attribute VB_Name = "������"
Option Explicit
'�ļ��������ļ�����ò�Ҫ������ansi�ַ�, �����ļ���������ansi�ַ����������鷳
Option Compare Text                       '�����ִ�Сд
Dim arrfiles(1 To 10000) As String                 '����һ������������Դ��path����(��ֵ�ɱ�)
Dim arrbase(1 To 10000) As String                   '�洢�ļ���
Dim arrextension(1 To 10000) As String             '�洢�ļ���չ��
Dim arrsize(1 To 10000) As String                   '�洢�ļ��Ĵ�С
Dim arrparent(1 To 10000) As String               '�洢�ļ�����λ��
Dim arrdate(1 To 10000) As String                  '�洢�ļ���������
Dim arrsizeb(1 To 10000) As Long                  '�ļ��Ĵ�С,��λ����
Dim arrmd5(1 To 10000) As String                    '�ļ�md5
Dim arredit(1 To 10000) As String                     '�ļ��޸�ʱ��
Dim arrcode(1 To 10000) As String                    '�ļ�·���쳣�ַ�
Dim arrcm(1 To 10000) As String                       '��ע
Dim arrfnansi(1 To 10000) As String                 '�ļ�����ansi��ע
Dim arrfpansi(1 To 10000) As String                  '�ļ�λ�÷�ansi��ע
Dim arrfilen() As Variant, arrfilesize() As Variant, arrfilemody() As Variant, arrfilemd5() As Variant, arrfilep() As Variant '�������,���ڱȽ�

Public fso As New FileSystemObject '------------------------------------------------------����
Public ShellxExist As Byte '�ж�Powershell�Ƿ����
Public AddFx As Byte '����ļ���Ӷ�����ִ��

Dim flc As Integer   '����ļ�������                              '% integer�����ݱ�ʶ��
Dim dl As Integer 'ɾ���ļ�������
Dim ls As Integer   '�����ļ�ͳ��                          'integer�������ݵķ�Χ-32768 �� 32767
Dim fc As Long '�ļ�����
Dim b As Byte, a As Integer 'Ŀ¼�����к�

Dim md5x As Integer
Dim umd5x As Integer
Dim idele As Integer 'ͳ��ɾ���˶���������
Dim ix As Integer
Dim deledic As New Dictionary '�洢ɾ�����ļ����ڵ��к�,����ɾ��
Dim rnglists As Range
Dim F As Byte '�����ж��ļ����Ƿ�����ӵ�״̬
Dim xi As Byte, c As Byte 'x���ڱ�����ݵĴ���, c����ļ��еĲ㼶
Dim Elow As Integer '������ݵ�����к�

Function ListAllFiles(ByVal addcode As Byte, ByVal FilePath As String) As Boolean               'addcode��ʾ���ݸ��µķ�ʽ,0ΪĬ�����,1Ϊ�������ļ���,2Ϊ���������ļ���
    Dim fd As Folder, fdatt As Long
    Dim i As Integer, j As Integer
    Dim tl As Single, t As Single '��¼ʱ��
    Dim rnglist As Range 'Ŀ¼
    Dim strp As String, strptemp As String, strfolder As String
    Dim bc As Integer, cN As Integer, tracenum As Byte
    Dim blow As Integer, cright As Byte 'Ŀ¼�������ݵķֲ������һ��
    Dim filec As Integer, ifilec As Integer, clm As Integer, bcx As Integer
    Dim xtemp As Variant, alow As Integer
    
    If addcode = 0 And FilePath = "NU" Then
        With Application.FileDialog(msoFileDialogFolderPicker) '�ļ���ѡ�񴰿�(�ļ����Դ�ֻ��ѡһ��)
            .Show
            If .SelectedItems.Count = 0 Then ListAllFiles = False: Exit Function 'δѡ���ļ������˳�sub
            strfolder = .SelectedItems(1)
        End With
        FilePath = strfolder '& "\" '��Ҫ������ļ���·��
    End If
    ListAllFiles = False
'    If CheckFileX(strfolder) = False And addcode = 0 Then MsgShow "�ļ��в�����Ŀ���ļ�", "Warning", 1500: Exit function '����cmd���ļ��н��п��ټ��,�ж��ļ������Ƿ����Ŀ���ļ�
    If CheckFileFrom(strfolder, 2) = True Then '����ļ��е���Դ,������ϵͳ�̵�λ��,ֻ����document,download,desktop����λ�õ��ļ������
        MsgBox "ϵͳ�̵��ļ�ֻ�����������:" & vbCr & "Desktop" & vbCr & "Downloads" & vbCr & "Documents"
        Exit Function
    End If
    If CheckPathAsWorkbook(strfolder) = True Then MsgBox "�ļ������λ������", vbInformation, "Tips": Exit Function '�������Ա����������������ļ�
    Set fd = fso.GetFolder(FilePath)          '��fdָ��·������
        If fd.IsRootFolder Then                    '��ֹ�������̷����
        MsgBox "�����̷����,����ʱ�����", vbOKOnly, "Careful!!!"
        Set fd = Nothing
        Exit Function
    End If
    fdatt = fd.Attributes
    If fd.ParentFolder.Path = Environ("SYSTEMDRIVE") & "\" Or fdatt = 18 Or fdatt = 1046 Then
    '------------------------��ֹ��������ļ���,ϵͳ�̵�һ���ļ���'��ֹ���ϵͳ�̵�һ���ļ���,��ֹ���ֲ����ļ����޷����ʵ�����,��Ŀ���ļ���
        Set fd = Nothing
        MsgBox "��ֹ�������ϵͳ�ļ���/�����ļ���", vbCritical, "Warning!!!"
        Exit Function
    End If
    
    ifilec = fd.Files.Count
    If ifilec = 0 And fd.SubFolders.Count = 0 Then
        Set fd = Nothing
        MsgBox "��ӵ��ļ���Ϊ��", vbOKOnly, "Careful!"
        Exit Function
    End If
'-------------------------------------------------------------------------------------------------����ļ��еĻ������
    With ThisWorkbook
    
        DisEvents '��ֹ����
        ' ��Ҫע����ִ�и��ӵĽ��̵�ʱ��,��Ҫ���Ǹ������ױ��������¼�,����ʱ�г�ͻ�Ľ��̻����ظ���ִ�еĽ���,�Լ��ٲ���Ҫ���¼��˷�
'        UserForm6.Show 0 'ע�����ｻ��Ϊ0��ʱ������Ĵ�����Լ���ִ��,Ϊ1ʱ,����ᱻ�ж�
        
        t = Timer '�����������ʱ��ĳ�ʼֵ
        
        flc = 0 '���̼�/ģ�鼶�����ĳ�ʼ�� '�漰�����ֱ�������������,���ǽ��е��Ի�����������,�������ֱ������Գ��ڴ���
        dl = 0
        ls = 0
        fc = 0
        b = 2 'Ŀ¼'λ��
        a = 0
        F = 0
        xi = 0
        c = 0
        ix = 0
        AddFx = 0
        idele = 0
        md5x = 0
        umd5x = 0
        Elow = 0 '�����ĳ�ʼ��
        c = UBound(Split(FilePath, "\")) '�ļ���λ�ڴ��̵�Ŀ¼�ȼ�
        
        If PSexist = True Then ShellxExist = 1 '��Ҫ�ж�powershell�汾����Ϣ,���ܵ���4.0
        
        DataTreat '�������еı������
        
        With .Sheets("Ŀ¼")
            blow = .[b65536].End(xlUp).Row
            cright = .Cells.SpecialCells(xlCellTypeLastCell).Column '�������Ҳ�
            If Len(.Range("b4").Value) = 0 Then '��δд������
                a = blow + 1
                xi = 1    '���ڱ�Ǳ���Ƿ��Ѿ���������
                GoTo 107
            Else
                If c = 1 Then                                          '�����δд������/һ��Ŀ¼�����ڹ���Ŀ¼'�����ж�����ӽ������ļ��е���Ϣ����Ҫ�ж�ֱ��д������
                    strp = FilePath & "\" '����"\"������Ϊ�˷�ֹc:\a,c:\ab��������ĳ���,��ģ���������޷�׼ȷ�ҵ�Ŀ��
                Else
                    xtemp = Split(FilePath, "\")
                    strp = xtemp(0) & "\" & xtemp(1) & "\" 'һ��Ŀ¼
                End If
                Set rnglist = .Cells(4, 3).Resize(blow, cright).Find(strp, lookat:=xlPart) 'ͨ�����һ���ļ���,�ж��Ƿ���ڹ����ļ���
                If rnglist Is Nothing Then
                    xi = 1 '����ļ�����δ�����ļ��б���ӽ��� ,Xi=1 ��ʾ����Ҫִ�бȽ�
                    a = blow + 1
                    GoTo 107
                End If
                '------------------------------------------------------------------------------------------------����Ƿ��Ѿ�������ص��ļ�����ӽ���
                strp = FilePath & "\"
                Set rnglist = .Cells(4, 3).Resize(blow, cright).Find(strp, lookat:=xlWhole) '��ȫƥ�����
                
                If Not rnglist Is Nothing Then '��ֵ '�ⲿ�ִ������ļ���λ��Ŀ¼,��ӽ�����Ϊ���ļ��� ''�������ݵ������
                    F = 1
                    filec = rnglist.Offset(0, 4).Value 'Ŀǰ�ļ��е��ļ�����
                    If addcode = 1 Then     '���µķ�ʽ��ͬ ,1,2�ֱ��ʾ�Ƿ�������ļ���
                        a = rnglist.Row
                        If fd.DateLastModified = rnglist.Offset(0, 2) Then
                            F = 3 '�����һ����ļ�������
                            GoTo 109
                        Else
                            If Int(Abs(filec - ifilec) / filec) > 50 Then '�ļ��з����仯��Χ�ķ����㹻�� 'abs����ֵ����
                                F = 4
                                GoTo 107
                            End If
                        End If
                    ElseIf addcode = 2 Then
                        If fd.DateLastModified = rnglist.Offset(0, 2) Then
                            GoTo 1001 '�ļ��е�����û�з����仯(ֻ�����ļ���һ��)'���������ļ���
                        Else
                            a = rnglist.Row
                            GoTo 107
                        End If
                    End If
                    MsgBox "���ļ������" '-�Ѿ���ӵ��ļ��н�ֹ������ӵķ�ʽ���
                    GoTo 1001
                Else
                    clm = cright - 5 '���ݵ����Ҳ�
                    bc = b + c
                    Do
                        strp = fd.Path
                        Set fd = fd.ParentFolder
                        strptemp = fd.Path & "\"
                        Set rnglist = .Cells(4, 3).Resize(blow, cright).Find(strptemp, lookat:=xlWhole) '����֮��Ϊ�����ļ���
                        If Not rnglist Is Nothing Then
                            F = 2
                            a = rnglist.Row + 1
                            .Cells(a, 1).EntireRow.Insert
                            GoTo 104
                        End If
                        '----------------------------------------------------------------------------------����ȷ���ļ����ڱ���λ��,ȷ��������ӽ����ĸ����ļ���֮���ܹ������ڵ�λ��
                        strp = strp & "\"
                        For bcx = bc To clm '����ѭ��������ǿ����������˳�����ν�������(Ĭ�ϵ�����˳���ǲ�ȷ����)
                            Set rnglist = .Cells(4, bcx).Resize(blow, 1).Find(strp, after:=.Cells(4, bcx), lookat:=xlPart)
                            If Not rnglist Is Nothing Then
                                F = 2
                                If c = 1 Then                        '��֤1���ļ���λ������ĵ�һ��
                                    For cN = rnglist.Row To 3 Step -1
                                        If .Cells(cN, 1) <> .Cells(rnglist.Row, 1) Then
                                            a = cN + 1
                                            .Cells(a, 1).EntireRow.Insert
                                            GoTo 104
                                        End If
                                    Next
                                Else
                                    If CInt(.Cells(rnglist.Row, 2)) = 1 Then '����һ���ļ�
                                        a = rnglist.Row + 1
                                    Else
                                        a = rnglist.Row
                                    End If
                                    .Cells(a, 1).EntireRow.Insert '�ǹ����ĸ����ļ���һ�ɷ�����һ��
                                End If
                                GoTo 104
                            End If
                        Next
                        bc = bc - 1          '�ļ��㼶�ı仯
                    Loop Until fd.IsRootFolder
                    a = blow + 1 '���û�ҵ�
            End If
104
            Set fd = fso.GetFolder(FilePath)          '���½�fdָ��·������
107
        End If
            .Cells(a, 2) = c
            .Cells(a, 2).NumberFormatLocal = "0_);[��ɫ](0)"
            .Cells(a, b + c) = fd.Path & "\"
            .Cells(a, b + c + 1) = fd.DateCreated
            .Cells(a, b + c + 1).NumberFormatLocal = "yyyy/m/d h:mm;@" '��ʽ����
            If F > 0 Then
                .Cells(a, 1) = .Cells(rnglist.Row, 1)
            ElseIf F = 0 Then
                alow = .[a65536].End(xlUp).Row
                .Range("a" & alow).AutoFill Destination:=.Range("a" & alow & ":" & "a" & a), type:=xlFillDefault
            End If
            .Cells(a, b + c + 2) = fd.DateLastModified
            .Cells(a, b + c + 2).NumberFormatLocal = "yyyy/m/d h:mm;@"
            .Cells(a, b + c + 3) = fd.SubFolders.Count
            .Cells(a, b + c + 3).NumberFormatLocal = "0_);[��ɫ](0)" '���õ�Ԫ��ĸ�ʽ,����ᱻExcelת��Ϊ�����ĸ�ʽ
            .Cells(a, b + c + 4) = fd.Files.Count
            .Cells(a, b + c + 4).NumberFormatLocal = "0_);[��ɫ](0)"
            .Cells(a, b + c + 5) = fd.Size
            .Cells(a, b + c + 5).NumberFormatLocal = "0_);[��ɫ](0)"
        End With
'------------------------------------------------------------------------------------------------------------------------------------------�ļ�����Ŀ¼��λ�õ�ȷ��,�������ݵķ�ʽ
109
        c = c + 1
        SearchFolders fd, addcode                                   '����sf��sub�������ļ��кͻ�ȡ�ļ������ļ�����Ϣ
        
        If flc = 0 Then
            If addcode > 0 Then
                GoTo 1001
            Else
                GoTo 1002 '��û�з���ֵ��ʱ��
            End If
        End If
            
        Call WriteData(2) 'д������
    
1002
        With .Sheets("������")                     '��������д����Ϣ����(��Ҫ�޸�)
            If Len(.Range("e37").Value) = 0 Then            '��δд������
                .Range("e37") = FilePath
                .Range("i37") = Now
            Else
                j = .[e65536].End(xlUp).Row
                Set rnglist = .Range("e37:e" & j).Find(FilePath, lookat:=xlWhole)
                If Not rnglist Is Nothing Then
                    GoTo 1001
                Else
                    .Range("e" & j + 1) = FilePath
                    .Range("i" & j + 1) = Now
                End If '
            End If
        End With
    End With
    
1001
    
    Set fd = Nothing
    Set rnglist = Nothing
    Set rnglists = Nothing
    ListAllFiles = True
    EnEvents
    
'    Unload UserForm6
'    tl = Timer - t
'    If tl > 30 Then
'        MsgBox "�������,����ʱ��: " & Format(tl, "0.0000") & "s" & vbCr _
'        & "�ܹ�����: " & fc & "���ļ�" & vbCr _
'        & "�ɹ�����: " & flc & "���ļ�" & vbCr _
'        & "���ֿ����ص��ļ�: " & ls & "��" & vbCr _
'        & "ɾ���ص��ļ�" & dl & "��"
'    Else
'        MsgBox "�������! " & vbCr _
'        & "�ܹ�����: " & fc & "���ļ�" & vbCr _
'        & "�ɹ�����: " & flc & "���ļ�" & vbCr _
'        & "���ֿ����ص��ļ�: " & ls & "��" & vbCr _
'        & "ɾ���ص��ļ�" & dl & "��"
'    End If
End Function

Private Sub DataTreat(Optional addcode As Byte) '���ɱ�������
    Dim itemp As Integer
    With ThisWorkbook.Sheets("���")               '���ⲿ�ֽ����ں���ӽ������ļ����бȽ�,�����find,���ٶ����и��õ�����,��Ҫ�����������ǳ����ʱ������ڴ��ռ��,�������㹻���ʱ��,��find���,�ٶȵ����ƿ�ʼ������С
        Elow = .[c65536].End(xlUp).Row + 1
        If Elow < 5 Then MsgBox "���ṹ���ƻ�", vbCritical, "Warning!!!":  Exit Sub
        If Elow = 6 Then Exit Sub
        itemp = Elow - 6
        If itemp = 1 Then
            If fso.fileexists(.Cells(6, "e").Value) = False Then   '����ļ��ر��ٵ�ʱ�������һ���ԭ�е��ļ��Ƿ񻹴�����Ŀ¼��Ӧ��λ��
                .rows(6).Delete Shift:=xlShiftUp '������ݱ�ȫ�����
                Elow = 6 '�µ�λ��
                If addcode = 0 Then
                    ClearAll (0)
                Else
                    ClearAll (1) '���µķ�ʽ-���Ŀ¼����
                End If
                Exit Sub
            End If
            ReDim arrfilen(1 To 1, 1 To 1)
            ReDim arrfilesize(1 To 1, 1 To 1) 'redim�����޸��������������
            ReDim arrfilemody(1 To 1, 1 To 1)
            ReDim arrfilemd5(1 To 1, 1 To 1)
            ReDim arrfilep(1 To 1, 1 To 1)
            arrfilen(1, 1) = .Cells(6, "c").Value
            arrfilesize(1, 1) = .Cells(6, "g").Value
            arrfilemody(1, 1) = .Cells(6, "h").Value
            arrfilemd5(1, 1) = .Cells(6, "k").Value
            arrfilep(1, 1) = .Cells(6, "e").Value
        ElseIf itemp > 1 Then
            itemp = Elow - 1
            arrfilen = .Range("c6:c" & itemp).Value
            arrfilesize = .Range("g6:g" & itemp).Value '��Ҫʹ��forѭ��ȡֵ,̫��.application.transpose������,�ڴ�����34000+������ʱ,��������
            arrfilemody = .Range("h6:h" & itemp).Value
            arrfilemd5 = .Range("k6:k" & itemp).Value
            arrfilep = .Range("e6:e" & itemp).Value
        End If
    End With
End Sub

Private Sub WriteData(ByVal xi As Byte) '����ȡ��������д�뵽���,�������������
    Dim itemp2 As Integer, i As Integer, k As Integer
    Dim arrdele() As Integer, arrdelex() As Integer
    
    Elow = Elow - idele
    With ThisWorkbook.Sheets("���")                       'Ĭ�ϵ�һά������һ��,transpose�������Խ�����ת������ʽ����,
    '----------------------------------------------ע��Application.Transposeת�õ����ݵ�����,��ת�ñ��������ʱ������Ϊ34446����(��ͬ���豸����office��������������ͬ)
        .Range("k" & Elow).Resize(flc) = ThisWorkbook.Application.Transpose(arrmd5)  '�ļ�md5
        .Range("ab" & Elow).Resize(flc) = ThisWorkbook.Application.Transpose(arrcode) '�쳣�ַ����
        .Range("ac" & Elow).Resize(flc) = ThisWorkbook.Application.Transpose(arrcm)   '��ע
        .Range("ae" & Elow).Resize(flc) = ThisWorkbook.Application.Transpose(arrfnansi)
        .Range("af" & Elow).Resize(flc) = ThisWorkbook.Application.Transpose(arrfpansi)
        .Range("c" & Elow).Resize(flc) = ThisWorkbook.Application.Transpose(arrbase) '�ļ���
        .Range("d" & Elow).Resize(flc) = ThisWorkbook.Application.Transpose(arrextension) '�ļ���չ��
        .Range("e" & Elow).Resize(flc) = ThisWorkbook.Application.Transpose(arrfiles) '�ļ�·��
        .Range("f" & Elow).Resize(flc) = ThisWorkbook.Application.Transpose(arrparent) '�ļ�����λ��
        .Range("g" & Elow).Resize(flc) = ThisWorkbook.Application.Transpose(arrsizeb) '�ļ���ʼ��С
        .Range("h" & Elow).Resize(flc) = ThisWorkbook.Application.Transpose(arredit) '�ļ��޸�ʱ��
        .Range("i" & Elow).Resize(flc) = ThisWorkbook.Application.Transpose(arrsize) '�ļ���С
        .Range("j" & Elow).Resize(flc) = ThisWorkbook.Application.Transpose(arrdate) '�ļ�����ʱ��
        itemp2 = flc + Elow - 1
        .Range("x" & Elow & ":" & "x" & itemp2) = Now '���Ŀ¼��ʱ��
        .Range("ad" & Elow & ":" & "ad" & itemp2) = xi '��ע�ļ�����Դ(��ͨ������ļ��еķ�ʽ��ӽ�����)
        .Range("b" & Elow) = .Cells(5, 1).Value
        itemp2 = itemp2 + 1
        .Range("b" & Elow).AutoFill Destination:=.Range("b" & Elow & ":" & "b" & itemp2), type:=xlFillDefault '���ͳһ����
        .Cells(5, 1).Value = .Range("b" & itemp2) 'ȷ�����е�ֵ������Ψһ��
        .Range("b" & itemp2).ClearContents
        If UF3Show = 1 Or UF3Show = 3 Then AddFx = 1 '����������ڴ������ݸ���ʹ��
        Erase arrfiles '�������,��̬���������Ͷ�̬������������΢�Ĳ��,��̬���齫����ȫĨ��,��̬������Ȼ���Ա�����Χ,ֵ��Ĩ��
        Erase arrbase
        Erase arrextension
        Erase arrsize
        Erase arrparent
        Erase arrdate
        Erase arrsizeb
        Erase arrmd5
        Erase arredit
        Erase arrcode
        Erase arrcm
        Erase arrfnansi
        Erase arrfpansi
        Erase arrfilen
        Erase arrfilesize
        Erase arrfilemody
        Erase arrfilep
        Erase arrfilemd5
        If idele > 0 Then
            i = deledic.Count - 1
            ReDim arrdele(i)
            ReDim arrdelex(i)
            For k = 0 To i
                arrdele(k) = deledic.Keys(k)
            Next
            arrdelex = Down(arrdele) '��������
            For k = 0 To i
                .rows(arrdelex(k)).Delete Shift:=xlShiftUp
                DeleFileOver .Cells(arrdelex(k), 6).Value
            Next
    '        deledic.RemoveAll
            Erase arrdele
            Erase arrdelex
            Set deledic = Nothing
        End If
    End With
End Sub

Private Sub SearchFolders(ByVal fd As Folder, ByVal addcodex As Integer)              'ByVal,Ϊ��ֵ���ݷ�ʽ, �б���byref(�����÷�ʽ����)
    Dim flx As File, fdatt As Long
    Dim sfd As Folder
    Dim strTemp As String
    Dim bc As Integer, bcx As Integer, clm As Integer
    Dim blow As Integer, cright As Integer 'Ŀ¼�������ݵķֲ������һ��
    Dim filec As Integer, ifilec As Integer, strp As String
    
    ix = ix + 1 '���ڱ���ļ���
    If F = 3 Then GoTo 1007 '���ļ��и��µ�ʱ��,��һ���ļ���û�з����仯,ֱ�ӵ����鿴���ļ���/�����ļ�����û���ļ�
    For Each flx In fd.Files                   '�����ļ�
        fc = fc + 1 'ͳ���ļ�����
        FileIn flx
    Next flx
1007
    If addcodex = 2 Then Exit Sub 'ֻ�����ļ���һ�㲻�漰���ļ���
    
    If fd.SubFolders.Count = 0 Then Exit Sub  '���ļ�����ĿΪ�����˳�sub
    
    For Each sfd In fd.SubFolders             '�������ļ���
        ifilec = sfd.Files.Count
        If ifilec = 0 And sfd.SubFolders.Count = 0 Then GoTo 107 'û���ļ�,���ļ���Ϊ��
        fdatt = sfd.Attributes
        If fdatt = 18 Or fdatt = 1046 Then GoTo 107 '���˵������ļ���
        F = 0                          'fΪģ�����,ʹ�ú�����
        With ThisWorkbook.Sheets("Ŀ¼")
            blow = .[b65536].End(xlUp).Row
            cright = .Cells.SpecialCells(xlCellTypeLastCell).Column '��������Ҫ�ظ�����ʱ,�����ظ�ȡֵ
            If xi = 1 Then
                a = blow + 1 '����Ҫִ�бȽ�
                GoTo 109
            End If
            strTemp = sfd.Path & "\"
            bc = b + c
            clm = cright - 5
            Set rnglists = .Cells(4, 3).Resize(blow, cright).Find(strTemp, after:=.Cells(4, 3), searchorder:=xlByColumns, lookat:=xlWhole)
            If Not rnglists Is Nothing Then '���Ŀ¼�Ѿ�����
                F = 1
                a = rnglists.Row
                If sfd.DateLastModified <> rnglists.Offset(0, 2) Then '�ļ����޸�ʱ���Ѿ������仯
                    filec = Rng.Offset(0, 4).Value
                    If filec = 0 Then F = 0: GoTo 109 'ԭ�����ļ���û������
                    If 100 * Int(Abs(ifilec - filec) / filec) > 50 Then F = 4 '�ļ��е������仯��Χ intת��Ϊinteger��,abs,����ֵ,���ļ��е����ݷ�����ı仯ʱ
                    GoTo 109
                Else
                    F = 3       'ʱ����ͬ,�������ļ�
                    If sfd.SubFolders.Count = 0 Then
                        GoTo 107
                    Else
                        GoTo 106
                    End If
                End If
            Else
                Do
                    strp = sfd.Path & "\"
                    Set sfd = sfd.ParentFolder
                    Set rnglists = Nothing
                    For bcx = bc To clm
                        Set rnglists = .Cells(4, bcx).Resize(blow, 1).Find(strp, after:=.Cells(4, bcx), lookat:=xlPart)
                        '-----------------��ģ��������ʱ��,������λ�ò���һ���ϸ�������ָ��������ʼ,ֻ��ͨ��ѭ���ķ�ʽ,ǿ������ÿһ��
                        If Not rnglists Is Nothing Then
                            If CInt(.Cells(rnglists.Row, 2)) >= c Then
                                a = rnglists.Row
                            Else
                                a = rnglists.Row + 1
                            End If
                            .Cells(a, 1).EntireRow.Insert
                            GoTo 110
                        End If
                    Next
                    bc = bc - 1 '�ļ��в㼶����
                Loop Until sfd.IsRootFolder
                a = blow + 1
            End If
    '-------------------------------------------------------------------------------------------------------------------------------����ȷ�����ļ�����Ŀ¼���ֵ�λ��
110
            Set sfd = fso.GetFolder(strTemp) '���¸�ֵ
109
            .Cells(a, 1) = .Cells(a - 1, 1)
            .Cells(a, 2) = c
            .Cells(a, 2).NumberFormatLocal = "0_);[��ɫ](0)"
            .Cells(a, b + c) = sfd.Path & "\"
            .Cells(a, b + c + 1) = sfd.DateCreated
            .Cells(a, b + c + 1).NumberFormatLocal = "yyyy/m/d h:mm;@"
            .Cells(a, b + c + 2) = sfd.DateLastModified
            .Cells(a, b + c + 2).NumberFormatLocal = "yyyy/m/d h:mm;@"
            .Cells(a, b + c + 3) = sfd.SubFolders.Count
            .Cells(a, b + c + 3).NumberFormatLocal = "0_);[��ɫ](0)"
            .Cells(a, b + c + 4) = ifilec
            .Cells(a, b + c + 4).NumberFormatLocal = "0_);[��ɫ](0)"
            .Cells(a, b + c + 5) = sfd.Size
            .Cells(a, b + c + 5).NumberFormatLocal = "0_);[��ɫ](0)"
        End With
106
        a = a + 1
        If sfd.SubFolders.Count > 0 Then
        c = c + 1 '����ļ��Ĳ㼶
        End If
108
        SearchFolders sfd, addcodex
107
    Next
    c = c - 1 'ÿһ�ζ���������ļ����µ�һ���ļ����µ����һ�����ļ���,��ִ�����֮��,���¼��㿪ʼ����һ��,����cҪ���еݼ�
End Sub

Sub AddFile() '--------------------------------------------����ļ�
    Dim fdx As FileDialog, fl As File, fd As Folder
    Dim selectfile As Variant
    Dim rngx As Range, strfd As String
    Dim i As Byte, k As Byte
    
    Set fdx = ThisWorkbook.Application.FileDialog(msoFileDialogFilePicker)
    On Error GoTo 100
    With fdx
        .AllowMultiSelect = True '����ѡ�����ļ�(ע�ⲻ���ļ���,�ļ���ֻ��ѡһ��)
        .Show
        i = .SelectedItems.Count
        If i = 0 Then Exit Sub
        If i > 10 Then MsgBox "����һ�����10���ļ�", vbOKOnly, "Careful!": Exit Sub
        ix = 0    'ģ�鼶���������
        AddFx = 0
        idele = 0
        dl = 0
        ls = 0
        md5x = 0
        umd5x = 0
        Elow = 0
        flc = 0
    
        DisEvents
        DataTreat
        '-------------------׼������
        For Each selectfile In .SelectedItems
            If k = 0 Then
                If CheckFileFrom(selectfile, 1) = True Or CheckPathAsWorkbook(selectfile, 1) = True Then Exit Sub '�����ļ�����Դ
            End If
            k = k + 1
            Set fl = fso.GetFile(selectfile)
            FileIn fl
        Next
    End With
    '--------------------------------------------------------------------------------------�ļ������
    If flc = 0 Then GoTo 100
    Call WriteData(1)
    strfd = fl.ParentFolder
    Set fd = fso.GetFolder(strfd)
    strfd = strfd & "\"
    With ThisWorkbook.Sheets("Ŀ¼") '��ӵ��ļ�,˵�������µ��ļ�,��ô��Ҫ���������ļ��е���Ϣ
        Set rngx = .Cells(4, 3).Resize(.[b65536].End(xlUp).Row, .Cells.SpecialCells(xlCellTypeLastCell).Column).Find(strfd, after:=.Cells(4, 3), searchorder:=xlByColumns, lookat:=xlWhole)
    End With
    With fd
        If Not rngx Is Nothing Then rngx.Offset(0, 2) = .DateLastModified: rngx.Offset(0, 4) = .Files: rngx.Offset(0, 5) = .Size
    End With
    '------------------------------------------------------------------------------------------------------------��Ϣд��Ŀ¼/����
100
    Set fdx = Nothing
    Set fl = Nothing
    Set fd = Nothing
    Set rngx = Nothing
    EnEvents
End Sub

Private Function FileIn(ByVal fl As File)
    '--------------------------------------�ļ���Ϣ���� 'fΪ0��ʱ��,��ʾ���ȫ�µ��ļ���,��ִ�еķ�ʽΪȫ������,f=1��ʱ���ʾ�����ļ���,
    '--------------------------------------���ָ���,f=2��ʱ��,�������еĹ����ļ���,ȫ������,f=3��ʱ��,�����ļ���,�޸�ʱ���Ŀ¼һֱ,ֱ������
    Dim filefd As String '�ļ������ļ���
    Dim filen As String '�ļ���
    Dim filemd5 As String '�ļ�md5
    Dim filex As String '�ļ���չ��
    Dim filez As Long '�ļ���С �����2G���ļ�
    Dim filep As String '�ļ�·��
    Dim filem As String '�ļ��޸�ʱ��
    Dim filect As String '�ļ�����ʱ��
    Dim p As Integer           '�Ƿ���������ַ�
    Dim k As Byte '�ж��ļ�����
    Dim md5t As Byte
    Dim j As Integer
    
    '--------------------------------------------------------- '���ļ��Ļ������Խ��л�ȡ�ʹ���-----'10-20-30-40,����������г���,����ê������ĳ��ֵ�����
    On Error GoTo 100
10
    With fl
        filep = .Path '·��
        filex = fso.GetExtensionName(filep)
        filex = LCase(filex) '��չ��'������ӽ�Ŀ¼���ļ����� ' l/ucaseΪת����С���� ,����ͳһʹ��Сд
        If filex Like "epub" Or filex Like "pdf" Or filex Like "mobi" Then
            k = 1
        ElseIf filex Like "do*" Or filex Like "xl*" Or filex Like "pp*" Or filex Like "tx*" Or filex Like "ac*" Then
            k = 2
        End If
        If k = 0 Then Exit Function
        filez = .Size '��С
        If Len(filex) = 0 Or filez = 0 Or fl.Attributes = 34 Then Exit Function '�ļ���չ��Ϊ��/�����ļ������� ,34��ʾhidden����,ע�ⲻҪֱ��ʹ��hidden����ʾ����,�޷�ʶ��
        filen = .Name '�ļ���
        filefd = .ParentFolder '�ļ���
        filect = .DateCreated '����ʱ��
        filem = .DateLastModified '�޸�ʱ��
    End With
    
    '-------------------------------------------------------------------���ļ�����Ϣ���бȽϺʹ���
20
    p = ErrCode(filen, 0, filefd)  '���з�ansi�ַ����
    If p = -1 Then GoTo 100 'δ��ȡ����Ч·��(��ȡ�����쳣)
    
    If Elow = 6 Then '�����δд������/����Ҫ���кͱ������ݱȽ�
        If k = 1 Then
            filemd5 = GetFileHashMD5(filep, p) '����md5
            If Len(filemd5) = 2 Then
                If md5x > 1 Then md5t = 1 '�����ļ����Ƚ�
            End If
            md5x = md5x + 1
        Else
            umd5x = umd5x + 1
        End If
    Else '�Ѵ�������
        If F = 1 Or F = 4 Then '�Ѹ��µķ�ʽ���Ƚ��ļ�,�Ƚ����ļ����Ƚ�,֮���ٽ���md5�Ƚ�
            If FileComp(filen, filep, filemd5, filez, filem, 3) = 5 Then Exit Function '�ļ�����ͬ,����'����ļ����Ѿ�������Ŀ¼,��ô�Ȳ�����md5,�Ƚ����ļ����Ƚ�
        End If
        If k = 1 Then
            filemd5 = GetFileHashMD5(filep, p) '����md5
            If Len(filemd5) = 2 Then            '���������Чmd5
                If FileComp(filen, filep, filemd5, filez, filem, 2) = 3 Then ls = ls + 1: Exit Function '��������ݽ��бȽ�
            Else
                If FileComp(filen, filep, filemd5, filez, filem, 1) = 1 Then Call DeleRFile(filep, filex, p): Exit Function '�ͱ�������һ��,ɾ��
            End If
            md5x = md5x + 1
        Else
            If FileComp(filen, filep, filemd5, filez, filem, 2) = 3 Then ls = ls + 1: Exit Function '��������ݽ��бȽ�
            umd5x = umd5x + 1
        End If
    End If
    
    '----------------------------- '������е��ļ��Ƚ� -�Ⱥͱ������ݱȽ�,�ڽ�������е��ļ��Ƚ�
30
    If md5x > 1 Then
        For j = 1 To flc '�����ôʵ�dic.exist��ȡ��
            If filemd5 = arrmd5(j) Then Call DeleRFile(filep, filex, p): Exit Function
        Next
    End If
    If umd5x > 1 Or md5t = 1 Then 'md5t��ʾ�޷�������Чmd5
        For j = 1 To flc
            If filen = arrbase(j) Then
                If filez = arrsizeb(j) Then
                    If filem = arredit(j) Then GoTo 100 '����
                End If
            End If
        Next
    End If
    '-----------------------------------------------------------------------------------------����д��
40
    flc = flc + 1
    arrmd5(flc) = filemd5
    arrfiles(flc) = filep                  '����洢�ļ���·��
    arrbase(flc) = filen                   '������չ�����ļ���
    arrextension(flc) = filex                    '�ļ���չ��
    arrparent(flc) = filefd '�ļ���һ��Ŀ¼
    arrdate(flc) = filect '�ļ�����
    arrsizeb(flc) = filez         '�ļ���С
    arredit(flc) = filem '�ļ��޸�ʱ��
    arrcode(flc) = IIf(p > 1, "ERC", "")              'https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/iif-function
    If Tagfnansi = True And Tagfpansi = True Then 'd
        arrfnansi(flc) = errcodenx
        arrfpansi(flc) = errcodepx
        arrcm(flc) = "EDC"
    ElseIf Tagfnansi = True And Tagfpansi = False Then 's
        arrfnansi(flc) = errcodenx
        arrfpansi(flc) = ""
        arrcm(flc) = "ENC" '���ֵ�λ��
    ElseIf Tagfnansi = False And Tagfpansi = True Then 's
        arrfnansi(flc) = ""
        arrfpansi(flc) = errcodepx
        arrcm(flc) = "EPC" '���ֵ�λ��
    ElseIf Tagfnansi = False And Tagfpansi = False Then
        arrfnansi(flc) = ""
        arrfpansi(flc) = ""
        arrcm(flc) = ""
    End If
    If filez < 1048576 Then
        arrsize(flc) = Format(filez / 1024, "0.00") & "KB"    '�ļ��ֽڴ���1048576��ʾ"MB",������ʾ"KB"
    Else
        arrsize(flc) = Format(filez / 1048576, "0.00") & "MB"
    End If
    Exit Function
100
    If Erl = 40 And flc > 1 Then flc = flc - 1
    Err.Clear
End Function

Private Function FileComp(ByVal filenx As String, ByVal filep As String, ByVal filemd5x As String, ByVal filezx As String, _
ByVal filemx As String, ByVal cmCode As Byte) As Byte '�ж��ļ��Ƿ�ͱ���Ѵ��ڵ�Ŀ¼�ص�
    Dim m As Integer, n As Integer
    Dim itemp As Integer
    
    itemp = Elow - 6
    With ThisWorkbook.Sheets("���")
        If cmCode = 1 Then
            For m = 1 To itemp
                If filemd5x = arrfilemd5(m, 1) Then 'md5�Ա�'�����ôʵ�dic.exist��ȡ�� ������Ҫע��ֻ��md5�������ݿ�����˴���,��Ϊ����Ψһ��
                    If fso.fileexists(arrfilep(m, 1)) = False Then 'ͬʱ����ļ��Ƿ����
                        n = m + 5
                        deledic(n) = ""
                        idele = idele + 1
                        FileComp = 0
                        Exit Function
                    Else
                        FileComp = 1
                    End If
                    Exit Function
                End If
            Next
            FileComp = 2 '������
        
        ElseIf cmCode = 2 Then
            For m = 1 To itemp
                If filenx = arrfilen(m, 1) Then
                    If filezx = arrfilesize(m, 1) Then
                        If filemx = arrfilemody(m, 1) Then
                            If fso.fileexists(arrfilep(m, 1)) = False Then 'ͬʱ����ļ��Ƿ����
                                n = m + 5
                                deledic(n) = "" '.Cells(n, 6).Value
                                idele = idele + 1
                                FileComp = 0
                                Exit Function
                            Else
                                FileComp = 3
                                Exit Function
                            End If
                        End If
                    End If
                End If
            Next
            FileComp = 4
        
        ElseIf cmCode = 3 Then
            For m = 1 To itemp
                If filep = arrfilep(m, 1) Then '���ļ����Ѿ����,��ִ������·���ıȽ�,����ļ����ڵ��ļ���������ı仯�Ͳ�����ļ��Ƿ����
                    If F = 4 Then '�ļ����ڵ��ļ��Ѿ������˴�ı仯
                        If fso.fileexists(arrfilep(m, 1)) = False Then 'ͬʱ����ļ��Ƿ����
                            n = m + 5
                            deledic(n) = ""
                            idele = idele + 1
                            FileComp = 0
                            Exit Function
                        End If
                    End If
                    FileComp = 5
                    Exit Function
                End If
            Next
            FileComp = 6
        End If
    End With
End Function

Private Function DeleRFile(ByVal filepx As String, ByVal filext As String, ByVal Px As Byte)  'ɾ���ظ����ļ�
    If FileTest(filepx, filext) = 0 Then '�ļ����ڹرյ�״̬
        If Px > 0 Then '���ڷ�ansi�ַ�
            fso.DeleteFile (filepx) '�ж��ļ��Ƿ��ڴ򿪵�״̬
        Else
            DeleteFiles (filepx) 'ɾ���ļ�
        End If
        With ThisWorkbook.Sheets("Ŀ¼")
            If ix = 1 Then '��һ���ļ���
                .Cells(a, b + c + 2) = Now 'ɾ���������ļ��е�ʱ�䷢���仯
            Else
                .Cells(a - 1, b + c + 1) = Now '��Ҫ�޸�
            End If
        End With
        dl = dl + 1 '�ۼ�ɾ���ļ�������
    End If
End Function

Private Function CheckPathAsWorkbook(ByVal strText, ByVal cmCode As Byte) As Boolean '�ж��ļ���/�ļ�����Դ�͹�������λ�ù�ϵ
    Dim WBPath As String
    
    WBPath = ThisWorkbook.Path & "\"
    CheckPathAsWorkbook = False
    If cmCode = 1 Then     '��ʾ�ļ�
        strText = fso.GetFile(strText).ParentFolder & "\"
    ElseIf cmCode = 2 Then
        strText = strText & "\"     '��ʾ�ļ���
    End If
    If Len(strText) > Len(WBPath) Then
        If InStr(strText, WBPath) > 0 Then CheckPathAsWorkbook = True
    Else
        If InStr(WBPath, strText) > 0 Then CheckPathAsWorkbook = True
    End If
End Function
