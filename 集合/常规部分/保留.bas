Attribute VB_Name = "����"
'Option Explicit
'
''Function Filein(ByVal fl As File)
''Dim filemd5 As String
''Dim rngmd As Range
''Dim FilePath As String
''Dim t As Integer       '���ڱ��hash�������ַ�ʽ���ɵ�
''Dim J As Integer        '���ڱ�Ƿǵ������ļ���·���Ƿ���������ַ�
''Dim p As Integer           '�Ƿ���������ַ�
''Dim flr As Variant        '���ڼ�¼������ҷ��ص�ֵ
''Dim faddress As String '���ڼ�¼ͬ���ļ��������ĵ�һ���ļ���λ��
''Dim ckn As Integer '���ڱ���ļ��Ѿ�������� 'ckn=0��ʾ�ļ�δ�����ļ������,=1��ʾ�������
''Dim exet As Integer
'
'If Len(fso.GetExtensionName(fl.Path)) = 0 Or fl.Size = 0 Or fl.Attributes = 34 Then GoTo 100 '�ļ���չ��Ϊ��/�����ļ������� ,34��ʾhidden����,ע�ⲻҪֱ��ʹ��hidden����ʾ����,�޷�ʶ��
'exet = fso.GetExtensionName(fl.Path) '������ӽ�Ŀ¼���ļ����� ' ucaseΪת����С����
'If Not exet Like "EPUB" And Not exet Like "PDF" And Not exet Like "MOBI" And Not exet Like "DO*" And Not exet Like "XL*" And Not exet Like "PP*" And Not exet Like "AC*" And Not exet Like "TX*" Then GoTo 100 '�ļ�����ɸѡ
'FilePath = fl.Path
'J = 0
't = 0
'p = Errcode(fl.ShortName, fl.ParentFolder, 0)
'
'If p > 0 Then
'    If exet Like "EPUB" Or exet Like "PDF" Or exet Like "MOBI" Then '�޶���pdf��mobi,epubʹ��md5 ,���ǵ��������ļ������ױ༭���ļ�,�ļ���hash��̬�仯
'        If f = 1 Then
'        t = 1
'        GoTo 1010
'        End If
'1011
'        If ShellxExist = 1 Then '�ж�powershell�Ƿ����, Environ("SystemRoot")ϵͳ��������(C:\windows,����ϵͳ��װ��C��)
'            DoEvents
'            t = 1
'            filemd5 = UCase(Hashpowershell(FilePath)) 'ucase,ת���ɴ�д
'            If Len(filemd5) = 2 Then GoTo 1002 'û�л�ȡ���ļ�hash
'        Else
'            If fl.Size < 20971520 Then '���ڲ���adodb.stream��ʽ����hash���ٶ�̫��,������Ҫ�����ļ��Ĵ�С, ������20M
'                DoEvents
'                t = 1
'                filemd5 = UCase(GetMD5Hash_File(FilePath, fl.Size))
'            Else
'1002
'                flc = flc + 1
'                t = 0
'                arrcode(flc) = "ERC"       '�쳣�ַ���� '�޸�-�޷�������Чֵ,�ͽ����ļ����Ƚ�
'                arrmd5(flc) = "UC"
'                If Tagfnansi = True And Tagfpansi = True Then 'd
'                    arrfnansi(flc) = errcodenx
'                    arrfpansi(flc) = errcodepx
'                    arrcm(flc) = "EDC"
'                ElseIf Tagfnansi = True And Tagfpansi = False Then 's
'                    arrfnansi(flc) = errcodenx
'                    arrfpansi(flc) = ""
'                    arrcm(flc) = "ENC" '���ֵ�λ��
'                ElseIf Tagfnansi = False And Tagfpansi = True Then 's
'                    arrfnansi(flc) = ""
'                    arrfpansi(flc) = errcodepx
'                    arrcm(flc) = "EPC" '���ֵ�λ��
'                End If
'            End If
'        End If
'        Else
'        J = 1
'        If f = 1 Then GoTo 1010
'    End If
'
'ElseIf p = 0 Then                                                             '����������
'    If exet Like "EPUB" Or exet Like "PDF" Or exet Like "MOBI" Then
'        t = 2
'        If f = 1 Then GoTo 1010
'1012
'        DoEvents                                                    '���ڼ���hash���ٶȻ����,������Ҫ��doevents����������ִ�еļ�����״��
'        filemd5 = UCase(FileHashes(FilePath)) '��ģ�����hash��Ч�����,����ͻ��2G��С������
'    Else
'        J = 2
'    End If
'ElseIf p = -1 Then GoTo 100 '�޷���ȡ�ļ�·��
'End If
'
''--------------------------------------------------------------------------------------�ж��ļ�·���Ƿ���ڷ�ansi�����ַ�/�Ƿ����md5/����md5�ķ�ʽ
'
'If elow > 5 Then '���Ѿ�������Ŀ¼��ʱ��
'    If J > 0 Then
'1010
'    With ThisWorkbook.Sheets("���")
'        Set rngmd = .Range("c6:c" & .[c65536].End(xlUp).Row).Find(fl.Name) '����Ƿ�ͬ���ļ��Ѵ���
'    End With
'    If rngmd Is Nothing Then '����ļ����Ƿ���ͬ
'        If f = 1 And ckn = 0 And J = 0 Then     '�����ļ����µ�ʱ��������ʽ���ļ��Ƚ����ļ����ж�,���������ͬ���ļ��ͽ���md5����
'            ckn = 1
'            If t = 1 Then
'                GoTo 1011 '���»�ȡ�ļ�����ϸ��Ϣ
'            ElseIf t = 2 Then
'                GoTo 1012
'            End If
'        End If
'1003
'        If w > 0 Then
'            flr = Filter(arrbase, fl.Name) '�Ƚ����ļ����ж�(���ٲ���Ҫ��ѭ��)
'            If UBound(flr) >= 0 Then
'                For h = 1 To flc        '�Ƚ����ļ����ǰ�ıȽ�
'                    If fl.Name = arrbase(h) And fl.Size = arrsizeb(h) And fl.DateLastModified = arredit(h) Then GoTo 100
'                Next
'            End If
'        End If
'
'        w = w + 1
'        If J = 2 Then
'            flc = flc + 1
'            arrmd5(flc) = ""
'            arrcode(flc) = ""
'            arrcm(flc) = ""
'            arrfnansi(flc) = ""
'            arrfpansi(flc) = ""
'        ElseIf J = 1 Then
'            flc = flc + 1
'            arrmd5(flc) = ""
'            arrcode(flc) = "ERC"
'            If Tagfnansi = True And Tagfpansi = True Then 'd
'                arrfnansi(flc) = errcodenx
'                arrfpansi(flc) = errcodepx
'                arrcm(flc) = "EDC"
'            ElseIf Tagfnansi = True And Tagfpansi = False Then 's
'                arrfnansi(flc) = errcodenx
'                arrfpansi(flc) = ""
'                arrcm(flc) = "ENC" '���ֵ�λ��
'            ElseIf Tagfnansi = False And Tagfpansi = True Then 's
'                arrfnansi(flc) = ""
'                arrfpansi(flc) = errcodepx
'                arrcm(flc) = "EPC" '���ֵ�λ��
'            End If
'        End If
'    Else
'
'        If f = 1 Then GoTo 100 '�ļ��Ѵ���Ŀ¼(���µķ�ʽ,���ڼ������Ƚ�)
'
'        If rngmd.Offset(0, 4).Value <> fl.Size Then '�ļ���С
'            faddress = rngmd.address '���������ļ�����ͬ�ĵ�һ��λ��
'            Do
'                With ThisWorkbook.Sheets("���") '����ļ�����ͬ,�ļ���С��һ��,��ִ����һ��ͬ���ļ����Ĵ�С�Ƚ�
'                    Set rngmd = .Range("c6:c" & .[c65536].End(xlUp).Row).FindNext(rngmd) '�����һ��ͬ���ļ�
'                End With
'                If rngmd.Offset(0, 4) = fl.Size Then
'                    Exit Do
'                    GoTo 1006 '��һ���ļ��Ĵ�С������ļ�һ��
'                End If
'            Loop Until faddress = rngmd.address 'ѭ�������е�Ŀ¼,�ص���һ��λ��
'            w = w + 1
'            If J = 2 Then
'                flc = flc + 1
'                arrmd5(flc) = ""
'                arrcode(flc) = ""
'                arrcm(flc) = ""
'                arrfnansi(flc) = ""
'                arrfpansi(flc) = ""
'            ElseIf J = 1 Then
'                flc = flc + 1
'                arrmd5(flc) = ""
'                arrcode(flc) = "ERC"
'                If Tagfnansi = True And Tagfpansi = True Then 'd
'                    arrfnansi(flc) = errcodenx
'                    arrfpansi(flc) = errcodepx
'                    arrcm(flc) = "EDC"
'                ElseIf Tagfnansi = True And Tagfpansi = False Then 's
'                    arrfnansi(flc) = errcodenx
'                    arrfpansi(flc) = ""
'                    arrcm(flc) = "ENC" '���ֵ�λ��
'                ElseIf Tagfnansi = False And Tagfpansi = True Then 's
'                    arrfnansi(flc) = ""
'                    arrfpansi(flc) = errcodepx
'                    arrcm(flc) = "EPC" '���ֵ�λ��
'                End If
'            End If
'        Else
'1006
'            If rngmd.Offset(0, 5) <> fl.DateLastModified Then '�ļ��޸ĵ�ʱ��
'                w = w + 1
'                If J = 2 Then
'                    flc = flc + 1
'                    arrmd5(flc) = "DP"
'                    arrcode(flc) = ""
'                    arrcm(flc) = ""
'                    arrfnansi(flc) = ""
'                    arrfpansi(flc) = ""
'                ElseIf J = 1 Then
'                    flc = flc + 1
'                    arrmd5(flc) = "DP"
'                    arrcode(flc) = "ERC"
'                    If Tagfnansi = True And Tagfpansi = True Then 'd
'                        arrfnansi(flc) = errcodenx
'                        arrfpansi(flc) = errcodepx
'                        arrcm(flc) = "EDC"
'                    ElseIf Tagfnansi = True And Tagfpansi = False Then 's
'                        arrfnansi(flc) = errcodenx
'                        arrfpansi(flc) = ""
'                        arrcm(flc) = "ENC" '���ֵ�λ��
'                    ElseIf Tagfnansi = False And Tagfpansi = True Then 's
'                        arrfnansi(flc) = ""
'                        arrfpansi(flc) = errcodepx
'                        arrcm(flc) = "EPC" '���ֵ�λ��
'                    End If
'                End If
'            Else
'                ls = ls + 1
'                GoTo 100
'            End If
'        End If
'        Set rngmd = Nothing
'        End If
'    End If
''---------------------------------------------------------������md5�ķ�ʽ���ж��ļ��Ƿ��ص�
'    If t > 0 Then        '�����md5����бȽ�
'        With ThisWorkbook.Sheets("���")
'            Set rngmd = .Range("k6:k" & .[k65536].End(xlUp).Row).Find(filemd5) '����Ƿ��ļ��Ѵ���(����hash����Ψһ��,����Ҫ���ж��αȽ�)
'        End With
'        If rngmd Is Nothing Then '�ļ��������ص�/��·�������������ַ�
'1004
'            If n > 0 Then '��ֵ��ʱ����бȽ�
'                flr = Filter(arrmd5, filemd5) 'filter����,����ɸѡ����
'                If UBound(flr) >= 0 Then GoTo 1005 '�����ص����ļ�(û���ص��ļ�,ֵΪ-1)
'            End If
'            n = n + 1
'            If t = 1 Then
'                flc = flc + 1                             '��¼�ж����е�����
'                arrmd5(flc) = filemd5
'                arrcode(flc) = "ERC"
'                If Tagfnansi = True And Tagfpansi = True Then 'd
'                    arrfnansi(flc) = errcodenx
'                    arrfpansi(flc) = errcodepx
'                    arrcm(flc) = "EDC"
'                ElseIf Tagfnansi = True And Tagfpansi = False Then 's
'                    arrfnansi(flc) = errcodenx
'                    arrfpansi(flc) = ""
'                    arrcm(flc) = "ENC" '���ֵ�λ��
'                ElseIf Tagfnansi = False And Tagfpansi = True Then 's
'                    arrfnansi(flc) = ""
'                    arrfpansi(flc) = errcodepx
'                    arrcm(flc) = "EPC" '���ֵ�λ��
'                End If
'            ElseIf t = 2 Then
'                flc = flc + 1                             '��¼�ж����е�����
'                arrmd5(flc) = filemd5
'                arrcode(flc) = ""
'                arrcm(flc) = ""
'                arrfnansi(flc) = ""
'                arrfpansi(flc) = ""
'            End If
'            Set rngmd = Nothing
'        Else
'1005
'            If f = 1 Then GoTo 100 '�ļ��Ѵ���Ŀ¼
'            Set rngmd = Nothing
'            dl = dl + 1
'            If t = 1 Then
'                fso.DeleteFile (FilePath) '��kill����һ����ֱ��ɾ�����ļ�
'            Else
'                DeleteFiles (FilePath) '����ļ���ͬ��ɾ���ļ�,�Ƴ�������վ(��֧���쳣�ַ���ɾ��,kill����ͬ������һ��������)
'                GoTo 100 '�ļ���ͬ,ִ����һ���ļ�
'            End If
'            With ThisWorkbook.Sheets("Ŀ¼")
'                .Cells(a, b + c + 2) = fd.DateCreated 'ɾ���������ļ��е�ʱ�䷢���仯
'            End With
'        End If
'    End If
''---------------------------------------------------------------------------------------------------------------����md�����ļ��ж�
'Else         '�����δд�����ݵ�ʱ��
'    If J > 0 Then
'        GoTo 1003
'    ElseIf t > 0 Then
'        GoTo 1004
'    End If
'End If
'
''-----------------------------------------------------------------------------------------------------------------------------------------�ж��ļ��Ƿ������Ŀ¼/��ӵ��ļ����Ƿ�����ص����ļ�/ִ��ɾ����������
'
'With fl
'    arrfiles(flc) = .Path                  '����洢�ļ���·��
'    arrbase(flc) = .Name                     '������չ�����ļ���
'    arrextension(flc) = fso.GetExtensionName(.Path)                    '�ļ���չ��
'    arrparent(flc) = .ParentFolder '�ļ���һ��Ŀ¼
'    arrdate(flc) = .DateCreated '�ļ�����
'    arrsizeb(flc) = .Size         '�ļ���С
'    arredit(flc) = .DateLastModified '�ļ��޸�ʱ��
'    If .Size < 1048576 Then
'        arrsize(flc) = Format(.Size / 1024, "0.00") & "KB"    '�ļ��ֽڴ���1048576��ʾ"MB",������ʾ"KB"
'    Else
'        arrsize(flc) = Format(.Size / 1048576, "0.00") & "MB"
'    End If
'End With
''---------------------------------------------------------------------------------------------------------------------д������
'100
''ckn = 0 '��������
''t = 0
''J = 0
''End Function
