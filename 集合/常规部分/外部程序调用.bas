Attribute VB_Name = "�ⲿ�������"
Option Explicit
'�ⲿ����ĵ�����Ҫ����һϵ�еĸ�������
'�ⲿ�����·��-�漰���������ַ�/�ո����
'�ⲿ����ִ�н���ķ���
'ִ�н�����صĴ����Ƿ�ͬ��
'ִ��Ŀ���·������
'ִ��ʧ�ܵĴ���
'��εȴ�ִ�н���ķ���, wsh.run֧��ͬ��, ��Ҫ����ִ�е�ʱ�����������,���Դ���shell����һ��ʹ��
'------------------------------��Ҫ�漰�ⲿ����:CMD, Powershell, 7Zip, BandZip
'---------Environ("comspec"), cmd��·��
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Function Webbrowser(ByVal url As String, Optional ByVal cmCode As Byte) '���������
    Dim exepath As String
    
    With ThisWorkbook.Sheets("temp")
        exepath = .Range("ab10").Value
        If Len(exepath) = 0 Or fso.fileexists(exepath) = False Then '���������������
            If Len(.Range("ab6").Value) > 0 Then
                If cmCode = 2 Then
                    Turlx = url
                    UserForm15.Show
                    Exit Function
                End If
                If cmCode = 1 Then UserForm3.Label57.Caption = "��վ�㲻֧��IE�����": Exit Function
                exepath = Environ("SYSTEMDRIVE") & "\Program Files\Internet Explorer\iexplore.exe" & Chr(32) '���û�����������������.Ĭ�ϵ���IE
            Else
                UserForm3.Label57.Caption = "��δ���������"
                Exit Function
            End If
        Else
            exepath = exepath & Chr(32)
        End If
    End With
    exepath = """" & exepath & """"
    Shell exepath & url, vbNormalFocus
End Function

Function CheckFileX(ByVal folderp As String) As Boolean '����ļ������Ƿ����Ŀ���ļ� 'https://www.cnblogs.com/zhaoqingqing/p/4620402.html
    Dim strx As String, strx2 As String, strx1 As String, strx3 As String, strx4 As String
    Dim c As Byte, i As Byte, xi As Variant, k As Byte

    CheckFileX = True
    strx4 = ThisWorkbook.Path & "\checklist.txt"
    If fso.fileexists(strx4) = True Then fso.DeleteFile (strx4): Sleep 25 '�ȴ��ļ�ɾ��
    xi = Split(folderp, "\")
    c = UBound(xi)
    If fso.GetDriveName(folderp) <> Environ("SYSTEMDRIVE") Then k = 1 '�ж��ļ�����Դ�Ƿ���ϵͳ��
    strx = Left(folderp, 2) '��ȡ�̷�
    i = InStrRev(folderp, "\")
    strx1 = Left(folderp, i - 1)
    strx2 = "cd " & strx1
    strx3 = xi(c)
    If c = 1 Then
        Shell ("cmd /c " & strx & " && for /r " & strx3 & " %i in (*.epub,*.pdf,*.mobi,*.xl*,*.doc*,*.txt,*.pp*,*.ac*) do @echo %i >" & strx4), vbHide
    ElseIf k = 0 And c > 1 Then
        Shell ("cmd /c " & strx2 & " && for /r " & strx3 & " %i in (*.epub,*.pdf,*.mobi,*.xl*,*.doc*,*.txt,*.pp*,*.ac*) do @echo %i >" & strx4), vbHide
    ElseIf k = 1 And c > 1 Then
        Shell ("cmd /c " & strx & " && strx2 " & " && for /r " & strx3 & " %i in (*.epub,*.pdf,*.mobi,*.xl*,*.doc*,*.txt,*.pp*,*.ac*) do @echo %i >" & strx4), vbHide
    End If
    Sleep 200 '�ȴ��ļ�����
    If fso.fileexists(strx4) = False Then CheckFileX = False 'û��Ŀ���ļ�
End Function

Function ZipCompress(ByVal xpath As String, Optional ByVal cmbagname As String, Optional ByVal cmCode As Byte, Optional ByVal passwordx As String) '����7zip��ѹ���ļ� '·��'ѹ��������'ѹ��������'ѹ�����õ�����
    Dim exepath As String, ax As Variant, i As Byte
    Dim strx As String, strx1 As String, strx2 As String, strx3 As String, strx4 As String  '�����ⲿ����ʱ,������cmd����Powershell,������Ҫע��
    
'    exepath = "C:\Program Files\7-Zip\7z.exe" '�����װ������λ��
    exepath = ThisWorkbook.Sheets("temp").Cells(32, "ab").Value
    If fso.fileexists(exepath) = False Then MsgBox "��δ���ý�ѹ���", vbOKOnly, "Tips": Exit Sub
    exepath = exepath & Chr(32)
    exepath = """" & exepath & """"
    i = InStrRev(xpath, "\") - 1
    If Len(cmbagname) = 0 Then 'ѹ����������
        If InStr(xpath, ".") > 0 Then '��ʾ�����ļ�
            ax = Split(xpath, "\")
            strx1 = ax(UBound(ax))
            strx1 = Split(strx1, ".")(0)
            cmbagname = strx1
        Else  '�����ļ���
            i = i + 1
            cmbagname = Right(xpath, Len(xpath) - i) '��ȡ�ļ�����
            i = i - 1
            xpath = xpath & "\"
        End If
    End If
    
    strx2 = Left(xpath, i) '���ѹ�������ļ��ͽ����ļ����ڵ��ļ�����, ���ѹ�������ļ��оͷ�����һ���ļ���
    cmbagname = cmbagname & ".7z" 'ʹ��7z��Ϊѹ�����ĸ�ʽ
    cmbagname = """" & cmbagname & """"
    xpath = """" & xpath & """"
     '������ʲô������ⲿ����,����Ҫע�������������еĸ�����, ��ո�,������,���ɼ��ַ��ȵ���Щ��������
    Select Case cmCode 'ѹ�� 'ѹ����ɾ�� 'ѹ������� 'ѹ������ܲ�ɾ��
        Case 0
        strx3 = exepath & " a " & cmbagname & " " & xpath
        Case 1
        strx3 = exepath & " a " & "-sdel " & cmbagname & " " & xpath 'a��ʾadd����ļ���ѹ����
        Case 2
        strx3 = exepath & " a " & "-p" & passwordx & " " & cmbagname & " " & xpath
        Case 3
        If Len(passwordx) = 0 Then passwordx = "password" '���ѡ��3,����û����������,Ĭ�Ͻ���������Ϊ"password"
        strx3 = exepath & " a " & "-p" & passwordx & " " & "-sdel " & cmbagname & " " & xpath
        Case Else
        Exit Function
    End Select
    '��ѹ�ļ���ʱ����õķ���Ҳ�����Ƶ�(���������ļ��Ľ�ѹ), ֻ������Ϊ" e "
    '������չһ��,�������д���ƽ�ѹ������С����,��Ȼ���볤�Ⱥ�����ʹ�õ��ַ�Ҫ������,�����ƽ������Ѷ���������ĸ��Ӷȳɼ�������(3λ��Ϊ��)10*10*10=1000 * 46.656/36*36*36=46656 *5.1=/62*62*62=238 328
    strx4 = "cd " & strx2 '��������Ŀ¼
                         'cmdִ�ж������ʱ,���� "&&" ��Ϊ���ӷ� 'ע�������ֱ�ӵ���7zip��������
    Shell ("cmd /c " & strx4 & "&&" & strx3), vbHide '����7zip����ѹ���ļ�,a��ʾadd,����ļ���ѹ����, -sdel,ѹ��֮��ɾ����Դ�ļ� ,-p��ʾ�������� '���������Ϣ�ɲ鿴7z��װĿ¼�µ�chm�ļ�
    '����ϲ�����Խ���صĹ��ܺϲ�����ģ�鷽�����,��ѹ��,sha1�����
End Function

Function FileisOpen(ByVal FilePath As String) As Boolean  '����powershell�ű����ж��ļ��Ƿ��ڴ򿪵�״̬,֧�ַ�ansi�ַ�·���ļ�'����powershellִ��ps�ű�������'��Ҫ����Ȩ��'�Թ���Ա���ִ��powershell ����: set-executionpolicy remotesigned �ر�����: Set-ExecutionPolicy Restricted
    Dim strOutput As String, Jsfilepath As String
    Dim WshShell As Object, WshShellExec As Object
    '-------------------------------------------------------------��Ҫע��txt�ļ��ǲ��������ڴ򿪵�״̬,���Կ������ɿ���txt�ĵ�
    Jsfilepath = ThisWorkbook.Sheets("temp").Range("ab7").Value
    If Len(Jsfilepath) = 0 Or PSexist = False Then Exit Function
    strCommand = "Powershell.exe -ExecutionPolicy ByPass " & Jsfilepath & Chr(32) & FilePath  'chr(32) ��ʾ�ո�� 'ע��strcommand�����е�����ո��
    Set WshShell = CreateObject("WScript.Shell")
    Set WshShellExec = WshShell.Exec(strCommand)
    strOutput = WshShellExec.StdOut.ReadAll     '����ִ�еĽ��,����ļ����ڴ򿪵�״̬���з���ֵ,������ǿ�ֵ
    If Len(strOutput) > 0 Then
        FileisOpen = True
    Else
        FileisOpen = False
    End If
    Set WshShell = Nothing
    Set WshShellExec = Nothing
End Function

Function HashPowershell(filepaths As String) As String 'ʹ��powershell������md5,ע�ⲻͬ��ʽ���ڷ��ŵ�����ʹ�� 'https://docs.microsoft.com/en-us/windows-server/administration/windows-commands/certutil
    Dim wsh As Object
    Dim wExec As Object
    Dim exepath As String
    Dim filepathx As String
    Dim Result As String, i As Byte, strx As String
    
    Set wsh = CreateObject("WScript.Shell")
     '���ﲻͬ��ִ�д�����Ҫע������ļ�·������ĳЩ������ַ���ɴ����޷���ִ�л���ִ�г��������,�����ļ�·������'�ո�()"�ȷǳ��и��ŵ��ַ�
    If InStr(filepaths, Chr(39)) > 0 Then
        filepathx = """" & filepaths & """"
        Set wExec = wsh.Exec("powershell certutil -hashfile " & """" & """" & filepathx & """" & """" & " md5")
    Else
        filepathx = "'" & filepaths & " '"
        Set wExec = wsh.Exec("powershell certutil -hashfile " & filepathx & " MD5")
    End If
    'Set wExec = wsh.Exec("powershell Get-FileHash " & filepathx & " -Algorithm MD5| Format-List")
    Result = wExec.StdOut.ReadAll
    If Len(Result) = 0 Then 'û�л�ȡ��hash
        HashPowershell = "UC"
        Exit Function
    End If
    i = UBound(Split(Result, ":")) - 1
    strx = Replace(Replace(Split(Result, ":")(i), Chr(10), ""), " ", "")
    HashPowershell = Left$(strx, Len(strx) - 8) '1��ʽ��ȡֵ�ķ�ʽ
    'Hashpowershell = UCase(Trim(Replace(Replace(Split(Split(Result, "Hash")(1), "Path")(0), ":", ""), Chr(10), ""))) 'chr10��ʾ���з�
    Set wExec = Nothing
    Set wsh = Nothing
End Function

Function ZipHash(ByVal FilePath As String) As String 'ʹ��7zip������sha1 or sha256/sha1�� ��֧��md5-���� 'https://www.7-zip.org/
    Dim wsh As Object
    Dim wExec As Object
    Dim exepath As String, filepathx As String
    Dim filemdx As String, Result As String '�������ص����ݿ��Բ���zip��װĿ¼�µ�chm�ĵ�
    
'    exepath = "C:\Program Files\7-Zip\7z.exe "
    exepath = ThisWorkbook.Sheets("temp").Cells(32, "ab").Value
    If fso.fileexists(exepath) = False Then MsgBox "��δ���ý�ѹ���", vbOKOnly, "Tips": Exit Sub
    exepath = exepath & Chr(32)
    If fso.fileexists(Trim(exepath)) = False Or fso.fileexists(Trim(FilePath)) = False Then Exit Function
    filepathx = """" & FilePath & """"                          'ע������ķ��ŵ�ʹ��,��α�ʾһ������
    Set wsh = CreateObject("WScript.Shell")
    Set wExec = wsh.Exec(exepath & "h -scrcsha256 " & filepathx) '�˷�����ȱ�������޷�����ִ�д���
    Result = wExec.StdOut.ReadAll
    If Len(Result) = 0 Then
    ZipHash = "UC"
    Set wExec = Nothing
    Set wsh = Nothing
    Exit Function
    End If
    filemdx = Trim(Split(Result, "SHA256 for data:")(UBound(Split(Result, "SHA256 for data:"))))
    ZipHash = Replace(Left$(filemdx, Len(filemdx) - 18), Chr(10), "") 'chr(10)��ʾascii��Ļ��з�
    Set wExec = Nothing
    Set wsh = Nothing
End Function

Function PowerSHForceW(ByVal FilePath As String, ByVal filez As Long) 'ͨ��Powershellǿ�����ܱ������ļ�д������
    Dim ws As Object
    Dim strx As String, commandline As String, strx1 As String
    'Powershell����д�����ݵķ���������out-file ����system.IO�ķ�ʽ���޷�ֱ�����ܱ������ļ�д������
    filez = filez \ 2
    If filez > 104857600 Then filez = 104857600 '����100M
    PowerSHCreate filez
    strx = ThisWorkbook.Path & "\temp.txt"
    strx = "'" & strx & "'"
    If InStr(FilePath, Chr(39)) = 0 Then
        FilePath = "'" & FilePath & "'"
        strx1 = FilePath
    Else
        FilePath = """" & FilePath & """"
        strx1 = """" & """" & FilePath & """" & """"
    End If
    Set ws = CreateObject("wscript.shell")
    'https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/get-content?view=powershell-6
    'https://docs.microsoft.com/en-us/powershell/module/Microsoft.PowerShell.Utility/Out-File?view=powershell-6
    commandline = "powershell $bulk=" & strx & ";$file=" & strx1 & ";(Get-Content $bulk) | Set-Content $file -Encoding Unknown -Force" & ";Remove-Item $bulk"
    '�����ɵ����������ļ�,100M���ļ�������������(����ǵ�����ps1�ű����޴�����)
    ws.Run (commandline), 0, True
    Set ws = Nothing
    Sleep 50
End Function

Function PowerSHCreate(ByVal filez As Long) '����powershell��������һ�������С���ļ�
    Dim TempFile As String, commandline As String
    Dim ws As Object
    'https://docs.microsoft.com/en-us/dotnet/api/system.io.streamwriter?view=netframework-4.8
    TempFile = ThisWorkbook.Path & "\temp.txt"
    TempFile = "'" & TempFile & "'"
    commandline = "$tempFile=" & TempFile & _
    ";$fs=New-Object System.IO.FileStream($tempFile,[System.IO.FileMode]::OpenOrCreate)" & _
    ";$fs.Seek(" & filez & " ,[System.IO.SeekOrigin]::Begin)" & _
    ";$fs.WriteByte(1)" & _
    ";$fs.Close()"
    Set ws = CreateObject("wscript.shell")
    ws.Run ("powershell " & commandline), 0, True
    Set ws = Nothing
    Sleep 50
End Function

Sub CmdOpenFile(ByVal FilePath As String) 'ͨ��cmd��Ӵ��ļ�
    '--------------------------------��Ҫ��Ĭ�ϴ��ļ��Ĺ����������
    FilePath = """" & FilePath & """"
    Shell ("cmd /c" & FilePath), vbHide
End Sub

Function PowerSHOpen(ByVal FilePath As String) '����powershell ���ļ�
    Dim ws As Object
    Dim commandline As String
    
    If InStr(FilePath, Chr(39)) > 0 Then
        FilePath = """" & FilePath & """"
        commandline = "powershell " & "invoke-item " & """" & """" & FilePath & """" & """"
    Else
        FilePath = "'" & FilePath & "'"
        commandline = "powershell " & "invoke-item " & FilePath
    End If
    Set ws = CreateObject("wscript.shell")
    ws.Run (commandline), 0
    Set ws = Nothing
End Function

Sub ZipExtract(ByVal FilePath As String, ByVal Folderpath As String) '����7zip�����֧����cmd commandline����������,���ò�����bandzip
    Dim i As Byte
    Dim strx As String, strx1 As String, exe As String
    Dim wsh As Object
    '��7zip���,ò���������г�����΢�Ĳ���,�����������˳����,��������Ĳ�����ȫ����ͬ��
    Set wsh = CreateObject("WScript.Shell")
'    exe = "C:\Program Files\Bandizip\bc.exe "
    exe = ThisWorkbook.Sheets("temp").Cells(52, "ab").Value
    If fso.fileexists(exe) = False Then MsgBox "��δ���ý�ѹ���", vbOKOnly, "Tips": Exit Sub
    exe = exe & Chr(32)
    strx = """" & FilePath & """"
    strx1 = """" & Folderpath & """"
    exe = """" & exe & """"
    wsh.Run (exe & " e " & "-aoa " & strx & " -o " & strx1), vbNormalFocus '֧��������������������
    Set wsh = Nothing
End Sub

Function TerminateEXE(ByVal exename As String, Optional ByVal cmCode As Byte = 0) As Byte '��ֹ�ض�����
    Dim obj As Object, targetexe As Object, targetexex As Object
                                                            'https://docs.microsoft.com/en-us/windows/win32/wmisdk/wmi-tasks--processes
    On Error GoTo 100
    TerminateEXE = 0
    Set obj = GetObject("winmgmts:\\.\root\cimv2")          'https://docs.microsoft.com/en-us/windows/win32/wmisdk/connecting-to-wmi-with-vbscript
    If obj Is Nothing Then MsgBox "�޷���������": Exit Function
    Set targetexe = obj.ExecQuery("select * from win32_process where name=" & Chr(39) & exename & Chr(39)) '("select * from win32_process where name='iexplore.exe'")' chr(39)='
    If targetexe Is Nothing Then Exit Function
    For Each targetexex In targetexe
    If cmCode = 1 Then
        targetexex.Terminate
    Else
        TerminateEXE = 1: Exit For
    End If
    Next
    Set obj = Nothing
    Set targetexe = Nothing
    Exit Function
100
    Set obj = Nothing
    Set targetexe = Nothing
    Err.Clear
End Function

Sub CopyFileClipboard(ByVal FilePath As String) 'ͨ��powershell�����ļ������а�
    Dim commandline As String, strx As String
    
    FilePath = """" & FilePath & """"   '���ļ�:'filepath = "'C:\text.txt','D:\text.txt','D:\text.jpg'"
    strx = """" & """" & FilePath & """" & """"
    commandline = "powershell $filelist =" & strx & vbCrLf & _
    "$col = New-Object Collections.Specialized.StringCollection " & vbCrLf & _
    "foreach($file in $filelist){$col.add($file)}" & vbCrLf & _
    "Add-Type -AssemblyName System.Windows.Forms" & vbCrLf & _
    "[Windows.Forms.Clipboard]::setfiledroplist($col)"
    Shell commandline, vbHide
End Sub

Sub cmdFindText(ByVal FilePath As String, ByVal Keyword As String) '��cmd����txt�ı��ڴ��йؼ��ʵ���
    Keyword = """" & Keyword & """"
    FilePath = """" & FilePath & """"
    Shell "cmd /c" & "findstr /c:" & strx & " " & FilePath & " | clip", vbHide '���shell����һ��ʹ��
    '---------------------------���ҵ���������������а�
End Sub
