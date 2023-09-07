Attribute VB_Name = "����"
Option Explicit
Private Declare Function _
  InternetGetConnectedState _
  Lib "wininet.dll" (ByRef lpdwFlags As Long, _
  ByVal dwReserved As Long) As Long             'https://www.cnblogs.com/fuchongjundream/p/3853716.html
Private Const _
  INTERNET_CONNECTION_MODEM_BUSY As Long = &H8   'https://docs.microsoft.com/en-us/windows/win32/wininet/wininet-functions
Private Const _
  INTERNET_RAS_INSTALLED As Long = &H10
Private Const _
  INTERNET_CONNECTION_OFFLINE As Long = &H20
Private Const _
  INTERNET_CONNECTION_CONFIGURED As Long = &H40
  '------------------------------------------------������ģ��
Private Const SAFT48kHz16BitStereo = 39
Private Const SSFMCreateForWrite As Byte = 3 ' Creates file even if file exists and so destroys or overwrites the existing file
Private Const SSFMOpenForRead As Byte = 0
Private Const SSFMOpenReadWrite As Byte = 1
Private Const SSFMCreate As Byte = 2
'Enum SpeechStreamFileMode
'    SSFMOpenForRead = 0
'    SSFMOpenReadWrite = 1
'    SSFMCreate = 2
'    SSFMCreateForWrite = 3
'End Enum
'-----------------------------https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms720892(v=vs.85)?redirectedfrom=MSDN
'Enum SpeechVoiceSpeakFlags
'    'SpVoice Flags
'    SVSFDefault = 0
'    SVSFlagsAsync = 1
'    SVSFPurgeBeforeSpeak = 2
'    SVSFIsFilename = 4
'    SVSFIsXML = 8
'    SVSFIsNotXML = 16
'    SVSFPersistXML = 32
'
'    'Normalizer Flags
'    SVSFNLPSpeakPunc = 64
'
'    'TTS Format
'    SVSFParseSapi =
'    SVSFParseSsml =
'    SVSFParseAutoDetect =
'
'    'Masks
'    SVSFNLPMask = 64
'    SVSFParseMask =
'    SVSFVoiceMask = 127
'    SVSFUnusedFlags = -128
'End Enum
'Enum SpeechStreamSeekPositionType
'    SSSPTRelativeToStart = 0
'    SSSPTRelativeToCurrentPosition = 1
'    SSSPTRelativeToEnd = 2
'End Enum
'-----------------------------------------��������
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'ʱ�� -����ѵ��
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Public SpEnd As Boolean
Public SpStart As Boolean 'sapi���źͲ��Ž�������

Private Declare Function CoCreateGuid Lib "ole32" (id As Any) As Long

Private Function CreateGUID() As String '����GUID
    Dim idBytes(0 To 15) As Byte
    Dim Cnt As Long, GUID As String
    '----------------------------https://www.cnblogs.com/snandy/p/3261754.html
    If CoCreateGuid(idBytes(0)) = 0 Then
        For Cnt = 0 To 15
            CreateGUID = CreateGUID + IIf(idBytes(Cnt) < 16, "0", "") + Hex$(idBytes(Cnt))
        Next Cnt
        CreateGUID = Left$(CreateGUID, 8) + "-" + Mid$(CreateGUID, 9, 4) + "-" + Mid$(CreateGUID, 13, 4) + "-" + Mid$(CreateGUID, 17, 4) + "-" + Right$(CreateGUID, 12)
    End If
End Function

Function TextToVoice(ByVal Oputfile As String, ByVal cmCode As Byte, Optional ByVal Inputfile As String, Optional ByVal strText As String) As Boolean '���ı�ת��Ϊ����
    Dim oFileStream As Object, oVoice As Object, oFileOpen As Object
    
    'https://docs.microsoft.com/zh-cn/previous-versions/windows/desktop/ms722561(v=vs.85)
    'https://docs.microsoft.com/zh-cn/previous-versions/windows/desktop/ms723602(v=vs.85)
    'https://www.cnblogs.com/sntetwt/p/3533632.html
    'EN �汾��windows û��CN�����԰�,��Ҫ���ⰲװ
    'windows7���ı���EN����Ч������(Ҳ�������������԰�?���Ը���?)
    '����������Կ��Կ��ٽ��ı�ת��Ϊ�����ļ�, �絥�ʵĴʿ�
    TextToVoice = True
    Set oFileStream = CreateObject("SAPI.SpFileStream")
    oFileStream.Format.type = SAFT48kHz16BitStereo  '�������Ƶ�ļ�
    '------------------------------------------------��������wav��ʽ���ļ�
    If fso.fileexists(Oputfile) = True Then fso.DeleteFile Oputfile
    oFileStream.Open Oputfile, SSFMCreateForWrite 'C:\Users\***\Downloads\A111\Sample.wav"
    Set oVoice = CreateObject("SAPI.SpVoice")
    Set oVoice.AudioOutputStream = oFileStream
    If cmCode = 1 Then
        If fso.fileexists(Inputfile) = False Then TextToVoice = False: Set oVoice = Nothing: Exit Function
        Set oFileOpen = CreateObject("SAPI.SpFileStream") '�����txt�ļ�
        oFileOpen.Open Inputfile, SSFMOpenForRead, False ''C:\Users\***\Downloads\A111\Sample.txt" û�в��Դ��ļ���Ч��, ����޷�ֱ�Ӷ�ȡ,�����Ƚ��ı���������ȡ����,�ٽ�����תΪ����
        oVoice.SpeakStream oFileOpen 'ע�����ﲻ��.speak, speak ������string���͵����� ,�� speak "hello, world"
        oFileOpen.Close
        Set oFileOpen = Nothing
    Else
        oVoice.Speak strText
    End If
    oFileStream.Close
    Set oFileStream = Nothing
    Set oVoice = Nothing
End Function

Function IsNetConnectOnline() As Boolean '���õ����������ӷ���
    IsNetConnectOnline = InternetGetConnectedState(0&, 0&)
End Function

Function GetOsVersion() As String '��ȡϵͳ�İ汾��Ϣ
    Dim objWMIService As Object, colItems As Object, objItem As Variant, WinOSversion As String
    
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
    For Each objItem In colItems
        WinOSversion = objItem.Version
    Next
    Select Case Left$(WinOSversion, 3)
        Case "5.2": WinOSversion = "Windows Server 2003"
        Case "5.0": WinOSversion = "Windows 2000"
        Case "5.1": WinOSversion = "Windows XP"
        Case "6.0": WinOSversion = "windows vista"
        Case "6.1": WinOSversion = "Win7"
        Case "6.2": WinOSversion = "Win8"
        Case "6.3": WinOSversion = "Win8.1"
        Case "10.": WinOSversion = "Win10"
        Case Else: WinOSversion = "Unknow"
    End Select
     GetOsVersion = WinOSversion
     Set objWMIService = Nothing
     Set colItems = Nothing
End Function

Function WmiCheckFileOpen(ByVal FilePath As String, Optional ByVal exename As String) As Boolean '�ж��ļ��Ƿ��ڴ򿪵�״̬
    Dim strComputer As String, commandline As String
    Dim objWMIService As Object, colItems As Object, objItem As Object
    '------------------------------------------------------------------���ַ����ĺô����Ǳ�����open�����޷��������뱣�����ļ�, ����Ҳ���Խ��txt�ļ�������
    '------------------------------------------------------------------�޷�����excel�������ļ�
    strComputer = "."
    WmiCheckFileOpen = False
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
    If objWMIService Is Nothing Then MsgBox "�޷���������": Exit Function
    If Len(exename) = 0 Then
        commandline = "select * from win32_process"   '���ݲ�ͬ���������ɸѡ
    Else
        commandline = "select * from win32_process where name=" & Chr(39) & exename & Chr(39)
    End If
    Set colItems = objWMIService.ExecQuery(commandline)
    For Each objItem In colItems
        If InStr(objItem.commandline, FilePath) > 0 Then WmiCheckFileOpen = True: Exit For
    Next
    Set objWMIService = Nothing
    Set colItems = Nothing
End Function

Sub TerminateEXEs(ByVal exename As String) '��ֹ�ض�����
    Dim obj As Object, targetexe As Object, targetexex As Object
                                                            'https://docs.microsoft.com/en-us/windows/win32/wmisdk/wmi-tasks--processes
    On Error GoTo 100
    Set obj = GetObject("winmgmts:\\.\root\cimv2")          'https://docs.microsoft.com/en-us/windows/win32/wmisdk/connecting-to-wmi-with-vbscript
    If obj Is Nothing Then MsgBox "�޷���������": Exit Sub
    Set targetexe = obj.ExecQuery("select * from win32_process where name=" & Chr(39) & exename & Chr(39)) '("select * from win32_process where name='iexplore.exe'")' chr(39)='
    If targetexe Is Nothing Then Exit Sub
    For Each targetexex In targetexe
        targetexex.Terminate
        Exit For
    Next
    Set obj = Nothing
    Set targetexe = Nothing
    Exit Sub
100
    Set obj = Nothing
    Set targetexe = Nothing
    Err.Clear
    '������ֱ����ֹ����,���Բ鿴�ļ�ռ�õĳ�����,ֱ����ֹ������ܵ������ݶ�ʧ������
End Sub

Function IsPing(strMachines As String) As Boolean '����Ӧ���豸�Ƿ������ӵ�״̬
    Dim aMachines() As String
    Dim machine As Variant
    Dim objPing As Object
    Dim objStatus As Object
    
    aMachines = Split(strMachines, ";")
    For Each machine In aMachines
        Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_PingStatus where address = '" & machine & "'")
        If objPing Is Nothing Then IsPing = False: Exit Function '���û�ɹ���������
        For Each objStatus In objPing
            If IsNull(objStatus.StatusCode) Or objStatus.StatusCode <> 0 Then
                IsPing = False
            Else
                IsPing = True
            End If
        Next
    Next
    Set objPing = Nothing
End Function

Function GetPYA(ByVal textchar As String) As String
    Dim i%, pyArr As Variant, str$, ch$, Py As String, k As Integer
    '-------------------------------------------------------------------https://docs.microsoft.com/zh-cn/office/vba/api/excel.worksheetfunction.lookup
    pyArr = [{"߹","A";"��","B";"�c","C";"��","D";"��","E";"��","F";"�","G";"��","H";"آ","J";"��","K";"��","L";"��","M";"�p","N";"��","O";"��","P";"��","Q";"��","R";"��","S";"��","T";"��","W";"Ϧ","X";"Ѿ","Y";"��","Z"}]
    str = Replace(Replace(textchar, " ", ""), "��", "") '�滻���ո�
    k = Len(str)
    For i = 1 To k
        ch = Mid(str, i, 1)
        If ch Like "[һ-��]" Then   '����Ǻ��֣�����ת��
            GetPYA = GetPYA & WorksheetFunction.Lookup(Mid(str, i, 1), pyArr)
        Else
            GetPYA = GetPYA & UCase(ch)     '������Ǻ��֣�ֱ�����
        End If
    Next
End Function

Function GetPY(str As String) As String '��ȡ����ƴ��������ĸ(������ܿ�����������ʱ,���ƴ������ĸ��ģ������)
    Dim i As Integer
    
    For i = 0 To Len(str) - 1
        GetPY = GetPY & _
        IIf(IsChinese(asc(Mid(str, i + 1, 1))), _
        GetPYChar(Mid(str, i + 1, 1)), "")
    Next
    GetPY = LCase(GetPY)
End Function

Private Function IsChinese(ByVal AscVal As Integer) As Boolean ''�ж�ĳ��ASC���Ƿ�ָ��һ������
    IsChinese = IIf(Len(Hex(AscVal)) > 2, True, False)
End Function

Private Function GetPYChar(Char As String) As String ''��ȡ��Ӧ��������ĸ(���ֲ�����©)
    Dim lChar As Long
    lChar = 65536 + asc(Char)
    If (lChar >= 45217 And lChar <= 45252) Then
        GetPYChar = "A"
    ElseIf (lChar >= 45253 And lChar <= 45760) Then
        GetPYChar = "B"
    ElseIf (lChar >= 47761 And lChar <= 46317) Then
        GetPYChar = "C"
    ElseIf (lChar >= 46318 And lChar <= 46825) Then
        GetPYChar = "D"
    ElseIf (lChar >= 46826 And lChar <= 47009) Then
        GetPYChar = "E"
    ElseIf (lChar >= 47010 And lChar <= 47296) Then
        GetPYChar = "F"
    ElseIf (lChar >= 47297 And lChar <= 47613) Then
        GetPYChar = "G"
    ElseIf (lChar >= 47614 And lChar <= 48118) Then
        GetPYChar = "H"
    ElseIf (lChar >= 48119 And lChar <= 49061) Then
        GetPYChar = "J"
    ElseIf (lChar >= 49062 And lChar <= 49323) Then
        GetPYChar = "K"
    ElseIf (lChar >= 49324 And lChar <= 49895) Then
        GetPYChar = "L"
    ElseIf (lChar >= 49896 And lChar <= 50370) Then
        GetPYChar = "M"
    ElseIf (lChar >= 50371 And lChar <= 50613) Then
        GetPYChar = "N"
    ElseIf (lChar >= 50614 And lChar <= 50621) Then
        GetPYChar = "O"
    ElseIf (lChar >= 50622 And lChar <= 50905) Then
        GetPYChar = "P"
    ElseIf (lChar >= 50906 And lChar <= 51386) Then
        GetPYChar = "Q"
    ElseIf (lChar >= 51387 And lChar <= 51445) Then
        GetPYChar = "R"
    ElseIf (lChar >= 51446 And lChar <= 52217) Then
        GetPYChar = "S"
    ElseIf (lChar >= 52218 And lChar <= 52697) Then
        GetPYChar = "T"
    ElseIf (lChar >= 52698 And lChar <= 52979) Then
        GetPYChar = "W"
    ElseIf (lChar >= 52980 And lChar <= 53640) Then
        GetPYChar = "X"
    ElseIf (lChar >= 53689 And lChar <= 54480) Then
        GetPYChar = "Y"
    ElseIf (lChar >= 54481 And lChar <= 52289) Then
        GetPYChar = "Z"
    End If
End Function

'--------------------------------�����ı����ݵ�ʾ��,���ڴ���hosts�ļ������ݸ���
Sub DataClean() '��ϴ����
    Dim strx As String, arr() As String, arrTemp() As Variant
    Dim m As Long, n As Long, i As Long, k As Byte
    
    With ThisWorkbook.Sheets(1)
        m = .[a65536].End(xlUp).Row
        If m = 1 Then
        ReDim arrTemp(1, 1)
        arrTemp(1, 1) = .Cells(1, 1).Value
        Else
        arrTemp = .Range("a1:a" & m).Value
        End If
        n = 1
        ReDim arr(1 To m)
        For i = 1 To m
            If Len(arrTemp(i, 1)) = 0 Then GoTo 100
            If InStr(arrTemp(i, 1), ".") = 0 Then GoTo 100
                strx = Trim(arrTemp(i, 1))
                k = InStr(strx, Chr(35))
                If k > 1 Then
                    arr(n) = Trim(Split(Split(strx, Chr(35))(0), Chr(32))(1))
                    n = n + 1
                ElseIf k = 0 Then
                    arr(n) = Trim(Split(strx, Chr(32))(1))
                    n = n + 1
                End If
100
        Next
        .Cells(1, 2).Resize(n, 1) = Application.Transpose(arr)
    End With
End Sub
'��ȡ�ı�,���ı������ݶ�ȡ��Excel��, ����Excel��ȥ��,�ٽ����������ݺϲ�,�ٽ���������д��hosts�ļ�
Sub WriteHosts()            '��ȡ/д��hosts������
    Dim arr As Variant, arrTemp() As Variant, arrtempx() As String
    Dim i As Long, m As Long, k As Long
    Dim strx1 As String
    Dim fl As File
    Dim flop As Object, t As Single
    
    t = Timer
    With ThisWorkbook.Sheets(1)
        m = .[a65536].End(xlUp).Row
        If m = 0 Then Exit Sub
        strx1 = "C:\Users\***\Desktop\hosts"
        Open strx1 For Input As #1          '���ı�������ȫ����ȡ����,���зֿ�
        arr = Split(StrConv(InputB(LOF(1), 1), vbUnicode), vbNewLine)
        Close #1
        i = UBound(arr)
        ReDim arrtempx(i)
        For k = 0 To i
            arrtempx(k) = Trim(Split(arr(k), Chr(32))(1))
        Next
        .Cells(m + 1, 1).Resize(i + 1, 1) = Application.Transpose(arrtempx)
        m = .[a65536].End(xlUp).Row
        .Range("a267:a" & m).RemoveDuplicates Columns:=1, Header:=xlNo 'ֱ������Excel�Դ���ȥ��
        m = .[a65536].End(xlUp).Row
        arrTemp = .Range("a1:a" & m).Value
        Set fl = fso.GetFile(strx1)
        Set flop = fl.OpenAsTextStream(ForWriting, TristateUseDefault)
        
        With flop
        For i = 1 To m
            .Write "0.0.0.0 " & arrTemp(i, 1) & Chr(13)
        Next
            .Close
        End With
    End With
    Set fl = Nothing
    Set flop = Nothing
    Debug.Print Format(Timer - t, "0.0000")
End Sub
'-----------------------------------------------hosts����

Function CheckPort(ByVal Server As String, ByVal Port As Long) As Boolean '���Զ�˷������Ķ˿��Ƿ��ڿ��õ�״̬
    Dim SockObject As Object, i As Byte
    '��Ҫ����MSWINSCK.OCX,system32�µ�����ؼ�(x86),win7֧��
    'https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa733709(v=vs.60)?redirectedfrom=MSDN
    'https://blog.csdn.net/u013082684/article/details/47131235
    Set SockObject = CreateObject("MSWinsock.Winsock.1")
    If SockObject Is Nothing Then CheckPort = False: MsgBox "��������ʧ��", vbCritical, "Warning": Exit Function
    With SockObject
        .Protocol = 0 ' TCP
        .RemoteHost = Server
        .RemotePort = Port
        .Connect
        Do While .State = 6 And i < 9 '״̬6��ʾ��������
            DoEvents
            Sleep 250
            i = i + 1 'ѭ���Ĵ���
        Loop
        If (.State = 7) Then
            CheckPort = True '���ӳɹ�
        Else
            CheckPort = False ''.State = 9����ʧ�� /������״̬�����϶�����ʧ��
        End If
        .Close
    End With
    Set SockObject = Nothing
End Function

Function DownloadFileA(ByVal url As String, ByVal FilePath As String) As Boolean '��api�ķ�ʽ�����ļ�
    Dim arrHttp() As Variant, startpos As Long, FileName As String
    Dim j As Integer, t As Long, xi As Variant, ThreadCount As Byte, i As Byte, p As Byte, endpos As Long
    Dim ado As Object, ohttp As Object, Filesize As Long, strx As String, remaindersize As Long, blockSize As Long, upbound As Byte
    
    ThreadCount = 4 '��װ���߳�
    xi = Split(url, "/")
    FileName = xi(UBound(xi)) '��ȡ�ļ���
    FileName = CheckRname(FileName) '������ȡ�����ļ���(��Ϊ���ܰ����Ƿ����ַ�)
    If Left(FilePath, 1) <> "\" Then FilePath = FilePath & "\" '�ļ����λ��
    Set ohttp = CreateObject("msxml2.serverxmlhttp")
    '-------------------------------------------------https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms766431(v=vs.85)?redirectedfrom=MSDN
    Set ado = CreateObject("adodb.stream")
    If ohttp Is Nothing Or ado Is Nothing Then DownloadFileA = False: Exit Function
    With ado
        .type = 1 '���ص��������� adTypeBinary  =1 adTypeText  =2
        .Mode = 3
        .Open
    End With
    
    With ohttp
        .Open "Head", url, True
        .Send
        Do While .readyState <> 4 And j < 256 '��������Ҫע�� ,Ҫ����ѭ���Ĵ���,��ֹ�������γ���ѭ��
            DoEvents
            j = j + 1
            Sleep 25
        Loop
        If .readyState <> 4 Then DownloadFileA = False: GoTo 100
        Filesize = .getResponseHeader("Content-Length") '����ļ���С
        .abort
    End With
    
    strx = FilePath & "TmpFile" '��ʱ�ļ�
    fso.CreateTextFile(strx, True, False).Write (Space(Filesize)) '����һ����С��ͬ�Ŀ��ļ�/��ǰռ�ݴ��̵�λ��
    ado.LoadFromFile (strx)
    blockSize = Fix(Filesize / ThreadCount)
    '--------------------fix����---https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/int-fix-functions
    remaindersize = Filesize - ThreadCount * blockSize
    upbound = ThreadCount - 1

    ReDim arrHttp(upbound) '�������msxml2.xmlhttp���������,����Ա�������ǡ��̡߳���
    
    For i = 0 To upbound
        startpos = i * blockSize
        endpos = (i + 1) * blockSize - 1
        If i = upbound Then endpos = endpos + remaindersize
        Set arrHttp(i) = CreateObject("msxml2.xmlhttp")
        With arrHttp(i)
            .Open "Get", url, True
            '�ֶ�����
            .setRequestHeader "Range", "bytes=" & startpos & "-" & endpos
            .Send
        End With
    Next
    
    Do
        t = timeGetTime
        Do While timeGetTime - t < 300
            DoEvents
            Sleep 25 '��Ҫ��ȫʹ��sleep, sleep�Ὣ�������̶�����
        Loop
        For i = 0 To upbound
            If arrHttp(i).readyState = 4 Then
                'ÿ��ģ��������Ͼͽ���д����ʱ�ļ�����Ӧλ��
                ado.Position = i * blockSize
                ado.Write arrHttp(i).responseBody
                arrHttp(i).abort
                p = p + 1
            End If
        Next
        If p = ThreadCount Then Exit Do
    Loop
    FilePath = FilePath & FileName
    If fso.fileexists(FilePath) Then fso.DeleteFile (FilePath)
    fso.DeleteFile (strx)
    ado.SaveToFile (FilePath)
    '--------------------------https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/savetofile-method-ado?redirectedfrom=MSDN
    DownloadFileA = True
100
    Set otthp = Nothing
    Set ado = Nothing
End Function

Function CheckRname(ByVal FileName As String) As String '���ļ����еķǷ��ַ��滻��
    Dim Char As String, i As Integer, k As Byte, strx As String, strx2 As String, j As Byte
    '---------------------------------------------------------------------------------------�������漰���ļ�������Ҳ���Ե������ģ��
    i = Len(FileName)
    strx = Right$(FileName, i - InStrRev(FileName, ".") + 1) '����״̬,��չ�� ��: http://*.*.*/*.jpg
    If i > 100 Then i = 100 '���Ƴ���(windowsϵͳ֧��200+����ļ���)
    k = i - Len(strx)
    strx2 = Left$(FileName, k)
    For j = 1 To k
        Char = Mid$(strx2, j, 1)
        Select Case asc(Char)
              Case asc("/"), asc("\"), asc(":"), asc("*"), asc("?"), asc("<"), asc(">"), asc("|") 'windows ���Ƶ��ַ�
              Char = "-"
              Mid(strx2, j, 1) = Char '����Щ�ַ�ͳһ�滻��
        End Select
    Next
    CheckRname = strx2 & strx
End Function

Sub ContrlWindowA(ByVal exepath As String, ByVal cmCode As Byte) 'ͨ��sendkey�ķ�ʽ�����ƴ���
    Dim i As Long, strx As String
    'https://docs.microsoft.com/zh-cn/office/vba/api/excel.application.sendkeys
    exepath = exepath & Chr(32)
    i = Shell(exepath, vbMinimizedFocus) 'vbMinimizedFocus�ܶ���򲢲�һ��֧��
    ThisWorkbook.Application.Wait (Now + TimeValue("0:00:01"))
    Select Case cmCode
        Case 1: strx = "% n" '��С��
        Case 2: strx = "% x" '��� '%��ʾ alt��," " ��ʾ��ո��
        Case 3: strx = "% r" '�ָ�
    End Select
'    AppActivate i 'ֱ��ʹ�� i ���ֳ�������(��Ҫ����title)
    ThisWorkbook.Application.SendKeys (strx)
    ' �����Ƕ����еĳ�����Ч, ��Щ��������,�������״̬, sendkey�Ͳ�������
End Sub

'----------------TextBox�Ҽ��˵�����,����,ճ��
Sub Copyx()
    ThisWorkbook.Application.SendKeys "^C"
End Sub

Sub Cutx()
    ThisWorkbook.Application.SendKeys "^X"
End Sub
Sub Pastex()
    ThisWorkbook.Application.SendKeys "^V"
End Sub

Sub TestFileGR(Optional ByVal ic As Integer = 1000) '���ɲ����ļ�, Ĭ������1000���ļ�
    Dim FilePath As String, k As Integer, strx As String, strx1 As String
    Dim fl As Object
    
    strx = ".txt"
    strx1 = Environ("UserProfile") & "\Desktop\test" & Format(Now, "yyyymmddhhmmss")
    fso.CreateFolder strx1
    For k = 1 To ic
        FilePath = strx1 & "\test000" & CStr(k) & strx
        Set fl = fso.CreateTextFile(FilePath, True)
        fl.WriteLine String(1024, 1)
        fl.Close
    Next
    Set fl = Nothing
End Sub

Function CheckIPisOK(ByVal ipaddress As String) As Boolean '���ip�Ƿ����
    Dim objPing As Object, objStatus As Object
    
    Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_PingStatus where address='" & ipaddress & "'")
    For Each objStatus In objPing
        If IsNull(objStatus.StatusCode) Or objStatus.StatusCode <> 0 Then
            CheckIPisOK = False
        Else
            CheckIPisOK = True
        End If
    Next
    Set objPing = Nothing
    Set objStatus = Nothing
End Function

Function ObtainKeyWord(ByVal xtext As String, ByVal strlenx As Byte) As String() '��ȡ�����ؼ���,��������ݵĽ��ж�������,��"office 2013 ����̳�"����"office" & "2013" & "����̳�"
    Dim arr() As String, matchrule As Variant
    Dim i As Byte, k As Byte, p As Byte
    Dim myreg As Object, match As Object, Matches As Object

'    k = Len(xtext)
    k = strlenx
'    If k = 0 Then Exit Sub
    matchrule = Array("[a-zA-Z]{3,}", "[0-9]{2,}", "[\u4e00-\u9fa5]{2,}") 'ƥ�䵥��, ���ȴ��ڵ���3 'ƥ����ĸ,���ȴ��ڵ���2 '���� '���ȴ��ڵ���2
    Set myreg = CreateObject("VBScript.RegExp")
    ReDim arr(k)
    ReDim AbtainKeyWord(k)
    For p = 0 To 2
        With myreg
            .Pattern = matchrule(p) '��ȡ��������
            .Global = True
            .IgnoreCase = True '�����ִ�Сд
            Set Matches = .Execute(xtext)
            For Each match In Matches
                arr(i) = match.Value
                i = i + 1
            Next
        End With
        Set match = Nothing
        Set Matches = Nothing
    Next
    ObtainKeyWord = arr
    Erase arr
    Set myreg = Nothing
End Function

Sub CreateLBworkSheet(ByVal n As Integer) '���������ݸ��Ƹ��Ƶ�csv�ļ�, ����ָ�ʽ��������
    Dim wb As Workbook
    Dim strx As String
    
    strx = ThisWorkbook.Path & "\LBCopy.csv"
    If fso.fileexists(strx) = True Then fso.DeleteFile strx, True
    ThisWorkbook.Application.ScreenUpdating = False
    ThisWorkbook.Application.DisplayAlerts = False
    Set wb = Workbooks.Add
    With wb
        .Sheets(1).Name = "LB"
        ThisWorkbook.Sheets("���").Range("b5:e" & n).Copy .Sheets(1).Cells(1, 1)
        .SaveAs strx
        .Close savechanges:=True
    End With
    Set wb = Nothing
    ThisWorkbook.Application.DisplayAlerts = True
    ThisWorkbook.Application.ScreenUpdating = True
End Sub

Function CreateDB() As Boolean '����access���ݿ�,���ָ����������ݵ����ݿ�
    Dim myCat As New ADOX.Catalog
    Dim FilePath As String
    Dim Elow As Integer
    Dim fl As File
    Dim k As Byte
    
    CreateDB = False
    FilePath = ThisWorkbook.Path & "\LB.accdb"     'ָ�����ݿ�λ��/����
    Elow = ThisWorkbook.Sheets("���").[c65536].End(xlUp).Row
    If Elow = 5 Then CreateDB = False: Exit Function
    If fso.fileexists(FilePath) = True Then
        Set fl = fso.GetFile(FilePath)
        If DateDiff("d", fl.DateCreated, Now) = 0 Then 'https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/datediff-function
            CreateDB = True
            Set fl = Nothing
            Exit Function
        Else
            fso.DeleteFile FilePath
            Set fl = Nothing
        End If
    End If
'    Set Conn = CreateObject("ADODB.Connection")
    myCat.Create "provider=microsoft.ace.oledb.12.0;data source=" & FilePath '�������ݿ��ļ�
    With Conn '�������ݿ�
        .Provider = "microsoft.ace.oledb.12.0"
        .Open FilePath
    End With
    '---------------------����(��ͷ,����(��󳤶�,���ֲ���Ҫָ��),�޶�����)
    SQL = "create table LB(ͳһ��� text(12) not null," & "�ļ��� text(255) not null,�ļ����� text(6) not null," & "�ļ�·�� text(255) not null)"
    Conn.Execute SQL
    '---------------------�������
    SQL = "INSERT INTO LB SELECT f1 as ͳһ���,f2 as �ļ���,f3 as �ļ�����,f4 as �ļ�·�� from [Excel 12.0;hdr=no;Database=" & ThisWorkbook.fullname & ";].[���$b6:e" & Elow & "]" 'f:filed
    Conn.Execute SQL
    '---------------------�������(ָ�����������)
    Conn.Close
    Set Conn = Nothing
    Set myCat = Nothing
    CreateDB = True
End Function

Function CheckUnicode(ByVal strText As String) As Boolean '����Ƿ���Unicode�ַ�
    Dim strx As String
    CheckUnicode = False
    strx = Replace(strText, "?", "") '�Ƚ�"?"����ȥ��
    strx = StrConv(strx, vbFromUnicode) '��תΪansi, Unicode�ַ�����"?"�滻��
    strx = StrConv(strx, vbUnicode) 'ת��Unicode�ַ�
    If InStr(strx, "?") > 0 Then CheckUnicode = True
End Function

Function ObtainMediaLen(ByVal FilePath As Variant) As String '��ȡý���ļ��ĳ���
    Dim FileName As Variant
    Dim obj As Object, fd As Object, fditem As Object
    '-----------------https://docs.microsoft.com/en-us/windows/win32/shell/shell-namespace
    '---------------������Ҫע��,filename, filepath������string���͵�����,������vi���͵�����, �������ִ���
    ' ( ByVal vDir As Variant ) As Folder
    FileName = Right(FilePath, Len(FilePath) - InStrRev(FilePath, "\")) '���ܼ�$����
    FilePath = Left(FilePath, Len(FilePath) - Len(FileName) - 1) '�ļ���
    Set obj = CreateObject("Shell.Application")
    Set fd = obj.Namespace(FilePath)
    Set fditem = fd.ParseName(FileName)
    ObtainMediaLen = fd.getdetailsof(fditem, 27) 'https://docs.microsoft.com/en-us/windows/win32/shell/folder-getdetailsof
    Set obj = Nothing
    Set fd = Nothing
    Set fditem = Nothing
End Function

Sub GetFileMoreDetail(ByVal FilePath As String) '��ȡ������ļ�����Ϣ
    Dim obj As Object, n As Byte

    Set obj = CreateObject("shell.application").Namespace(sPath)
    For n = 0 To 255
            Debug.Print .getdetailsof(, n)
            Debug.Print .getdetailsof(obj, n)
        End If
    Next
Set obj = Nothing
End Sub

Sub MiniAllWindows() '��С�����еĴ���
    Dim objShell As Object
    Set objShell = CreateObject("shell.application")
    objShell.MinimizeAll
    Set objShell = Nothing
End Sub
'0 Open the application with a hidden window.
'1 Open the application with a normal window. If the window is minimized or maximized, the system restores it to its original size and position.
'2 Open the application with a minimized window.
'3 Open the application with a maximized window.
'4 Open the application with its window at its most recent size and position. The active window remains active.
'5 Open the application with its window at its current size and position.
'7 Open the application with a minimized window. The active window remains active.
'10 Open the application with its window in the default state specified by the application.
Sub ShellOpenFile(ByVal FilePath As String, Optional ByVal cmCode As Byte = 1)
    Dim objShell As Object
    Set objShell = CreateObject("shell.application")
    objShell.ShellExecute FilePath, "", "", "open", cmCode
    Set objShell = Nothing
End Sub

Function CheckIsServiceRunning(Optional ByVal servicename As String = "Spooler") As Boolean '�������Ƿ�����, SpoolerΪ��ӡ����, ��pdf��صĳ�������Ҫʹ��
    Dim objShell As Object
    Dim bReturn As Boolean
    Set objShell = CreateObject("shell.application")
    bReturn = objShell.IsServiceRunning(servicename)
    CheckIsServiceRunning = bReturn
    Set objShell = Nothing
End Function

Function msToMinute(ByVal ms As Long) As String '����תΪ����
msToMinute = Format(ms / 1000 / 24 / 60 / 60, "hh:mm:ss") 'ת��Ϊʱ����
End Function

Function getGUID() '����GUID
    getGUID = LCase(Mid$(CreateObject("Scriptlet.TypeLib").GUID, 2, 36))
End Function

Function Get_CPU_name() As String '��ȡCPU����
    Dim objWMIService As Object
    Dim colItems As Object, objItem As Object
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor", , 48)
    For Each objItem In colItems
        Get_CPU_name = objItem.Name: Exit For
    Next
    Set objWMIService = Nothing
    Set colItems = Nothing
End Function

Function Sys_MemoryCapacity() As Long '��ȡϵͳ�ڴ�����
    Dim objItem As Object, objWMIMemory As Object
    Dim i As Long
    Set objWMIMemory = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_PhysicalMemory")
    'objWMIMemory.Count
    For Each objItem In objWMIMemory
        i = i + objItem.Capacity
    Next
    Sys_MemoryCapacity = i
    Set objWMIMemory = Nothing
End Function
