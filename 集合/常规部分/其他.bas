Attribute VB_Name = "其他"
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
  '------------------------------------------------网络检查模块
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
'-----------------------------------------语音参数
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) '时间 -单词训练
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Public SpEnd As Boolean
Public SpStart As Boolean 'sapi播放和播放结束控制

Private Declare Function CoCreateGuid Lib "ole32" (id As Any) As Long

Private Function CreateGUID() As String '创建GUID
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

Function TextToVoice(ByVal Oputfile As String, ByVal cmCode As Byte, Optional ByVal Inputfile As String, Optional ByVal strText As String) As Boolean '将文本转换为语音
    Dim oFileStream As Object, oVoice As Object, oFileOpen As Object
    
    'https://docs.microsoft.com/zh-cn/previous-versions/windows/desktop/ms722561(v=vs.85)
    'https://docs.microsoft.com/zh-cn/previous-versions/windows/desktop/ms723602(v=vs.85)
    'https://www.cnblogs.com/sntetwt/p/3533632.html
    'EN 版本的windows 没有CN的语言包,需要额外安装
    'windows7中文本的EN发音效果很烂(也许有其他的语言包?可以改善?)
    '利用这个特性可以快速将文本转换为语音文件, 如单词的词库
    TextToVoice = True
    Set oFileStream = CreateObject("SAPI.SpFileStream")
    oFileStream.Format.type = SAFT48kHz16BitStereo  '输出的音频文件
    '------------------------------------------------仅适用于wav格式的文件
    If fso.fileexists(Oputfile) = True Then fso.DeleteFile Oputfile
    oFileStream.Open Oputfile, SSFMCreateForWrite 'C:\Users\***\Downloads\A111\Sample.wav"
    Set oVoice = CreateObject("SAPI.SpVoice")
    Set oVoice.AudioOutputStream = oFileStream
    If cmCode = 1 Then
        If fso.fileexists(Inputfile) = False Then TextToVoice = False: Set oVoice = Nothing: Exit Function
        Set oFileOpen = CreateObject("SAPI.SpFileStream") '输入的txt文件
        oFileOpen.Open Inputfile, SSFMOpenForRead, False ''C:\Users\***\Downloads\A111\Sample.txt" 没有测试大文件的效果, 如果无法直接读取,可以先将文本的内容提取出来,再将内容转为语音
        oVoice.SpeakStream oFileOpen '注意这里不是.speak, speak 后面是string类型的数据 ,如 speak "hello, world"
        oFileOpen.Close
        Set oFileOpen = Nothing
    Else
        oVoice.Speak strText
    End If
    oFileStream.Close
    Set oFileStream = Nothing
    Set oVoice = Nothing
End Function

Function IsNetConnectOnline() As Boolean '更好的网络检查连接方法
    IsNetConnectOnline = InternetGetConnectedState(0&, 0&)
End Function

Function GetOsVersion() As String '获取系统的版本信息
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

Function WmiCheckFileOpen(ByVal FilePath As String, Optional ByVal exename As String) As Boolean '判断文件是否处于打开的状态
    Dim strComputer As String, commandline As String
    Dim objWMIService As Object, colItems As Object, objItem As Object
    '------------------------------------------------------------------这种方法的好处就是避免了open方法无法访问密码保护的文件, 而且也可以解决txt文件的问题
    '------------------------------------------------------------------无法处理excel关联的文件
    strComputer = "."
    WmiCheckFileOpen = False
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
    If objWMIService Is Nothing Then MsgBox "无法创建对象": Exit Function
    If Len(exename) = 0 Then
        commandline = "select * from win32_process"   '根据不同的任务进行筛选
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

Sub TerminateEXEs(ByVal exename As String) '终止特定进程
    Dim obj As Object, targetexe As Object, targetexex As Object
                                                            'https://docs.microsoft.com/en-us/windows/win32/wmisdk/wmi-tasks--processes
    On Error GoTo 100
    Set obj = GetObject("winmgmts:\\.\root\cimv2")          'https://docs.microsoft.com/en-us/windows/win32/wmisdk/connecting-to-wmi-with-vbscript
    If obj Is Nothing Then MsgBox "无法创建对象": Exit Sub
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
    '不建议直接终止程序,可以查看文件占用的程序名,直接终止程序可能导致数据丢失的问题
End Sub

Function IsPing(strMachines As String) As Boolean '检查对应的设备是否处于连接的状态
    Dim aMachines() As String
    Dim machine As Variant
    Dim objPing As Object
    Dim objStatus As Object
    
    aMachines = Split(strMachines, ";")
    For Each machine In aMachines
        Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_PingStatus where address = '" & machine & "'")
        If objPing Is Nothing Then IsPing = False: Exit Function '如果没成功创建对象
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
    pyArr = [{"吖","A";"八","B";"攃","C";"咑","D";"妸","E";"发","F";"旮","G";"哈","H";"丌","J";"咔","K";"垃","L";"妈","M";"乸","N";"噢","O";"帊","P";"七","Q";"冄","R";"仨","S";"他","T";"屲","W";"夕","X";"丫","Y";"帀","Z"}]
    str = Replace(Replace(textchar, " ", ""), "　", "") '替换掉空格
    k = Len(str)
    For i = 1 To k
        ch = Mid(str, i, 1)
        If ch Like "[一-龥]" Then   '如果是汉字，进行转换
            GetPYA = GetPYA & WorksheetFunction.Lookup(Mid(str, i, 1), pyArr)
        Else
            GetPYA = GetPYA & UCase(ch)     '如果不是汉字，直接输出
        End If
    Next
End Function

Function GetPY(str As String) As String '获取汉字拼音的首字母(这个功能可用以在搜索时,针对拼音首字母的模糊搜索)
    Dim i As Integer
    
    For i = 0 To Len(str) - 1
        GetPY = GetPY & _
        IIf(IsChinese(asc(Mid(str, i + 1, 1))), _
        GetPYChar(Mid(str, i + 1, 1)), "")
    Next
    GetPY = LCase(GetPY)
End Function

Private Function IsChinese(ByVal AscVal As Integer) As Boolean ''判断某个ASC码是否指向一个汉字
    IsChinese = IIf(Len(Hex(AscVal)) > 2, True, False)
End Function

Private Function GetPYChar(Char As String) As String ''获取对应汉字首字母(出现不少遗漏)
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

'--------------------------------处理文本内容的示例,用于处理hosts文件的内容更新
Sub DataClean() '清洗数据
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
'读取文本,将文本的数据读取到Excel中, 利用Excel的去重,再将两部分内容合并,再将内容重新写入hosts文件
Sub WriteHosts()            '读取/写入hosts的数据
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
        Open strx1 For Input As #1          '将文本的内容全部读取出来,按行分开
        arr = Split(StrConv(InputB(LOF(1), 1), vbUnicode), vbNewLine)
        Close #1
        i = UBound(arr)
        ReDim arrtempx(i)
        For k = 0 To i
            arrtempx(k) = Trim(Split(arr(k), Chr(32))(1))
        Next
        .Cells(m + 1, 1).Resize(i + 1, 1) = Application.Transpose(arrtempx)
        m = .[a65536].End(xlUp).Row
        .Range("a267:a" & m).RemoveDuplicates Columns:=1, Header:=xlNo '直接利用Excel自带的去重
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
'-----------------------------------------------hosts处理

Function CheckPort(ByVal Server As String, ByVal Port As Long) As Boolean '检查远端服务器的端口是否处于可用的状态
    Dim SockObject As Object, i As Byte
    '需要引用MSWINSCK.OCX,system32下的这个控件(x86),win7支持
    'https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa733709(v=vs.60)?redirectedfrom=MSDN
    'https://blog.csdn.net/u013082684/article/details/47131235
    Set SockObject = CreateObject("MSWinsock.Winsock.1")
    If SockObject Is Nothing Then CheckPort = False: MsgBox "创建对象失败", vbCritical, "Warning": Exit Function
    With SockObject
        .Protocol = 0 ' TCP
        .RemoteHost = Server
        .RemotePort = Port
        .Connect
        Do While .State = 6 And i < 9 '状态6表示正在链接
            DoEvents
            Sleep 250
            i = i + 1 '循环的次数
        Loop
        If (.State = 7) Then
            CheckPort = True '链接成功
        Else
            CheckPort = False ''.State = 9连接失败 /其他的状态连不上都归于失败
        End If
        .Close
    End With
    Set SockObject = Nothing
End Function

Function DownloadFileA(ByVal url As String, ByVal FilePath As String) As Boolean '非api的方式下载文件
    Dim arrHttp() As Variant, startpos As Long, FileName As String
    Dim j As Integer, t As Long, xi As Variant, ThreadCount As Byte, i As Byte, p As Byte, endpos As Long
    Dim ado As Object, ohttp As Object, Filesize As Long, strx As String, remaindersize As Long, blockSize As Long, upbound As Byte
    
    ThreadCount = 4 '假装多线程
    xi = Split(url, "/")
    FileName = xi(UBound(xi)) '获取文件名
    FileName = CheckRname(FileName) '修正获取到的文件名(因为可能包含非法的字符)
    If Left(FilePath, 1) <> "\" Then FilePath = FilePath & "\" '文件存放位置
    Set ohttp = CreateObject("msxml2.serverxmlhttp")
    '-------------------------------------------------https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms766431(v=vs.85)?redirectedfrom=MSDN
    Set ado = CreateObject("adodb.stream")
    If ohttp Is Nothing Or ado Is Nothing Then DownloadFileA = False: Exit Function
    With ado
        .type = 1 '返回的数据类型 adTypeBinary  =1 adTypeText  =2
        .Mode = 3
        .Open
    End With
    
    With ohttp
        .Open "Head", url, True
        .Send
        Do While .readyState <> 4 And j < 256 '在这里需要注意 ,要限制循环的次数,防止在这里形成死循环
            DoEvents
            j = j + 1
            Sleep 25
        Loop
        If .readyState <> 4 Then DownloadFileA = False: GoTo 100
        Filesize = .getResponseHeader("Content-Length") '获得文件大小
        .abort
    End With
    
    strx = FilePath & "TmpFile" '临时文件
    fso.CreateTextFile(strx, True, False).Write (Space(Filesize)) '创建一个大小相同的空文件/提前占据磁盘的位置
    ado.LoadFromFile (strx)
    blockSize = Fix(Filesize / ThreadCount)
    '--------------------fix函数---https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/int-fix-functions
    remaindersize = Filesize - ThreadCount * blockSize
    upbound = ThreadCount - 1

    ReDim arrHttp(upbound) '定义包含msxml2.xmlhttp对象的数组,·成员数量便是“线程”数
    
    For i = 0 To upbound
        startpos = i * blockSize
        endpos = (i + 1) * blockSize - 1
        If i = upbound Then endpos = endpos + remaindersize
        Set arrHttp(i) = CreateObject("msxml2.xmlhttp")
        With arrHttp(i)
            .Open "Get", url, True
            '分段下载
            .setRequestHeader "Range", "bytes=" & startpos & "-" & endpos
            .Send
        End With
    Next
    
    Do
        t = timeGetTime
        Do While timeGetTime - t < 300
            DoEvents
            Sleep 25 '不要完全使用sleep, sleep会将整个进程都挂起
        Loop
        For i = 0 To upbound
            If arrHttp(i).readyState = 4 Then
                '每个模块下载完毕就将其写入临时文件的相应位置
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

Function CheckRname(ByVal FileName As String) As String '将文件名中的非法字符替换掉
    Dim Char As String, i As Integer, k As Byte, strx As String, strx2 As String, j As Byte
    '---------------------------------------------------------------------------------------其他的涉及到文件命名的也可以调用这个模块
    i = Len(FileName)
    strx = Right$(FileName, i - InStrRev(FileName, ".") + 1) '理想状态,扩展名 如: http://*.*.*/*.jpg
    If i > 100 Then i = 100 '限制长度(windows系统支持200+的最长文件名)
    k = i - Len(strx)
    strx2 = Left$(FileName, k)
    For j = 1 To k
        Char = Mid$(strx2, j, 1)
        Select Case asc(Char)
              Case asc("/"), asc("\"), asc(":"), asc("*"), asc("?"), asc("<"), asc(">"), asc("|") 'windows 限制的字符
              Char = "-"
              Mid(strx2, j, 1) = Char '把这些字符统一替换掉
        End Select
    Next
    CheckRname = strx2 & strx
End Function

Sub ContrlWindowA(ByVal exepath As String, ByVal cmCode As Byte) '通过sendkey的方式来控制窗体
    Dim i As Long, strx As String
    'https://docs.microsoft.com/zh-cn/office/vba/api/excel.application.sendkeys
    exepath = exepath & Chr(32)
    i = Shell(exepath, vbMinimizedFocus) 'vbMinimizedFocus很多程序并不一定支持
    ThisWorkbook.Application.Wait (Now + TimeValue("0:00:01"))
    Select Case cmCode
        Case 1: strx = "% n" '最小化
        Case 2: strx = "% x" '最大化 '%表示 alt键," " 表示大空格键
        Case 3: strx = "% r" '恢复
    End Select
'    AppActivate i '直接使用 i 部分程序会出错(需要换成title)
    ThisWorkbook.Application.SendKeys (strx)
    ' 并不是对所有的程序都有效, 有些程序开启后,进入输出状态, sendkey就不起作用
End Sub

'----------------TextBox右键菜单复制,剪切,粘贴
Sub Copyx()
    ThisWorkbook.Application.SendKeys "^C"
End Sub

Sub Cutx()
    ThisWorkbook.Application.SendKeys "^X"
End Sub
Sub Pastex()
    ThisWorkbook.Application.SendKeys "^V"
End Sub

Sub TestFileGR(Optional ByVal ic As Integer = 1000) '生成测试文件, 默认生成1000个文件
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

Function CheckIPisOK(ByVal ipaddress As String) As Boolean '检测ip是否可用
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

Function ObtainKeyWord(ByVal xtext As String, ByVal strlenx As Byte) As String() '获取搜索关键词,将混合内容的进行独立分离,如"office 2013 经典教程"拆解成"office" & "2013" & "经典教程"
    Dim arr() As String, matchrule As Variant
    Dim i As Byte, k As Byte, p As Byte
    Dim myreg As Object, match As Object, Matches As Object

'    k = Len(xtext)
    k = strlenx
'    If k = 0 Then Exit Sub
    matchrule = Array("[a-zA-Z]{3,}", "[0-9]{2,}", "[\u4e00-\u9fa5]{2,}") '匹配单词, 长度大于等于3 '匹配字母,长度大于等于2 '中文 '长度大于等于2
    Set myreg = CreateObject("VBScript.RegExp")
    ReDim arr(k)
    ReDim AbtainKeyWord(k)
    For p = 0 To 2
        With myreg
            .Pattern = matchrule(p) '获取豆瓣评分
            .Global = True
            .IgnoreCase = True '不区分大小写
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

Sub CreateLBworkSheet(ByVal n As Integer) '将书库的内容复制复制到csv文件, 会出现格式兼容问题
    Dim wb As Workbook
    Dim strx As String
    
    strx = ThisWorkbook.Path & "\LBCopy.csv"
    If fso.fileexists(strx) = True Then fso.DeleteFile strx, True
    ThisWorkbook.Application.ScreenUpdating = False
    ThisWorkbook.Application.DisplayAlerts = False
    Set wb = Workbooks.Add
    With wb
        .Sheets(1).Name = "LB"
        ThisWorkbook.Sheets("书库").Range("b5:e" & n).Copy .Sheets(1).Cells(1, 1)
        .SaveAs strx
        .Close savechanges:=True
    End With
    Set wb = Nothing
    ThisWorkbook.Application.DisplayAlerts = True
    ThisWorkbook.Application.ScreenUpdating = True
End Sub

Function CreateDB() As Boolean '创建access数据库,添加指定区域的数据到数据库
    Dim myCat As New ADOX.Catalog
    Dim FilePath As String
    Dim Elow As Integer
    Dim fl As File
    Dim k As Byte
    
    CreateDB = False
    FilePath = ThisWorkbook.Path & "\LB.accdb"     '指定数据库位置/名称
    Elow = ThisWorkbook.Sheets("书库").[c65536].End(xlUp).Row
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
    myCat.Create "provider=microsoft.ace.oledb.12.0;data source=" & FilePath '创建数据库文件
    With Conn '连接数据库
        .Provider = "microsoft.ace.oledb.12.0"
        .Open FilePath
    End With
    '---------------------表名(表头,类型(最大长度,数字不需要指定),限定条件)
    SQL = "create table LB(统一编号 text(12) not null," & "文件名 text(255) not null,文件类型 text(6) not null," & "文件路径 text(255) not null)"
    Conn.Execute SQL
    '---------------------创建表格
    SQL = "INSERT INTO LB SELECT f1 as 统一编号,f2 as 文件名,f3 as 文件类型,f4 as 文件路径 from [Excel 12.0;hdr=no;Database=" & ThisWorkbook.fullname & ";].[书库$b6:e" & Elow & "]" 'f:filed
    Conn.Execute SQL
    '---------------------添加数据(指定区域的数据)
    Conn.Close
    Set Conn = Nothing
    Set myCat = Nothing
    CreateDB = True
End Function

Function CheckUnicode(ByVal strText As String) As Boolean '检查是否有Unicode字符
    Dim strx As String
    CheckUnicode = False
    strx = Replace(strText, "?", "") '先将"?"符号去掉
    strx = StrConv(strx, vbFromUnicode) '先转为ansi, Unicode字符将被"?"替换掉
    strx = StrConv(strx, vbUnicode) '转会Unicode字符
    If InStr(strx, "?") > 0 Then CheckUnicode = True
End Function

Function ObtainMediaLen(ByVal FilePath As Variant) As String '获取媒体文件的长度
    Dim FileName As Variant
    Dim obj As Object, fd As Object, fditem As Object
    '-----------------https://docs.microsoft.com/en-us/windows/win32/shell/shell-namespace
    '---------------这里需要注意,filename, filepath不能是string类型的数据,必须是vi类型的数据, 否则会出现错误
    ' ( ByVal vDir As Variant ) As Folder
    FileName = Right(FilePath, Len(FilePath) - InStrRev(FilePath, "\")) '不能加$符号
    FilePath = Left(FilePath, Len(FilePath) - Len(FileName) - 1) '文件夹
    Set obj = CreateObject("Shell.Application")
    Set fd = obj.Namespace(FilePath)
    Set fditem = fd.ParseName(FileName)
    ObtainMediaLen = fd.getdetailsof(fditem, 27) 'https://docs.microsoft.com/en-us/windows/win32/shell/folder-getdetailsof
    Set obj = Nothing
    Set fd = Nothing
    Set fditem = Nothing
End Function

Sub GetFileMoreDetail(ByVal FilePath As String) '获取更多的文件的信息
    Dim obj As Object, n As Byte

    Set obj = CreateObject("shell.application").Namespace(sPath)
    For n = 0 To 255
            Debug.Print .getdetailsof(, n)
            Debug.Print .getdetailsof(obj, n)
        End If
    Next
Set obj = Nothing
End Sub

Sub MiniAllWindows() '最小化所有的窗口
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

Function CheckIsServiceRunning(Optional ByVal servicename As String = "Spooler") As Boolean '检测服务是否启动, Spooler为打印服务, 在pdf相关的程序中需要使用
    Dim objShell As Object
    Dim bReturn As Boolean
    Set objShell = CreateObject("shell.application")
    bReturn = objShell.IsServiceRunning(servicename)
    CheckIsServiceRunning = bReturn
    Set objShell = Nothing
End Function

Function msToMinute(ByVal ms As Long) As String '毫秒转为分钟
msToMinute = Format(ms / 1000 / 24 / 60 / 60, "hh:mm:ss") '转换为时分秒
End Function

Function getGUID() '生成GUID
    getGUID = LCase(Mid$(CreateObject("Scriptlet.TypeLib").GUID, 2, 36))
End Function

Function Get_CPU_name() As String '获取CPU名称
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

Function Sys_MemoryCapacity() As Long '获取系统内存容量
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
