Attribute VB_Name = "调用API"
Option Explicit
'api的使用因为谨慎, 特别是涉及到内存控制的, 如果出现错误将导致潜在的文件被破坏的风险(可能风险不在Excel的保护范围内)
'https://blog.csdn.net/weixin_34334744/article/details/93278317
'https://docs.microsoft.com/zh-cn/windows/win32/api/stringapiset/nf-stringapiset-widechartomultibyte
Private Const CP_UTF8 As Long = 65001
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

'https://docs.microsoft.com/en-us/previous-versions/aa671659(v=vs.71)?redirectedfrom=MSDN
#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long '64位系统 'https://docs.microsoft.com/zh-cn/office/vba/Language/Reference/user-interface-help/ptrsafe-keyword
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
#Else
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
#End If
'-----------------------------------------------------------------------剪切板
Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long '判断数组是否已经初始化'https://stackoverflow.com/questions/183353/how-do-i-determine-if-an-array-is-initialized-in-vb6/444810#444810
'https://docs.microsoft.com/zh-cn/windows/win32/api/oleauto/nf-oleauto-safearraygetdim
'--------------------------------------------------------------------------------------判断数组
Private Declare Function StrFormatByteSize Lib "shlwapi" Alias "StrFormatByteSizeA" (ByVal dw As Long, ByVal pszBuf As String, ByRef cchBuf As Long) As String
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Const OFS_MAXPATHNAME = 128
Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type
Private Const OF_READ = &H0
'-------------------------------------api方式转换文件大小显示方式
Dim hHook As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" _
(ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function GetKeyNameText Lib "user32" Alias "GetKeyNameTextA" (ByVal lParam As Long, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Const WH_KEYBOARD = 2
'----------------------------------控制键盘

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, _
lpSecurityAttributes As String, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const INVALID_HANDLE_VALUE = -1
'------------------------------------------------- '判断程序是否处于运行状态
Private Declare Function InternetOpenA Lib "wininet" _
        (ByVal lpszAgent As String, ByVal dwAccessType As Long, _
        ByVal lpszProxyName As String, ByVal lpszProxyBypass As String, _
        ByVal dwFlags As Long) As Long

Private Declare Function InternetCloseHandle Lib "wininet" _
        (ByVal hinternet As Long) As Long

Private Declare Function InternetOpenUrlA Lib "wininet" _
        (ByVal hinternet As Long, ByVal lpszUrl As String, _
        ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, _
        ByVal dwFlags As Long, ByVal dwContext As Long) As Long
'-------------------------------------------------------------------检测网络是否连接正常
Private Declare Function LCMapString Lib "kernel32" Alias "LCMapStringA" (ByVal Locale As Long, ByVal dwMapFlags As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, ByVal lpDestStr As String, ByVal cchDest As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
'-----------------------------------------------------------------------------------------------------简繁转换

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
   
Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2
'-----------------------------------------------------------------------------------------句柄-窗体
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_HIDE = 0
Private Const SW_MAXIMIZE = 3
Private Const SW_SHOWMINIMIZED = 2

'https://docs.microsoft.com/en-us/windows/win32/debug/imagehlp-reference
'https://baike.baidu.com/item/SearchTreeForFile/2576526
Private Declare Function SearchTreeForFile Lib "ImageHlp.dll" (ByVal lpRoot As String, ByVal lpInPath As String, ByVal lpOutPath As String) As Long '搜索文件
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long '获取磁盘
'getdrivetype的类型
' 5 表示cd-rom/光盘/光驱
' 3 表示固定cipan
' 4 表示远程磁盘(网络端)
' 6 表示内存盘(ram disk)
' 2 表示可移动盘
' http://www.jasinskionline.com/WindowsApi/ref/g/getdrivetype.html
Function SearchFilex(ByVal FileName As String, Optional ByVal diskc As Byte = 15) As String '查找特定文件的所在位置 '注意文件名是需要精确(具体)的文件名如查找
    Dim k As Long, i As Long, SearchPath As String                                          ' 查找C:\Program Files\7-Zip\7z.exe的路径, 需要使用 "7z.exe"这个参数
    'diskc = fso.Drives.Count
    For i = 0 To diskc '检查的磁盘数量, 缺省值15
        SearchPath = Chr$(i + 65) & ":\" '查找盘符, 从A开始
        If GetDriveType(SearchPath) = 3 Then
            SearchFilex = String$(1024, 0) '占位
            k = SearchTreeForFile(SearchPath, FileName, SearchFilex)
            If k <> 0 Then SearchFilex = Split(SearchFilex, Chr(0))(0): Exit Function
        End If
    Next
    SearchFilex = ""
End Function

Sub ContrlWindow(ByVal Classname As String, ByVal cmCode As Byte) '通过句柄来控制窗体的的状态, 通过vb这种方式无法有效控制窗体的状态
    Dim lHwnd As Long
     '获取句柄
     '---------Classname可以通过spy++来获取
     '---------https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-showwindow?redirectedfrom=MSDN
    lHwnd = FindWindow(Classname, vbNullString)
    Select Case cmCode
        Case 1: ShowWindow lHwnd, SW_SHOWMINIMIZED ' 最小化显示
        Case 2: ShowWindow lHwnd, SW_MAXIMIZE '最大化显示
        Case 3: ShowWindow lHwnd, SW_HIDE '隐藏
    End Select
End Sub

Sub GetHwnd() '获取句柄和窗体文本
    Dim hwnd As Long
    Dim iNum As Long
    Dim lpStr As String * 256
    Dim lgh As Long
    hwnd = GetWindow(GetDesktopWindow, GW_CHILD)    '取得第1个子窗口句柄
    Do                                              '循环枚举
        If hwnd = 0 Then Exit Do                    '句柄为0时结束循环
        iNum = iNum + 1                             '子窗口计数
        Cells(iNum, 1) = iNum                       '记录子窗口序号
        Cells(iNum, 2) = hwnd                       '记录子窗口句柄
        lgh = GetWindowText(hwnd, lpStr, 255)       '获取窗口标题文本的长度
        Cells(iNum, 3) = Left(lpStr, lgh)           '记录子窗口标题文本
        hwnd = GetWindow(hwnd, GW_HWNDNEXT)         '查找下一个子窗口
    Loop
End Sub

Sub BlocKey() '控制键盘 '可以用于控制某些控件的内容不被错误修改
    Dim i As Long
    
    i = GetCurrentThreadId
    hHook = SetWindowsHookEx(WH_KEYBOARD, AddressOf KeyBoardProc, 0, i)
End Sub

Function KeyBoardProc(ByVal code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long '键盘控制程序
    If code < 0 Then
        KeyBoardProc = CallNextHookEx(hHook, code, wParam, lParam)
    Else
        '将钩子程序的返回值设置为非0，表示钩子程序处理了这个消息，不继续发送到目标窗口程序
        '就可以屏蔽所有键盘的按键
        KeyBoardProc = 1
    End If
End Function

Sub UnBlocKey() '解锁键盘
    UnhookWindowsHookEx hHook
End Sub



Function GetClipText() As String '获取剪切板的内容
    Dim objdata As New DataObject  '需要引用“Microsoft Forms 2.0 Object Library”-即窗体
    
    objdata.GetFromClipboard
    GetClipText = objdata.GetText
    Set objdata = Nothing
End Function

Function IsExeRun(ByVal pFile As String) As Boolean '检查程序是否处于运行状态-备用
    Dim xi As Long
    
    xi = CreateFile(pFile, GENERIC_READ Or GENERIC_WRITE, 0&, vbNullString, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)
    IsExeRun = (xi = INVALID_HANDLE_VALUE)
    CloseHandle xi
End Function

Function TestURL(url As String) As Boolean '测试网络链接是否正常
    Dim hInt As Long, hInt2 As Long
    
    hInt = InternetOpenA("Excel", 0, vbNullString, _
            vbNullString, &H200000)
    hInt2 = InternetOpenUrlA(hInt, url, vbNullString, 0, 0, 0)
    If hInt2 Then TestURL = True: InternetCloseHandle hInt2
    InternetCloseHandle hInt
End Function

Function SC2TC(ByVal str As String) As String '简转繁
    Dim strLen As Long
    Dim TC As String
    
    strLen = lstrlen(str)
    TC = Space(strLen)
    LCMapString &H804, &H4000000, str, strLen, TC, strLen
    SC2TC = TC
End Function

Function TC2SC(ByVal str As String) As String '繁转简
    Dim strLen As Long
    Dim sc As String
    
    strLen = lstrlen(str)
    sc = Space(strLen)
    LCMapString &H804, &H2000000, str, strLen, sc, strLen
    TC2SC = sc
End Function

Sub ClearClipboard() '清空剪切板
    Dim lngRet As Long
    
    lngRet = OpenClipboard(Application.hwnd)
    If lngRet Then
        EmptyClipboard
        CloseClipboard
    End If
End Sub

Function FormatFileSize(ByVal FilePath As String) As String 'api方法转换文件的大小,注意不支持文文件夹,仅支持单个文件
    Dim FileHandle As Long
    Dim strFileName As String
    Dim leRpOpenBuff As OFSTRUCT
    Dim Amount As Long
    Dim lngFileSize As Long

    Dim Buffer As String
    Dim Result As String
    FileHandle = OpenFile(FilePath, leRpOpenBuff, OF_READ)
    Buffer = Space$(255)
    Amount = GetFileSize(FileHandle, lngFileSize)
    Result = StrFormatByteSize(Amount, Buffer, Len(Buffer))
    If InStr(Result, vbNullChar) > 1 Then FormatFileSize = Left$(Result, InStr(Result, vbNullChar) - 1)
End Function

