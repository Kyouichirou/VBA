Attribute VB_Name = "����API"
Option Explicit
'api��ʹ����Ϊ����, �ر����漰���ڴ���Ƶ�, ������ִ��󽫵���Ǳ�ڵ��ļ����ƻ��ķ���(���ܷ��ղ���Excel�ı�����Χ��)
'https://blog.csdn.net/weixin_34334744/article/details/93278317
'https://docs.microsoft.com/zh-cn/windows/win32/api/stringapiset/nf-stringapiset-widechartomultibyte
Private Const CP_UTF8 As Long = 65001
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

'https://docs.microsoft.com/en-us/previous-versions/aa671659(v=vs.71)?redirectedfrom=MSDN
#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long '64λϵͳ 'https://docs.microsoft.com/zh-cn/office/vba/Language/Reference/user-interface-help/ptrsafe-keyword
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
#Else
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
#End If
'-----------------------------------------------------------------------���а�
Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long '�ж������Ƿ��Ѿ���ʼ��'https://stackoverflow.com/questions/183353/how-do-i-determine-if-an-array-is-initialized-in-vb6/444810#444810
'https://docs.microsoft.com/zh-cn/windows/win32/api/oleauto/nf-oleauto-safearraygetdim
'--------------------------------------------------------------------------------------�ж�����
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
'-------------------------------------api��ʽת���ļ���С��ʾ��ʽ
Dim hHook As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" _
(ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function GetKeyNameText Lib "user32" Alias "GetKeyNameTextA" (ByVal lParam As Long, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Const WH_KEYBOARD = 2
'----------------------------------���Ƽ���

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, _
lpSecurityAttributes As String, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const INVALID_HANDLE_VALUE = -1
'------------------------------------------------- '�жϳ����Ƿ�������״̬
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
'-------------------------------------------------------------------��������Ƿ���������
Private Declare Function LCMapString Lib "kernel32" Alias "LCMapStringA" (ByVal Locale As Long, ByVal dwMapFlags As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, ByVal lpDestStr As String, ByVal cchDest As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
'-----------------------------------------------------------------------------------------------------��ת��

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
   
Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2
'-----------------------------------------------------------------------------------------���-����
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_HIDE = 0
Private Const SW_MAXIMIZE = 3
Private Const SW_SHOWMINIMIZED = 2

'https://docs.microsoft.com/en-us/windows/win32/debug/imagehlp-reference
'https://baike.baidu.com/item/SearchTreeForFile/2576526
Private Declare Function SearchTreeForFile Lib "ImageHlp.dll" (ByVal lpRoot As String, ByVal lpInPath As String, ByVal lpOutPath As String) As Long '�����ļ�
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long '��ȡ����
'getdrivetype������
' 5 ��ʾcd-rom/����/����
' 3 ��ʾ�̶�cipan
' 4 ��ʾԶ�̴���(�����)
' 6 ��ʾ�ڴ���(ram disk)
' 2 ��ʾ���ƶ���
' http://www.jasinskionline.com/WindowsApi/ref/g/getdrivetype.html
Function SearchFilex(ByVal FileName As String, Optional ByVal diskc As Byte = 15) As String '�����ض��ļ�������λ�� 'ע���ļ�������Ҫ��ȷ(����)���ļ��������
    Dim k As Long, i As Long, SearchPath As String                                          ' ����C:\Program Files\7-Zip\7z.exe��·��, ��Ҫʹ�� "7z.exe"�������
    'diskc = fso.Drives.Count
    For i = 0 To diskc '���Ĵ�������, ȱʡֵ15
        SearchPath = Chr$(i + 65) & ":\" '�����̷�, ��A��ʼ
        If GetDriveType(SearchPath) = 3 Then
            SearchFilex = String$(1024, 0) 'ռλ
            k = SearchTreeForFile(SearchPath, FileName, SearchFilex)
            If k <> 0 Then SearchFilex = Split(SearchFilex, Chr(0))(0): Exit Function
        End If
    Next
    SearchFilex = ""
End Function

Sub ContrlWindow(ByVal Classname As String, ByVal cmCode As Byte) 'ͨ����������ƴ���ĵ�״̬, ͨ��vb���ַ�ʽ�޷���Ч���ƴ����״̬
    Dim lHwnd As Long
     '��ȡ���
     '---------Classname����ͨ��spy++����ȡ
     '---------https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-showwindow?redirectedfrom=MSDN
    lHwnd = FindWindow(Classname, vbNullString)
    Select Case cmCode
        Case 1: ShowWindow lHwnd, SW_SHOWMINIMIZED ' ��С����ʾ
        Case 2: ShowWindow lHwnd, SW_MAXIMIZE '�����ʾ
        Case 3: ShowWindow lHwnd, SW_HIDE '����
    End Select
End Sub

Sub GetHwnd() '��ȡ����ʹ����ı�
    Dim hwnd As Long
    Dim iNum As Long
    Dim lpStr As String * 256
    Dim lgh As Long
    hwnd = GetWindow(GetDesktopWindow, GW_CHILD)    'ȡ�õ�1���Ӵ��ھ��
    Do                                              'ѭ��ö��
        If hwnd = 0 Then Exit Do                    '���Ϊ0ʱ����ѭ��
        iNum = iNum + 1                             '�Ӵ��ڼ���
        Cells(iNum, 1) = iNum                       '��¼�Ӵ������
        Cells(iNum, 2) = hwnd                       '��¼�Ӵ��ھ��
        lgh = GetWindowText(hwnd, lpStr, 255)       '��ȡ���ڱ����ı��ĳ���
        Cells(iNum, 3) = Left(lpStr, lgh)           '��¼�Ӵ��ڱ����ı�
        hwnd = GetWindow(hwnd, GW_HWNDNEXT)         '������һ���Ӵ���
    Loop
End Sub

Sub BlocKey() '���Ƽ��� '�������ڿ���ĳЩ�ؼ������ݲ��������޸�
    Dim i As Long
    
    i = GetCurrentThreadId
    hHook = SetWindowsHookEx(WH_KEYBOARD, AddressOf KeyBoardProc, 0, i)
End Sub

Function KeyBoardProc(ByVal code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long '���̿��Ƴ���
    If code < 0 Then
        KeyBoardProc = CallNextHookEx(hHook, code, wParam, lParam)
    Else
        '�����ӳ���ķ���ֵ����Ϊ��0����ʾ���ӳ������������Ϣ�����������͵�Ŀ�괰�ڳ���
        '�Ϳ����������м��̵İ���
        KeyBoardProc = 1
    End If
End Function

Sub UnBlocKey() '��������
    UnhookWindowsHookEx hHook
End Sub



Function GetClipText() As String '��ȡ���а������
    Dim objdata As New DataObject  '��Ҫ���á�Microsoft Forms 2.0 Object Library��-������
    
    objdata.GetFromClipboard
    GetClipText = objdata.GetText
    Set objdata = Nothing
End Function

Function IsExeRun(ByVal pFile As String) As Boolean '�������Ƿ�������״̬-����
    Dim xi As Long
    
    xi = CreateFile(pFile, GENERIC_READ Or GENERIC_WRITE, 0&, vbNullString, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)
    IsExeRun = (xi = INVALID_HANDLE_VALUE)
    CloseHandle xi
End Function

Function TestURL(url As String) As Boolean '�������������Ƿ�����
    Dim hInt As Long, hInt2 As Long
    
    hInt = InternetOpenA("Excel", 0, vbNullString, _
            vbNullString, &H200000)
    hInt2 = InternetOpenUrlA(hInt, url, vbNullString, 0, 0, 0)
    If hInt2 Then TestURL = True: InternetCloseHandle hInt2
    InternetCloseHandle hInt
End Function

Function SC2TC(ByVal str As String) As String '��ת��
    Dim strLen As Long
    Dim TC As String
    
    strLen = lstrlen(str)
    TC = Space(strLen)
    LCMapString &H804, &H4000000, str, strLen, TC, strLen
    SC2TC = TC
End Function

Function TC2SC(ByVal str As String) As String '��ת��
    Dim strLen As Long
    Dim sc As String
    
    strLen = lstrlen(str)
    sc = Space(strLen)
    LCMapString &H804, &H2000000, str, strLen, sc, strLen
    TC2SC = sc
End Function

Sub ClearClipboard() '��ռ��а�
    Dim lngRet As Long
    
    lngRet = OpenClipboard(Application.hwnd)
    If lngRet Then
        EmptyClipboard
        CloseClipboard
    End If
End Sub

Function FormatFileSize(ByVal FilePath As String) As String 'api����ת���ļ��Ĵ�С,ע�ⲻ֧�����ļ���,��֧�ֵ����ļ�
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

