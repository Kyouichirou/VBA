Attribute VB_Name = "���ռ��"
Option Explicit
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
'-https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getsystemmenu
'-Enables the application to access the window menu (also known as the system menu or the control menu) for copying and modifying.
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
'Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
'Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112

Sub KillProcess(ByVal FilePath As String, ByVal exepath As String) 'ͨ��΢���handle.exe����ȡ�ļ���ռ��,ʹ��cmd����ռ�ý���
    Dim hwnd As Long
    Dim hMenu As Long, hSubMenu As Long, lSelectAllID As Long, lCopyID As Long
    Dim lPasteID As Long
    Dim strx As String
    Dim strx1 As String
    Dim strx2 As String, pid As Long
    '------------------------------cmd shell֮��������Ľ�����Щ����
    '------------------------------https://docs.microsoft.com/zh-cn/sysinternals/downloads/handle
    '-----------------------------handle.exe�ļ����Ե�����΢�����վ������,�����ǹ��߰�-SysinternalsSuite,Ҳ�������ļ�
    '-----------------------------https://blog.csdn.net/weixin_42124234/article/details/98076625
    pid = Shell("cmd /k", vbHide) 'cmd /k, ִ����ɺ���cmd����, ��ȡִ��cmd��pid
    strx2 = " -a -u "
    strx1 = FilePath
    strx1 = """" & strx1 & """"
    strx = exepath
    strx = """" & strx & """"
    ThisWorkbook.Application.Wait (Now + TimeValue("0:00:01")) '�ȴ�cmd�����
    hwnd = FindWindow("ConsoleWindowClass", vbNullString)  '��ȡcmd���ھ��
    If hwnd = 0 Then Exit Sub                              '�����ȡ��0��˵��cmd����δ�򿪣����˳�����
    
    '------------------------------------------------------------------ͨ��DataObject�ķ��������ݸ��Ƶ����а�, ��֧�ַ�ansi�ַ�
    'dim d As Object
'    Set d = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
'    d.Clear
'    d.SetText strx & strx2 & strx1 & vbCrLf
'    d.PutInClipboard
    '����û�취ֱ����shell����ִ��handle(ֱ��ִ�л��⵽�ܶ��鷳, ����handle.exe��·�����ڿո�, ��Ҫ�����ļ���·�����ڿո�����, û����ȫ��������,���϶��˫����,Ҳ����������)
    SetClipboard strx '�ĳ�api�ķ�ʽ��������
    '---------------------------------------------------------'���þ��-���а�-����-ճ���ķ�ʽ,���ִ��
    hMenu = GetSystemMenu(hwnd, False)                     '��ȡcmd����ϵͳ�˵��ľ��
    hSubMenu = GetSubMenu(hMenu, 7)                        '��ȡ���༭���Ӳ˵��ľ��
    lPasteID = GetMenuItemID(hSubMenu, 2)                  '��ȡ��ճ�����˵����ID
    SendMessage hwnd, WM_SYSCOMMAND, lPasteID, ByVal 0     '��cmd���ڷ���ճ�������Ϊ�Իس�������������Ӧdosָ��ᱻִ��
    Application.Wait (Now + TimeValue("0:00:10"))
    '------------------------------------------------------����handle.exeִ�е��ٶȺ���, ��Ҫ�ȴ���ȡִ�н��
    lSelectAllID = GetMenuItemID(hSubMenu, 3)              '��ȡ��ȫѡ���˵����ID
    SendMessage hwnd, WM_SYSCOMMAND, lSelectAllID, ByVal 0 '��cmd���ڷ���ȫѡ����
    lCopyID = GetMenuItemID(hSubMenu, 1)                   '��ȡ�����ơ��˵����ID
    SendMessage hwnd, WM_SYSCOMMAND, lCopyID, ByVal 0      '��cmd���ڷ��͸�������
    GetClipboard
    '-------------------------------������Ҫ�����ؽ��, ������Ҫ��ȡִ�еĽ��������ǿ�ƽ��ռ��
    '(���ﻹû����ִ�н��, ��Ҫ���ô�sub����Ҫ�Լ������ⲿ�ֵ�����)
    '------------------------------������ֱ��ǿ����taskkill����ռ�ý���,�⽫�Ե�����Ϣ��ʧ�Ŀ���, ���Է���ִ�н��������ռ�õ�pid,Ȼ���ֶ��˳�����
    'Shell ("cmd /c taskkill /pid " & pid & " /f"), vbHide '����cmd taskkill /fǿ�Ƶķ�ʽ��֮ǰִ�е�cmd����ر�
End Sub

