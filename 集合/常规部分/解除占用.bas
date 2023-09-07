Attribute VB_Name = "解除占用"
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

Sub KillProcess(ByVal FilePath As String, ByVal exepath As String) '通过微软的handle.exe来获取文件的占用,使用cmd结束占用进程
    Dim hwnd As Long
    Dim hMenu As Long, hSubMenu As Long, lSelectAllID As Long, lCopyID As Long
    Dim lPasteID As Long
    Dim strx As String
    Dim strx1 As String
    Dim strx2 As String, pid As Long
    '------------------------------cmd shell之后和其他的进程有些差异
    '------------------------------https://docs.microsoft.com/zh-cn/sysinternals/downloads/handle
    '-----------------------------handle.exe文件可以单独从微软的网站上下载,或者是工具包-SysinternalsSuite,也包含此文件
    '-----------------------------https://blog.csdn.net/weixin_42124234/article/details/98076625
    pid = Shell("cmd /k", vbHide) 'cmd /k, 执行完成后保留cmd窗体, 获取执行cmd的pid
    strx2 = " -a -u "
    strx1 = FilePath
    strx1 = """" & strx1 & """"
    strx = exepath
    strx = """" & strx & """"
    ThisWorkbook.Application.Wait (Now + TimeValue("0:00:01")) '等待cmd窗体打开
    hwnd = FindWindow("ConsoleWindowClass", vbNullString)  '获取cmd窗口句柄
    If hwnd = 0 Then Exit Sub                              '如果获取到0，说明cmd窗口未打开，则退出程序
    
    '------------------------------------------------------------------通过DataObject的方法将内容复制到剪切板, 不支持非ansi字符
    'dim d As Object
'    Set d = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
'    d.Clear
'    d.SetText strx & strx2 & strx1 & vbCrLf
'    d.PutInClipboard
    '由于没办法直接用shell命令执行handle(直接执行或遭到很多麻烦, 假如handle.exe的路径存在空格, 需要检测的文件的路径存在空格等情况, 没法完全消除干扰,加上多个双引号,也还是有问题)
    SetClipboard strx '改成api的方式复制命令
    '---------------------------------------------------------'利用句柄-剪切板-复制-粘贴的方式,间接执行
    hMenu = GetSystemMenu(hwnd, False)                     '获取cmd窗口系统菜单的句柄
    hSubMenu = GetSubMenu(hMenu, 7)                        '获取“编辑”子菜单的句柄
    lPasteID = GetMenuItemID(hSubMenu, 2)                  '获取“粘贴”菜单项的ID
    SendMessage hwnd, WM_SYSCOMMAND, lPasteID, ByVal 0     '向cmd窗口发送粘贴命令，因为以回车结束，所以相应dos指令会被执行
    Application.Wait (Now + TimeValue("0:00:10"))
    '------------------------------------------------------由于handle.exe执行的速度很慢, 需要等待获取执行结果
    lSelectAllID = GetMenuItemID(hSubMenu, 3)              '获取“全选”菜单项的ID
    SendMessage hwnd, WM_SYSCOMMAND, lSelectAllID, ByVal 0 '向cmd窗口发送全选命令
    lCopyID = GetMenuItemID(hSubMenu, 1)                   '获取“复制”菜单项的ID
    SendMessage hwnd, WM_SYSCOMMAND, lCopyID, ByVal 0      '向cmd窗口发送复制命令
    GetClipboard
    '-------------------------------这里需要处理返回结果, 假如需要获取执行的结果或者是强制解除占用
    '(这里还没处理执行结果, 需要调用此sub的需要自己增加这部分的内容)
    '------------------------------不建议直接强制用taskkill结束占用进程,这将对导致信息丢失的可能, 可以返回执行结果或者是占用的pid,然后手动退出进程
    'Shell ("cmd /c taskkill /pid " & pid & " /f"), vbHide '利用cmd taskkill /f强制的方式将之前执行的cmd窗体关闭
End Sub

