Attribute VB_Name = "�ļ�x64"
Option Explicit
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lLonpProcName As String) As Long
Private Declare Function Wow64DisableWow64FsRedirection Lib "kernel32" (ByRef oldvalue As Long) As Boolean
Private Declare Function Wow64RevertWow64FsRedirection Lib "kernel32" (ByVal oldvalue As Long) As Boolean

'https://www.samlogic.net/articles/sysnative-folder-64-bit-windows.html
'https://docs.microsoft.com/en-us/windows/win32/api/wow64apiset/nf-wow64apiset-wow64revertwow64fsredirection
'https://docs.microsoft.com/en-us/windows/win32/winprog64/file-system-redirector?redirectedfrom=MSDN
'http://blog.sina.com.cn/s/blog_792da39c01013bzh.html
'https://www.cnblogs.com/lhglihuagang/p/3930874.html
'�����������Ҫ����x86�ĳ������system32�µ�x64�ĳ���,�޷�ֱ�ӷ���
Private Function IsSupport(ByVal strDLL As String, strFunctionName As String) As Boolean
    Dim hMod As Long, lPA As Long
    hMod = LoadLibrary(strDLL)
    If hMod Then
        lPA = GetProcAddress(hMod, strFunctionName)
        FreeLibrary hMod
        If lPA Then
            IsSupport = True
        End If
    End If
End Function
'��Ҫ�������x64��, ��Ϊsystem32�µ�snippΪx64����, �޷�ֱ�ӷ���
'������Ϊ����,���x86 ���� x64������
Sub OpenSnipp() '��system32�µ�snipping��ͼ����
    Dim fsRedirect As Long
    If IsSupport("Kernel32", "Wow64DisableWow64FsRedirection") = True And IsSupport("Kernel32", "Wow64RevertWow64FsRedirection") = True Then
        fsRedirect = Wow64DisableWow64FsRedirection(fsRedirect)
        If fsRedirect Then
            Shell "c:\windows\system32\SnippingTool.exe", vbNormalFocus
            Wow64RevertWow64FsRedirection fsRedirect
            Exit Sub
        End If
        Shell "c:\windows\system32\SnippingTool.exe"
    End If
End Sub
