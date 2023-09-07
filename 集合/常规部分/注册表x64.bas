Attribute VB_Name = "ע���x64"
Option Explicit
'1.��д��HKEY_LOCAL_MACHINE�ⲿ�ֵ�ʱ��, ��x64�µ�64bit officeҲ�����޷�ֱ��д������, ��Ҫ����ԱȨ�޲���д��
#If VBA7 Then
    Private Declare PtrSafe Function GetSystemWow64Directory Lib "kernel32.dll" Alias "GetSystemWow64DirectoryA" (ByVal lpBuffer As String, ByVal uSize As Long) As Long
#Else
    Private Declare Function GetSystemWow64Directory Lib "kernel32.dll" Alias "GetSystemWow64DirectoryA" (ByVal lpBuffer As String, ByVal uSize As Long) As Long
#End If

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_DYN_DATA = &H80000006

Private Const STANDARD_RIGHTS_READ = &H20000
Private Const KEY_QUERY_VALUE = &H1&
Private Const KEY_ENUMERATE_SUB_KEYS = &H8&
Private Const KEY_NOTIFY = &H10&
Private Const SYNCHRONIZE = &H100000
'Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or _
'                        KEY_QUERY_VALUE Or _
'                        KEY_ENUMERATE_SUB_KEYS Or _
'                        KEY_NOTIFY) And _
'                        (Not SYNCHRONIZE))
Private Const MAXLEN = 256
Private Const ERROR_SUCCESS = &H0&

Const REG_NONE = 0
Const REG_SZ = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY = 3
Const REG_DWORD = 4
Const REG_DWORD_LITTLE_ENDIAN = 4
Const REG_DWORD_BIG_ENDIAN = 5
Const REG_LINK = 6
Const REG_MULTI_SZ = 7
Const REG_RESOURCE_LIST = 8

Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Declare Function apiRegOpenKeyEx Lib "Advapi32.dll" _
        Alias "RegOpenKeyExA" (ByVal hKey As Long, _
        ByVal lpSubKey As String, ByVal ulOptions As Long, _
        ByVal samDesired As Long, ByRef phkResult As Long) _
        As Long

Private Declare Function apiRegCloseKey Lib "Advapi32.dll" _
        Alias "RegCloseKey" (ByVal hKey As Long) As Long

Private Declare Function apiRegQueryValueEx Lib "Advapi32.dll" _
        Alias "RegQueryValueExA" (ByVal hKey As Long, _
        ByVal lpValueName As String, ByVal lpReserved As Long, _
        ByRef lpType As Long, lpData As Any, _
        ByRef lpcbData As Long) As Long

Private Declare Function apiRegQueryInfoKey Lib "Advapi32.dll" _
        Alias "RegQueryInfoKeyA" (ByVal hKey As Long, _
        ByVal lpClass As String, ByRef lpcbClass As Long, _
        ByVal lpReserved As Long, ByRef lpcSubKeys As Long, _
        ByRef lpcbMaxSubKeyLen As Long, _
        ByRef lpcbMaxClassLen As Long, _
        ByRef lpcValues As Long, _
        ByRef lpcbMaxValueNameLen As Long, _
        ByRef lpcbMaxValueLen As Long, _
        ByRef lpcbSecurityDescriptor As Long, _
        ByRef lpftLastWriteTime As FILETIME) As Long
     
Declare Function RegEnumKeyEx Lib "Advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, _
ByVal dwIndex As Long, ByVal lpName As String, lpcName As Long, ByVal lpReserved As Long, _
ByVal lpClass As String, ByVal lpcClass As Long, lpftLastWriteTime As FILETIME) As Long

'Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, _
'lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumKeyExW Lib "Advapi32.dll" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As Long, ByRef lpcName As Long, _
Optional ByVal lpReserved As Long, Optional ByVal lpClass As Long, Optional ByRef lpcClass As Long, Optional ByVal lpftLastWriteTime As Long) As Long
'-----------------------------------------------------------------https://docs.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-internetsetoptiona
Private Declare Function internetsetoption Lib "wininet.dll" Alias "InternetSetOptionA" _
  (ByVal hinternet As Long, ByVal dwoption As Long, ByRef lpBuffer As Any, ByVal dwbufferlength As Long) As Long
  
Function APIReturnRegKeyValue(ByVal lngKeyToGet As Long, ByVal strKeyName As String, ByVal strValueName As String) As String '��һ�����ȡע����µ�ĳ��ֵ
'-------------------------����- Debug.Print APIReturnRegKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\360DrvMgr", "DisplayIcon")
    Dim lnghKey As Long
    Dim strClassName As String
    Dim lngClassLen As Long
    Dim lngReserved As Long
    Dim lngSubKeys As Long
    Dim lngMaxSubKeyLen As Long
    Dim lngMaxClassLen As Long
    Dim lngValues As Long
    Dim lngMaxValueNameLen As Long
    Dim lngMaxValueLen As Long
    Dim lngSecurity As Long
    Dim ftLastWrite As FILETIME
    Dim lngType As Long
    Dim lngData As Long
    Dim lngTmp As Long
    Dim strRet As String
    Dim varRet As Variant
    Dim lngRet As Long
    Const KEY_READ = &H20019, KEY_WOW64_64KEY = &H100&
        On Error GoTo APIReturnRegKeyValue_Err
'    lngTmp = apiRegOpenKeyEx(lngKeyToGet, _
'                strKeyName, 0&, KEY_READ, lnghKey)
    lngTmp = apiRegOpenKeyEx(lngKeyToGet, _
                strKeyName, 0&, KEY_READ Or KEY_WOW64_64KEY, lnghKey) '��ע��� ,KEY_READ Or KEY_WOW64_64KEY �ǹؼ�����, �������������x86�ĳ������x64��ע���
    If Not (lngTmp = ERROR_SUCCESS) Then Err.Raise lngTmp + vbObjectError
    lngReserved = 0&
    strClassName = String$(MAXLEN, 0):  lngClassLen = MAXLEN
    'Get boundary values
    '----------------------------------------------------lnghKey���ؾ�������ں�������ע���Ĳ���
    lngTmp = apiRegQueryInfoKey(lnghKey, strClassName, _
        lngClassLen, lngReserved, lngSubKeys, lngMaxSubKeyLen, _
        lngMaxClassLen, lngValues, lngMaxValueNameLen, _
        lngMaxValueLen, lngSecurity, ftLastWrite)
        
    If Not (lngTmp = ERROR_SUCCESS) Then Err.Raise lngTmp + vbObjectError
    'Now grab the value for the key
    strRet = String$(MAXLEN - 1, 0)
    lngTmp = apiRegQueryValueEx(lnghKey, strValueName, _
                lngReserved, lngType, ByVal strRet, lngData)
    Select Case lngType
      Case REG_SZ  '----ע����ֵ����
        lngTmp = apiRegQueryValueEx(lnghKey, strValueName, _
                lngReserved, lngType, ByVal strRet, lngData)
        varRet = Left(strRet, lngData - 1)
      Case REG_DWORD
        lngTmp = apiRegQueryValueEx(lnghKey, strValueName, _
                lngReserved, lngType, lngRet, lngData)
        varRet = lngRet
      Case REG_BINARY
        lngTmp = apiRegQueryValueEx(lnghKey, strValueName, _
                lngReserved, lngType, ByVal strRet, lngData)
        varRet = Left(strRet, lngData)
    End Select
    
    If Not (lngTmp = ERROR_SUCCESS) Then Err.Raise _
                                lngTmp + vbObjectError
APIReturnRegKeyValue_Exit:
    APIReturnRegKeyValue = varRet
    lngTmp = apiRegCloseKey(lnghKey)
    Exit Function
APIReturnRegKeyValue_Err:
'    varRet = "Error: Key or Value Not Found."
    varRet = ""
    Resume APIReturnRegKeyValue_Exit
End Function

Function ListRegistry64(ByVal subKey As String) As String()
    Const KEY_READ = &H20019, KEY_WOW64_64KEY = &H100&
    Dim lhKey As Long
    Dim i As Long
    Dim sKeyName As String
    Dim lenKeyName As Long, strx As String, strx1 As String, k As Integer, arr() As String
    Dim tFT As FILETIME 'https://docs.microsoft.com/zh-cn/dotnet/api/system.runtime.interopservices.comtypes.filetime?view=netframework-3.0
    Dim n As Long
    
    i = 0 '����
    sKeyName = Space(1024) '��Ԥ�û������ĳ���
    lenKeyName = 1024
'    RegOpenKey HKEY_LOCAL_MACHINE, subKey, lhKey
    apiRegOpenKeyEx HKEY_LOCAL_MACHINE, _
                subKey, 0&, KEY_READ Or KEY_WOW64_64KEY, lhKey
    n = RegEnumKeyEx(lhKey, i, sKeyName, lenKeyName, 0, vbNullString, 0, tFT)
    k = 1
    ReDim arr(1 To 1)
    Do Until n <> 0 '��n��0ʱ����ʾ��������
        '��ȡʵ�ʵļ���
        strx = Left(sKeyName, lenKeyName)
        strx1 = Trim(APIReturnRegKeyValue(HKEY_LOCAL_MACHINE, subKey & strx, "DisplayName"))
        If Len(strx1) > 0 Then arr(k) = strx1: k = k + 1: ReDim Preserve arr(1 To k)
        lenKeyName = 1024  '���û������Ĵ�С
        i = i + 1
        n = RegEnumKeyEx(lhKey, i, sKeyName, lenKeyName, 0, vbNullString, 0, tFT)
    Loop
    apiRegCloseKey lhKey
    If k > 0 Then ReDim ListRegistry(1 To k)
    ListRegistry64 = arr
End Function

Function CheckOS() As Byte '�ж�ϵͳ��x86����x64
    Dim DirPath As String, Result As Byte
    DirPath = Space(255)
    Result = GetSystemWow64Directory(DirPath, 255) 'ͨ���ж�SystemWow64�Ĵ���
    CheckOS = IIf(Result <> 0, 64, 32)
End Function

'-----https://docs.microsoft.com/zh-cn/windows/win32/api/winreg/nf-winreg-regenumkeyexw
'https://docs.microsoft.com/en-us/windows/win32/wmisdk/requesting-wmi-data-on-a-64-bit-platform
'----http://www.office-cn.net/t/api/regenumkeyex.htm#regenumkeyex
'---https://docs.microsoft.com/en-us/windows/win32/winprog64/accessing-an-alternate-registry-view
'----https://docs.microsoft.com/en-us/windows/win32/winprog64/registry-redirector
'----https://docs.microsoft.com/zh-cn/windows/win32/wmisdk/requesting-wmi-data-on-a-64-bit-platform?redirectedfrom=MSDN
'https://docs.microsoft.com/en-us/archive/blogs/alejacma/how-to-read-a-registry-key-and-its-values-vbscript
'----https://stackoverflow.com/questions/12796624/key-wow64-32key-and-key-wow64-64key

' ��ȡ,�޸ĵȶ�ע�������ķ����ܶ�, ���������snippģ���е�x86�Ľ��̷���x64(system32��)���ļ�һ��, ������ֱ�ӷ���
' ע���, ��64λ��ϵͳ��Ҳ���ֳ�x86��x64����
' ��������, wmi, wsh, api
' wmi����vbs, ������x64���������Ͽ���ʵ�ֶ�x64ע���ķ���
' ��򵥵Ļ���powershell
' Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayIcon | Format-Table �CAutoSize
' Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayIcon | Format-Table �CAutoSize

Sub ListInstallPM() '�г�ϵͳ�ϰ�װ�����г���
    Dim strx1 As String, strx2 As String, i As Integer, j As Integer, k As Integer, p As Byte
    Dim arr As Variant, arrx As Variant
    Dim dic As New Dictionary
    'ע�ⳣ�� HKEY_LOCAL_MACHINE��ѡ��, �����Ҫ�鿴����ע���, ��Ҫ�޸��������
    strx1 = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
    p = CheckOS
    If p = 32 Then
        arr = ListRegistry(strx1)
    ElseIf p = 64 Then
        strx2 = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\" 'x86(x64ϵͳ��) ��װ����б�ע���λ��
        arr = ListRegistry86(strx2)
        arrx = ListRegistry64(strx1)
        i = UBound(arr)
        j = UBound(arrx)
        For k = 1 To i
            dic(arr(k)) = "" 'ȥ��
        Next
        For k = 1 To j
            dic(arrx(k)) = ""
        Next
    End If
    '--------�������dic��key���ǰ�װ����б�
End Sub

Function ListRegistry86(ByVal strKeyPath As String) As String() '��װ�ļ��б�-x86
    Dim arr() As String
    Dim strComputer As String
    Dim objReg As Object, strSubkey As Variant, arrSubkeys As Variant
    Dim Folderpath As String, Name As String, i As Integer
    
    strComputer = "."
    On Error Resume Next
    Set objReg = GetObject( _
    "winmgmts:{impersonationLevel=impersonate}!\\" & _
    strComputer & "\root\default:StdRegProv")
    
    objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubkeys
    
    If IsArray(arrSubkeys) = False Then MsgBox "Ϊ��ȡ���б���Ϣ", vbOKOnly, "Tips": Exit Function
    i = UBound(arrSubkeys)
    i = i + 1
    ReDim arr(1 To i)
    ReDim ListRegistry(1 To i)
    i = 0
    For Each strSubkey In arrSubkeys
        objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath & strSubkey, "DisplayName", Name
        If Len(Name) > 0 Then
               i = i + 1
               arr(i) = Trim(Name)
    '           objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath & strSubkey, "InstallLocation", folderpath '��һ����ȡע����������
        End If
    Next
    ListRegistry = arr
    Set objReg = Nothing
End Function

Function IEVersion() As Byte '��ȡIE������İ汾
    Dim strx As String
    Dim wsh As Object, regstr As String
    
    On Error GoTo 100
    Set wsh = CreateObject("Wscript.Shell")
'    strx = Split(ThisWorkbook.Application.OperatingSystem, " ")(0)
'    If LCase(strx) <> "windows" Then GoTo 100            '�����ֵ������x86 ����x64�����Զ�ȡ
    regstr = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\svcUpdateVersion" '��ȡע���,win7/win8.1 ,IE��������°汾,win10����������
    strx = wsh.RegRead(regstr)
    If Len(strx) = 0 Or Len(strx) = 1 Then IEVersion = 0: Exit Function
    IEVersion = Int(Left(strx, 2))
    Set wsh = Nothing
    Exit Function
100
    IEVersion = 0: Set wsh = Nothing
End Function

Sub ActivexUnTips() '���Excel����activex�ؼ��Ĳ���ȫ��ʾ, �ܶ�ؼ��ڼ��ص�ʱ�����ֲ���ȫ����ʾ
    Dim wsh As Object
    
    Set wsh = CreateObject("Wscript.Shell")
    wsh.RegWrite "HKCU\Software\Microsoft\VBA\Security\LoadControlsInForms", 1, "REG_DWORD"
    wsh.RegWrite "HKCU\Software\Microsoft\Office\Common\Security\UFIControls", 1, "REG_DWORD"
    Set wsh = Nothing
End Sub

Function ChangeProxy(ByVal ipaddress As String, ByVal cmCode As Byte) '�޸Ĵ��������
    '------------------https://ss64.com/vb/regwrite.html
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer", ipaddress, "REG_SZ"
    If cmCode = 0 Or cmCode = 1 Then wsh.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", cmCode, "REG_DWORD"
    Set wsh = Nothing
    '-----------------------https://docs.microsoft.com/zh-cn/windows/win32/wininet/option-flags
    '-----------------------39��ʾ,����ϵͳע����޸���Ч
    Call internetsetoption(0, 39, 0, 0) '��ʵ�ֲ�����ie�����ʵ�ִ���������Ч
End Function
