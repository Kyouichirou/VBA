Attribute VB_Name = "ɾ�������ļ�"
Option Explicit
Private Type uSHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As Long   'unicode�ַ����ĵ�ַ(�����ǰ���unicode)
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long
End Type
Private Type aSHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String  '�ַ���
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long
End Type
'https://docs.microsoft.com/zh-cn/windows/win32/api/shellapi/nf-shellapi-shfileoperationw
'ע�������A��W������, A��ʾ�������ݽ�����ʱ�򽫱�ת��Ϊansi����, w��ʾ����unicode����
'https://blog.csdn.net/Giser_D/article/details/103311433
'WΪΨһ,CreateProcessA����ansi��wide��ת����Ȼ����ײ㺯�����ײ㺯��ֻ��wide�İ汾��
Private Declare Function uSHFileOperation Lib "shell32.dll" Alias "SHFileOperationW" (ByRef lpFileOp As uSHFILEOPSTRUCT) As Long ' ����ɾ������unicode�ַ����ļ�
Private Declare Function aSHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (ByRef lpFileOp As aSHFILEOPSTRUCT) As Long
Private Const FO_Delete = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Const HWND_DESKTOP = 0
Private Const NOCONFIRMATION = &H10 '����ʾ
'varPtr��ָ�������ڴ����ڴ��ַ��������ı���str��û�и�ֵ������ڴ��ַ�ǲ����~
'strPtr��ָ�ַ����ĵ�ַ��Ҳ����Ϊ��ֵ�ĵ�ַ��������ı���str������ʱ�����ռ���0��û���κ�ֵ��ָ������strPtr(str)=0��
'��������str��ֵabc�󣨼�str=��abc������strPtr(str)Ҳ����ָ�������ַ�����Ҳ��������ֵ���ĵ�ַ~
'һ�仰������ȷ������varPtr���ǹ̶�����ģ�����strPtr����������ֵ�ı仯���仯����Ϊֵ���ˣ���ָ��ĵ�ַ�ͱ���~
Function DeleteFiles(ByVal FilePath As String, ByVal isUnicode As Boolean) '֧��ansi, Ҳ֧��ɾ������unicode�ַ����ļ�
    Dim aSHdele As aSHFILEOPSTRUCT
    Dim uSHdele As uSHFILEOPSTRUCT
    If isUnicode = False Then
        With aSHdele
            .hwnd = HWND_DESKTOP
            .pTo = ""
            .wFunc = FO_Delete
            .pFrom = Path + Chr(0)
            .fFlags = FOF_ALLOWUNDO + NOCONFIRMATION
        End With
        aSHFileOperation aSHdele
    Else
        With uSHdele
            .hwnd = HWND_DESKTOP
            .pTo = ""
            .wFunc = FO_Delete
            .pFrom = StrPtr(Path)    '�ؼ�������һ��
            .fFlags = FOF_ALLOWUNDO + NOCONFIRMATION
        End With
        uSHFileOperation uSHdele
    End If
End Function

Function DeleFilepShell(ByVal strx As String) '����powershell������ɾ���ļ�(֧�ַ�ansi�����ַ�),ɾ���ļ�֧�ֵ�����վ '�򿪵��ļ��������,����ʾ(����ɾ������)
    Dim ws As Object
    Dim filepathx As String
    
    If fso.fileexists(strx) = False Then Exit Function '�ļ�������,ֱ���˳�
    filepathx = "'" & strx & " '"               'powershell�������� " '(������) " ���ű� " "(˫����) " ���Ÿ���
    Set ws = CreateObject("wscript.shell")
    ws.Run ("powershell $testFile=" & filepathx & """;dir | Out-File $testFile;$shell = new-object -comobject 'Shell.Application';$item = $shell.Namespace(0).ParseName( (Resolve-Path $testFile).Path);$item.InvokeVerb('delete')"""), 0
    '0��ʾ���ش���'powershell���ö�������ķ���,ʹ�� ";"(�ֺ�)���� "|"
    Set ws = Nothing
End Function

Function DeleToRecycle(ByVal FilePath As String) As Boolean '�Ƴ��ļ�������վ-���� 'boolean�������ж�ɾ�������Ƿ�õ�ʵ�ʵ�ִ��(���ܴ����ļ�����������ռ�ö������޷�ɾ��������)
    Dim objReg As Object
    Dim objShell As Object
    Dim vStateArr As Variant, vBackupState As Variant
    '---------------------------https://docs.microsoft.com/en-us/previous-versions/tn-archive/ee176985(v=technet.10)?redirectedfrom=MSDN
    On Error GoTo 100
    Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    objReg.GetBinaryValue &H80000001, "Software\Microsoft\Windows\CurrentVersion\Explorer", "ShellState", vStateArr
    vBackupState = vStateArr
    vStateArr(4) = 39
    '-------------------�޸�ע���,���õ�����ʾ
    objReg.SetBinaryValue &H80000001, "Software\Microsoft\Windows\CurrentVersion\Explorer", "ShellState", vStateArr
    Set objShell = CreateObject("Shell.Application")
    objShell.Namespace(0).ParseName(FilePath).InvokeVerb ("delete") 'ɾ��ִ��
    objReg.SetBinaryValue &H80000001, "Software\Microsoft\Windows\CurrentVersion\Explorer", "ShellState", vBackupState
    '-------------------------�ָ�ע���
    DeleToRecycle = True 'ɾ�����ִ��
    Set objReg = Nothing
    Set objShell = Nothing
100
    Exit Function
    If Err.Number <> 0 Then
        Set objReg = Nothing
        Set objShell = Nothing
        Err.Clear
        DeleToRecycle = False
    End If
End Function

Sub EmptyRecycle() '��ջ���վ
    Dim retVal As Long
    retVal = SHEmptyRecycleBin(0&, vbNullString, SHERB_NORMAL)
End Sub

Sub SendFile2Recycle(ByVal FilePath As String) '֧��unicode�ַ�,ȱ��,����ȷ��ɾ������, ��Ҫ���ע���ȡ������
    '-----------------------------https://docs.microsoft.com/zh-cn/windows/win32/shell/invokeverbex
    Dim strFolderParent As Variant
    Dim strFileName As Variant
    Dim objShell As Object
    Dim objFolder As Object
    Dim objFolderItem As Object
    
    strFolderParent = fso.GetParentFolderName(FilePath)
    strFileName = fso.GetFileName(FilePath)
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.Namespace(strFolderParent)
    Set objFolderItem = objFolder.ParseName(strFileName)
    objFolderItem.InvokeVerbEx ("Delete")
    Set objShell = Nothing
    Set objFolder = Nothing
    Set objFolderItem = Nothing
End Sub
