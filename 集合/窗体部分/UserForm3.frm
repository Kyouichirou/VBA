VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "���ư�"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   18195
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
'-------------------------------------------https://blog.csdn.net/softman11/article/details/6124345 '�ַ�����
'API Ϊ���ڴ�����С����ť ' _���ű�ʾ����
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
  ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
  ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" ( _
  ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
  ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
  ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetActiveWindow Lib "user32.dll" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
  ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

'Constants
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const GWL_EXSTYLE = (-20)
Private Const HWND_TOP = 0
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_SHOWWINDOW = &H40
Private Const WS_EX_APPWINDOW = &H40000
Private Const GWL_STYLE = (-16)
Private Const WS_MINIMIZEBOX = &H20000
Private Const SWP_FRAMECHANGED = &H20
Private Const WM_SETICON = &H80
Private Const ICON_SMALL = 0&
Private Const ICON_BIG = 1&
'-------------------------------------------------------------------------------------�����С����ť
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const STILL_ALIVE = &H103
Private Const INFINITE = &HFFFF '(�������Ҫע��, �ȴ���ʱ��)
Private ExitCode As Long
Private hProcess As Long
Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long '�ж������Ƿ���ɳ�ʼ��
'-----------------------����shellִ��
Private Const BdUrl As String = "https://www.baidu.com"

Dim NewM As Boolean '���ڿ���textbox�˵�
Dim filepath5 As String, folderoutput As String 'md5/����-����б�
Dim filepathc As String, folderpathc As String '��ѹ�ļ�
Dim filepathset As String, folderpathset As String '����
'-------------------------------------------------------------------------------Ϊ��ֹ��ansi�ַ��ĸ���,��Ҫ��ʱ�洢ֵ��������(����ֱ��ʹ��Textbox,����label�ϵ�ֵ)
Dim arrcompress() As String, flc As Byte, arrfilez() As Double '����-��ѹ�ļ�
Dim imgx As Byte '���ڿ���ͼƬ�ؼ�
Dim pgx As Byte, pgx1 As Byte, pgx2 As Byte, pgx3 As Byte, pgx4 As Byte, pgx5 As Byte, pgx6 As Byte, pgx7 As Byte '��ҳ���л�ʱ���ƿؼ������ݵ�����
'----------------------------------------------------------------
Dim browser1 As Byte, browserkey As String '���ڿ���������ؼ�
Dim arraddfolder() As String
Dim docmx As Integer '�������һ��
Dim arrax() As Variant '���, �ļ���,�ļ���չ��,�ļ�·��,�ļ�λ��
Dim arrbx() As Variant '�򿪴���
Dim arrsx() As Variant '�Ƽ�ָ��/����
Dim arrux() As Variant '��ǩ1/��ǩ2
Dim spyx As Integer '�����洢����б�ֵ
Dim storagex As String
Dim searchx As Byte '�ı�����
'----------------------------------------------����/���ڴ洢��Ҫ����������
Dim voicex As Byte '�жϵ����Ƿ��ѯ����
Dim vbsx As Byte, vbfilex As String '�洢vbs·�� '���������õĲ����洢���ڴ���ȥ
'-------------------------------------------------------------��������
Dim wm As Object '������ʱ��Windows media���ڲ��ű������� ' WindowsMediaPlayer
'-----------------------------------------------------
Dim arrlx() As String 'treeview��ʱ�洢key
Dim arrch() As String 'treeview�洢ѡ���ļ�������Ŀ��key
Dim ich As Byte 'treeview����nodes����������
Dim s As Byte 'treeviewnodes�������ʱֵ
'--------------------------------------treeview���
Dim arrTemp() As Integer, arrtemp2() As String, arrtemp1() As String, arrtemp3() As String '�洢����ѵ��������ͽ��
Dim listnum As Byte '�����б�-�����
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'ʱ�� -����ѵ��
Private Declare Function timeGetTime Lib "winmm.dll" () As Long 'ʱ�� -����ѵ��
Dim Flagpause As Boolean, FlagStop As Boolean, Flagnext As Boolean '����ִ�еı���-����ѵ��
'--------------------------------------------------------����ѵ��
Dim clickx As Byte '����checkbox.click�¼���ִ��
'--------------------------------------------------------------------------------------��С��ico
Private Sub AddIcon() '������ʾͼ��
    Dim hwnd As Long
    Dim lngRet As Long
    Dim hIcon As Long
    'hIcon = Image1.Picture.Handle
    'hWnd = FindWindow(vbNullString, Me.Caption)
    lngRet = SendMessage(hwnd, WM_SETICON, ICON_SMALL, ByVal hIcon)
    'lngRet = SendMessage(hWnd, WM_SETICON, ICON_BIG, ByVal hIcon)
    lngRet = DrawMenuBar(hwnd)
End Sub

Private Sub AddMinimiseButton() '����������С����ť
    Dim hwnd As Long
    
    hwnd = GetActiveWindow
    Call SetWindowLong(hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) Or WS_MINIMIZEBOX)
    Call SetWindowPos(hwnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Private Sub AppTasklist(myForm) '��ӵ�������
    Dim wStyle As Long
    Dim Result As Long
    Dim hwnd As Long
    
    hwnd = FindWindow(vbNullString, myForm.Caption)
    wStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
    wStyle = wStyle Or WS_EX_APPWINDOW
    Result = SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or _
      SWP_HIDEWINDOW)
    Result = SetWindowLong(hwnd, GWL_EXSTYLE, wStyle)
    Result = SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or _
      SWP_SHOWWINDOW)
End Sub
'---------------------------------------------------------------------------------------------------��С��ico

Private Sub CheckBox10_Click() '����-��ֹ��ʾɾ���ļ���ʾ
    If clickx = 1 Then clickx = 0: Exit Sub
    If Me.CheckBox10.Value = True Then
        ThisWorkbook.Sheets("temp").Range("ab37").Value = 1
    Else
        ThisWorkbook.Sheets("temp").Range("ab37").Value = ""
    End If
End Sub

Private Sub CheckBox11_Click() '����-�Զ�����md5
    If clickx = 1 Then clickx = 0: Exit Sub
    With Me
        If .CheckBox11.Value = True Then
            ThisWorkbook.Sheets("temp").Range("ab35") = 1
            .CheckBox12.Enabled = True
        Else
            ThisWorkbook.Sheets("temp").Range("ab35") = ""
            .CheckBox12.Enabled = False
            ThisWorkbook.Sheets("temp").Range("ab36") = ""
        End If
    End With
End Sub

Private Sub CheckBox12_Click() '����-����������
    If clickx = 1 Then clickx = 0: Exit Sub
    If Me.CheckBox12.Value = True Then
        ThisWorkbook.Sheets("temp").Range("ab36") = 1
    Else
        ThisWorkbook.Sheets("temp").Range("ab36") = ""
    End If
End Sub

Private Sub CheckBox13_Click() '�ļ�����
    Dim yesno As Variant, strx As String, strx1 As String, strx2 As String
    
    If Fileptc > 0 Then Fileptc = 0: Exit Sub '��ֹ��checkbox��ֵ�����¼�
    With Me
        strx = .Label29.Caption
        If Len(strx) = 0 Or .Label55.Visible = True Then .CheckBox13.Value = False: Exit Sub
        With ThisWorkbook.Sheets("temp")
            If Len(.Cells(41, "ab").Value) = 0 Then
               strx2 = ThisWorkbook.Path & "\protect"
                .Cells(41, "ab") = strx2
                If fso.folderexists(strx2) = False Then fso.CreateFolder strx2 '�����ļ���
            Else
                strx2 = .Cells(41, "ab").Value
                If fso.folderexists(strx2) = False Then fso.CreateFolder strx2
            End If
            strx1 = strx2 & "\" & Filenamei '.Label23.Caption
        End With
        
        If .CheckBox13.Value = True Then
            If Len(.Label76.Caption) > 52428800 Then
                yesno = MsgBox("�ļ�̫��,ִ�н���,�Ƿ����?", vbYesNo, "Warning")
                If yesno = vbNo Then Exit Sub
            End If
            
            If FileStatus(strx, 2) = 4 Then '����ļ��Ƿ����
                Rng.Offset(0, 32) = 1
            Else
'                .Label55.Visible = ture
'                .TextBox1.Text = ""
                .Label57.Caption = "�ļ�������"
                DeleFileOverx strx
            End If
        Else
            SearchFile strx
            If Rng Is Nothing Then
'                .Label55.Visible = ture
'                .TextBox1.Text = ""
                .Label57.Caption = "�ļ�������"
                DeleFileOverx strx
            End If
            Rng.Offset(0, 32) = ""
            If fso.fileexists(strx1) = True Then fso.DeleteFile (strx1) 'ȡ���ļ�����,���Ƶ��ļ�����ɾ����
        End If
    End With
    Set Rng = Nothing
End Sub

Private Sub CheckBox16_Click() '��ѹ-ɾ��Դ�ļ�
    Dim yesno As Variant
    
    If Me.CheckBox16.Value = True Then
        yesno = MsgBox("ע��: ���ѡ,�����ļ��Ƿ��ѹ�ɹ�,Դ�ļ����ᱻɾ��" & vbCr _
               & "(��:�����ļ��������������)�Ƿ����?", vbYesNo, "Warning!!!")
        If yesno = vbNo Then Me.CheckBox16.Value = False
    End If
End Sub

Private Sub CheckBox17_Click()
    If clickx = 1 Then clickx = 0: Exit Sub
    With Me.CheckBox17
        If .Value = True Then
            ThisWorkbook.Sheets("temp").Cells(43, "ab") = 1
        Else
            ThisWorkbook.Sheets("temp").Cells(43, "ab") = ""
        End If
    End With
End Sub

Private Sub CheckBox18_Click() '���������
    If clickx = 1 Then clickx = 0: Exit Sub
    With ThisWorkbook.Sheets("temp")
        If Len(.Cells(6, "ab")) = 0 Then Me.CheckBox18.Value = False: Exit Sub
        If Me.CheckBox18.Value = ture Then
            .Cells(45, "ab") = 1
        Else
            .Cells(43, "ab") = ""
        End If
    End With
End Sub

Private Sub CheckBox19_Click() '�ļ������ٷ�ʽ
    If clickx = 1 Then clickx = 0: Exit Sub
    With ThisWorkbook.Sheets("temp")
        If Me.CheckBox19.Value = True Then
            .Cells(50, "ab") = 1
        Else
            .Cells(50, "ab") = ""
        End If
    End With
End Sub

Private Sub CheckBox21_Click() '�������ļ���
    If Me.CheckBox21.Value = True Then Me.CheckBox22.Value = False
End Sub

Private Sub CheckBox22_Click() '�����ļ���
    If Me.CheckBox22.Value = True Then Me.CheckBox21.Value = False
End Sub

Private Sub CheckBox23_Click() 'ͬ����ӷ���
    If clickx = 1 Then clickx = 0: Exit Sub
    With ThisWorkbook.Sheets("temp")
        If Me.CheckBox23.Value = True Then
            .Cells(53, "ab") = 1
        Else
            .Cells(53, "ab") = ""
        End If
    End With
End Sub

Private Sub CheckBox25_Click() '�����ִ�Сд
    With Me
    If .CheckBox25.Value = True Then
        .CheckBox26.Enabled = False
    Else
        .CheckBox26.Enabled = True
    End If
    End With
End Sub

Private Sub CheckBox26_Click() '����ģʽ
    With Me
    If .CheckBox26.Value = True Then
        .CheckBox25.Enabled = False
    Else
        .CheckBox25.Enabled = True
    End If
    End With
End Sub

Private Sub CheckBox27_Click() '����ģʽ
    If clickx = 1 Then clickx = 0: Exit Sub
    With ThisWorkbook.Sheets("temp")
        If Me.CheckBox27.Value = True Then
            .Cells(54, "ab") = 1
        Else
            .Cells(54, "ab") = ""
        End If
    End With
End Sub

Private Sub CommandButton106_Click() '�ļ���-��ʾ�����б�
    With Me
        If .CommandButton106.Caption = "��ʾ�����б�" Then
            .CommandButton106.Caption = "���ص����б�"
            .ListBox6.Visible = True
        Else
            .ListBox6.Visible = False
            .CommandButton106.Caption = "��ʾ�����б�"
        End If
    End With
End Sub

Private Sub CommandButton107_Click() '�ļ���-�����б�-Ĭ�ϵ���Ϊtext�ĵ�,����Ŀ¼Ϊdocuments 'https://docs.microsoft.com/zh-TW/office/vba/Language/Reference/User-Interface-Help/createtextfile-method
    Dim fl As Object, strx As String
    Dim i As Integer, k As Integer
    
    With Me.ListBox6
        k = .ListCount
        If k = 0 Then Exit Sub
        strx = Environ("UserProfile") & "\Desktop\" & CStr(Format(Now, "yyyymmddhhmmss")) & ".txt" '����ʱ�䴴��txt�ļ�
        If fso.fileexists(strx) = False Then
            Set fl = fso.CreateTextFile(strx, True)
            k = k - 1
            For i = 0 To k
                fl.WriteLine .List(i, 0) & "  " & .List(i, 1) '���б������д��
            Next
            fl.Close
            Me.Label57.Caption = "�����ɹ�"
        Else
            MsgBox "Ŀ¼���Ѵ�����ͬ���ļ�", vbOKOnly, "Warning"
            Exit Sub
        End If
    End With
    Set fl = Nothing
End Sub

Private Sub CommandButton108_Click() '�ļ���-�����ļ�
    Dim strfolder As String
    Dim i As Integer, k As Integer, strx3 As String, strx4 As String
    Dim strx As String, strx1 As String, fl As Object, strx2 As String, dicsize As Double, filez As Long, dics As String
    
    With Me.ListBox6
        k = .ListCount
        If k = 0 Then Exit Sub
        With ThisWorkbook.Application.FileDialog(msoFileDialogFolderPicker) '�ļ���ѡ�񴰿�(�ļ����Դ�ֻ��ѡһ��)
            .Show
            If .SelectedItems.Count = 0 Then Exit Sub 'δѡ���ļ������˳�sub
            strfolder = .SelectedItems(1)
        End With
        strfolder = strfolder & "\"
        k = k - 1
        Set fl = fso.CreateTextFile(strfolder & CStr(Format(Now, "yyyymmddhhmmss")) & ".txt", True, True) '����txt�ļ�
        dics = Left$(strfolder, 1)
        dicsize = fso.GetDrive(dics).AvailableSpace '���̵Ŀռ��С
        For i = 0 To k
            strx3 = .List(i, 0)
            SearchFile strx3    '�����ļ�
            If Rng Is Nothing Then GoTo 101
            strx = Rng.Offset(0, 3) '�ļ�·��
            strx4 = Rng.Offset(0, 1) '�ļ���
            strx1 = strfolder & strx4 '�µ��ļ�·��
            strx2 = strx4 '�ļ���
            If fso.fileexists(strx) = True And fso.fileexists(strx1) = False Then '�ļ����������ļ�����û���ظ����ļ�
                filez = fso.GetFile(strx).Size                                    '�������ϸ��-ǿ�Ƹ��»��ǵ�����ʾ����(overwrite or promotion)
                dicsize = dicsize - filez
                If dicsize < 209715200 Then MsgBox "���̿ռ䲻��!", vbCritical, "Warning": GoTo 100 '�����̵Ŀռ�С��200M��ʱ��
                strx = """" & strx & """"
                strfolder = """" & strfolder & """"  '------------'ע��cmd�����·������,���ļ���·�����ڿո�
                Shell ("cmd /c" & "copy " & strx & Chr(32) & strfolder), vbHide     'fso.CopyFile (strx), strfolder
                fl.WriteLine strx3 & Space(2) & strx2 & Space(3) & "Success" & Chr(13)   '�����ļ���ͬʱ,�����ļ����б�
            Else
101
                fl.WriteLine strx3 & Space(2) & strx2 & Space(3) & "Fail" & Chr(13) 'space���� https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/space-function
            End If
        Next
    End With
    Sleep 100
    Me.Label57.Caption = "�����ɹ�"
100
    fl.Close
    Set fl = Nothing
End Sub

Private Sub CommandButton109_Click() '�ļ���-����б�
    Me.ListBox6.Clear
End Sub

Private Sub CommandButton110_Click() '�ļ���ӵ������б�
    Dim i As Integer, k As Integer, p As Byte
    Dim strx As String, strx1 As String, strx2 As String
    
    With Me.ListView2
        k = .ListItems.Count
        If k = 0 Then Exit Sub
        For i = 1 To k
            If .ListItems(i).Selected = True Then
                strx = .SelectedItem.Text
                strx1 = .SelectedItem.SubItems(1)
                strx2 = .SelectedItem.SubItems(3)
                If CheckLb6(strx) = False Then
                    With Me.ListBox6
                        If .ListCount > 30 Then MsgBox "��������ﵽ����": Exit Sub
                        .AddItem
                        p = .ListCount - 1
                        .List(p, 0) = strx '���
                        .List(p, 1) = strx1 '�ļ���
'                        .List(p, 2) = strx2 '�ļ�·��
                    End With
                End If
            End If
        Next
    End With
End Sub

Private Sub CommandButton111_Click() '�ļ���-�Ƴ��б�
    Dim i As Integer, k As Integer
    
    With Me.ListBox6
        k = .ListCount
        If k = 0 Or .ListIndex < 1 Then Exit Sub '-1��ʾûѡ��
        k = k - 1
        For i = 0 To k
            If .Selected(i) = True Then .RemoveItem (i)
        Next
    End With
End Sub

Private Sub CommandButton112_Click() '�ļ���-�鿴�ļ�����
    Dim i As Integer, k As Integer, strx As String
    
    With Me.ListView2
        k = .ListItems.Count
        If k = 0 Then Exit Sub
        For i = 1 To k
            '.SelectedItem.Text, ����û��ѡ��, Ҳ�ᱻִ��, Ĭ��Ϊѡ�е�һ��, ע��checked��selected֮�������
            If .ListItems(i).Selected = True Then
                strx = .ListItems(i).Text
                SearchFile strx
                If Rng Is Nothing Then Me.Label57.Caption = "�ļ���ʧ": Exit Sub
                ShowDetail (strx)
                Exit For
            End If
        Next
    End With
    Set Rng = Nothing
End Sub

Private Sub CommandButton113_Click() '������-��С
    If Workbooks.Count = 1 Then
        ThisWorkbook.Application.Visible = False
    Else
        ThisWorkbook.Windows(1).WindowState = xlMinimized
    End If
End Sub

Private Sub CommandButton114_Click() '��������
    With ThisWorkbook
        If .Application.Visible = False Then .Application.Visible = True
        .Windows(1).WindowState = xlMaximized
        With .Sheets("temp")
            .Activate
            .Range("aa1").Select
        End With
    End With
    Unload Me
End Sub

Private Sub CommandButton115_Click() '����-�ؽ�
    Dim strx As String
    
    strx = ThisWorkbook.Path & "\lbrecord.xlsx"
    If fso.fileexists(strx) = True Then Me.Label57.Caption = "�ļ��Ѵ���": Exit Sub
    ThisWorkbook.Application.ScreenUpdating = False
    Call CreateWorksheet(strx)
    ThisWorkbook.Application.ScreenUpdating = True
    Me.Label57.Caption = "�����ɹ�"
End Sub

Private Sub CommandButton116_Click() '����-Ԥ��
    Dim strfolder As String
    Dim strx As String, strx1 As String
    Dim wsh As Object, fd As Folder
    Dim pid As Long, i As Integer
    
    With Application.FileDialog(msoFileDialogFolderPicker) '�ļ���ѡ�񴰿�(�ļ����Դ�ֻ��ѡһ��)
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub 'δѡ���ļ������˳�sub
        strfolder = .SelectedItems(1)
    End With
    Set fd = fso.GetFolder(strfolder)
    If fd.Files.Count = 0 And fd.SubFolders.Count = 0 Then Set fd = Nothing: Me.Label57.Caption = "���ļ���": Exit Sub
    strx = Split(strfolder, "\")(UBound(Split(strfolder, "\"))) '��ȡ�ļ��е�����
'    Set wsh = CreateObject("WScript.Shell")
    strx1 = Environ("UserProfile") & "\Desktop\" & strx & ".txt"
    
    strx1 = """" & strx1 & """"
    strfolder = """" & strfolder & """"
    '��ֹ���ֿո�ȸ���
'    wsh.Run "cmd /c tree " & strfolder & " /f >>" & strx1, 0 '����cmd��tree���� 'https://wenku.baidu.com/view/66979c4fcf84b9d528ea7a96.html,cmd����,0:hidden,3: max & activate
    
    pid = Shell("cmd /c tree " & strfolder & " /f >" & strx1, 0)
    
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0, pid)
'    Sleep 500
    Do
        GetExitCodeProcess hProcess, ExitCode
        DoEvents
        Sleep 50
        i = i + 1
    Loop While ExitCode = STILL_ALIVE And i < 75 '����ִ�е�ʱ��, �ȴ�ʱ�䲻�ܳ���75*50/1000 s
    If ExitCode = STILL_ALIVE Then Me.Label57.Caption = "���������ļ�����,�ļ�����̫�����Ժ�....": Exit Sub
    CloseHandle hProcess
    
    Shell "notepad.exe " & strx1, 3 '�����ɵ��ļ�
'    Set wsh = Nothing
    Set fd = Nothing
End Sub

Private Sub CommandButton117_Click() '����-�ļ��ض��ļ��б�
    Dim strx As String, strx1 As String, strlen1 As Byte, strx2 As String, strx3 As String, strx4 As String, strx5 As String, strx6 As String
    Dim strx7 As String, strx8 As String
    Dim wsh As Object, i As Byte, xi As Variant, k As Byte, fd As Folder, fdz As Double, j As Single, wt As Integer, p As Integer, n As Single, pidx As Long

    With Me
        strx = Trim(folderoutput)
        strx1 = Trim(.TextBox27.Text)
        strlen1 = Len(strx1)
        If Len(strx) = 0 And strlen1 = 0 Then Exit Sub
        If InStr(strx, Chr(92)) = 0 Then Exit Sub          'chr(92),\,chr32,�ո�
        Set fd = fso.GetFolder(folderoutput)
'        fdz = fd.Size
        If fso.folderexists(folderoutput) = False Then Exit Sub
        If fd.Files.Count = 0 And fd.SubFolders.Count = 0 Then Exit Sub
        i = InStr(strx1, Chr(32))
        strx5 = fd.Name
        strx7 = Environ("UserProfile") & "\Desktop\" & strx5 & ".txt" '�������
        If fso.fileexists(strx7) = True Then .Label57.Caption = "�ļ��Ѵ���": GoTo 100
        If strlen1 >= 6 And i = 0 Then .TextBox27.SetFocus: .Label57.Caption = "��ʽ����": Exit Sub '��������Ƿ���ȷ
        strx3 = "*."
        If strlen1 > 0 Then
            If i = 0 Then
                strx4 = strx3 & strx1
            Else
                xi = Split(strx1, Chr(32))
                i = UBound(xi)
                For k = 0 To i
                    strx4 = strx4 & strx3 & xi(k) & Chr(44) 'chr44,����
                Next
            End If
        Else
            strx4 = "*.*" '�����������չ��,��Ĭ��������չ��
        End If
        strx4 = Chr(40) & strx4 & Chr(41) 'chr40/41����
        strx6 = """" & strx & """"
        strx7 = """" & strx7 & """"
        strx8 = "for /r " & strx6 & " %a in " & strx4 & " do >>" & strx7 & " echo %~dpa%~nxa"  'for /r cmd����
        Set wsh = CreateObject("WScript.Shell")
        wsh.Run ("cmd /c " & strx8), vbHide, True 'CreateObject("WScript.Shell"), ֧��ͬ��ִ��, true����ʾͬ��ִ��, ������Ҫ���ǵ�ִ�е�ʱ���Ƿ����������
'        pidx = Shell("cmd /c " & strx8, vbHide)
        
'        If CheckProgramRun(pidx) = True Then .Label57.Caption = "Ŀ¼������,�Ժ����ֶ���": Exit Sub
    
'        If fdz < 1073741824 Then                    '�ȴ��ļ���ȫ���ɵ�ʱ��(����ȡ���ֶ�������ʱ��)
'            wt = 300
'        ElseIf fdz > 1073741824 And fdz < 10737418240# Then '1G
'            j = Int(fdz / 1073741824)
'            n = CSng(j)                'ת������Ϊ������
'            p = (n + n / 10 - 0.1 - n / 50) * 300
'            wt = Int(Round(fdz / 1073741824 / j, 3) * p)
'        ElseIf fdz >= 10737418240# Then '10G
'            .Label57.Caption = "������..."
'            j = Int(fdz / 10737418240#)
'            n = CSng(j)
'            p = (n + n / 10 - 0.1 - n / 50) * 500 '���ݲ�ͬ�豸�����ܽ���΢��
'            wt = Int(Round(fdz / 10737418240# / j, 3) * p)
'        End If
'
'        Sleep wt '��ʱ
100
        Shell "notepad.exe " & strx7, 3 '���ļ�
        Set fd = Nothing
        Set wsh = Nothing
        folderoutput = ""
        .Label57.Caption = "�������"
    End With
End Sub

Private Sub CommandButton119_Click() '����-CMD
    Shell ("cmd "), vbNormalFocus
End Sub

Private Sub CommandButton120_Click() '����-�������
    Dim i As Byte, k As Byte

    With Me
        If .OptionButton7.Value = True Then
        i = 0
        ElseIf .OptionButton8.Value = True Then
        i = 1
        ElseIf .OptionButton9.Value = True Then
        i = 2
        End If
        If IsNumeric(.ComboBox13.Text) = True Then
            k = Abs(Int(.ComboBox13.Text))
        Else
            Exit Sub
        End If
        .TextBox28.Text = PasswordGR(i, k)
        If i = 0 Then .Label57.Caption = "���������밲ȫ�Ե�"
    End With
End Sub

Private Sub CommandButton121_Click() '����-�ر�����������
    Dim i As Byte, k As Byte
    
    i = Workbooks.Count
    If i = 1 Then Exit Sub
    For k = 1 To i
        If Workbooks(k).Name <> ThisWorkbook.Name Then Workbooks(k).Close savechanges:=True
    Next
End Sub

Private Sub CommandButton122_Click() '�����ļ�
    If Len(Me.Label29.Caption) = 0 Then Exit Sub
    UserForm13.Show
End Sub

Private Sub CommandButton123_Click() '��ѹ�ļ�
    Dim strx As String, fd As Folder, fl As File, i As Byte, k As Byte, j As Byte, m As Byte, n As Byte
    Dim discz As Double, yesno As Variant
    Dim strx1 As String, strx2 As String, filez As Double, t As Long
    
    On Error GoTo 102
    With Me
        strx = Trim(.TextBox29.Text)
        m = Len(strx)
        If m = 0 Then Exit Sub
        strx1 = ThisWorkbook.Sheets("temp").Cells(42, "ab").Value
        If Len(strx1) = 0 Then MsgBox "��δ���ý�ѹ�ļ����λ��": Exit Sub
        If fso.folderexists(strx1) = False Then MsgBox "���ý�ѹ�ļ�������": Exit Sub
        strx2 = Left$(strx1, 1)
        discz = fso.GetDrive(strx2).AvailableSpace '��ȡ���̵Ĵ�С
        If .CheckBox16.Value = True Then k = 1
        
        If InStr(strx, ".") > 0 Then '��ʾ�ļ�
            strx = filepathc
            If m <> Len(filepathc) Then
            If fso.fileexists(strx) = False Then .Label57.Caption = "�ļ�������": Exit Sub
            If CheckFileFrom(strx, 1) = True Then MsgBox "�ļ���Դ����": Exit Sub
            If TerminateEXE("bc.exe", 0) = 1 Then '���
                yesno = MsgBox("�Ƿ�ر�����ʹ�õ�Bandzip", vbYesNo, "Tips")
                If yesno = vbYes Then
                    TerminateEXE "bc.exe", 1
                    .Label57.Caption = "ִ�н�����,����ʹ��Bandzip"
                Else
                    Exit Sub
                End If
            End If
            filez = fso.GetFile(strx).Size
            filez = filez * 5
            If filez > discz Then MsgBox "���̿ռ䲻��": Exit Sub
            ZipExtract strx, strx1
            If k = 1 Then
                t = timeGetTime
                Do
                    DoEvents
                    Sleep 200
                    i = TerminateEXE("bc.exe", 0)
                    If i = 0 Then Shell ("cmd /c" & "del /s " & strx), vbHide: .Label57.Caption = "�������": Exit Sub '��ִ�����֮��ִ��ɾ������
                Loop Until i = 0 Or timeGetTime - t > 60000
                If i = 1 Then TerminateEXE "bc.exe", 1: .Label57.Caption = "���������쳣": Exit Sub '����޷���60�����˳�����,���Զ����������˳�
            End If
            filepathc = ""
        Else                                '�ļ���
            strx = folderpathc
            If fso.folderexists(strx) = False Then .Label57.Caption = "�ļ��в�����": Exit Sub
            If CheckFileFrom(strx, 2) = True Then MsgBox "�ļ���Դ����": Exit Sub
            ReDim arrcompress(1 To 100)
            ReDim arrfilez(1 To 100)
            flc = 0
            Set fd = fso.GetFolder(strx)
            If .CheckBox15.Value = False Then j = 1
            
            FoldersCompFile fd, j
            
            If flc = 0 Then Exit Sub
            .Label57.Caption = "����ִ����...���������������"
            For i = 1 To flc
                filez = arrfilez(i)
                filez = filez * 5
                discz = discz - filez
                If discz < 209715200 Then MsgBox "���̿ռ䲻��": GoTo 100
                ZipExtract arrcompress(i), strx1
            Next
            
            If k = 1 Then
            t = timeGetTime
            Do
                DoEvents
                Sleep 200
                i = TerminateEXE("bc.exe", 0)
                If i = 0 Then GoTo 101
            Loop Until i = 0 Or timeGetTime - t > 300000 '�����ʱ��������5����
101
            If i = 1 Then TerminateEXE "bc.exe", 1: GoTo 102 '����޷���60�����˳�����,���Զ����������˳�
            For i = 1 To flc
                strx = arrcompress(i)
                strx = """" & strx & """" '��ֹ���ڿո�ȸ�������
                Shell ("cmd /c" & "del /s " & strx), vbHide
            Next
            End If
100
            Erase arrcompress
            Erase arrfilez
            folderpathc = ""
        End If
        Sleep 100
        .Label57.Caption = "�������"
    End With
Exit Sub
102
Me.Label57.Caption = "���������쳣"
Erase arrcompress
Erase arrfilez
Err.Clear
End Sub

Function FoldersCompFile(ByVal fd As Folder, Optional ByVal cmCode As Byte) '��һ��ȡѹ���ļ�
    Dim sfd As Folder, fl As File, filex As String, strx As String
    
    For Each fl In fd.Files
        If flc > 100 Then Exit Function
        strx = fl.Name
        If InStr(strx, ".") = 0 Then GoTo 100
        filex = LCase(Right$(strx, Len(strx) - InStrRev(strx, "."))) '�ļ���չ��
        If filex Like "rar" Or filex Like "zip" Or filex Like "7z" Then '�޶�������ѹ���ļ�
            flc = flc + 1
            arrcompress(flc) = fl.Path
            arrfilez(flc) = fl.Size
        End If
100
    Next
    If fd.SubFolders.Count = 0 Or cmCode = 1 Then Exit Function
    For Each sfd In fd.SubFolders
        FoldersCompFile sfd, cmCode
    Next
End Function

Private Sub CommandButton124_Click() 'תpdf
    Dim strx As String, i As Byte
    
    With Me
        strx = LCase(.Label24.Caption)
        If Len(strx) = 0 Then Exit Sub
        If strx Like "doc" Or strx Like "docx" Then
            If fso.fileexists(Filepathi) = True Then
                If Len(ThisWorkbook.Sheets("temp").Cells(43, "ab")) > 0 Then i = 1
                WordToPDF Filepathi, i
                .Label57.Caption = "�����ɹ�"
            End If
        Else
            .Label57.Caption = "�ļ����Ͳ�ƥ��,����Word"
        End If
    End With
End Sub

Private Sub CommandButton125_Click() '�༭-ժҪ-�Ű�
    Dim strx As String, xi As Variant, i As Byte, k As Byte, strx1 As String
    
    With Me
        If .Label55.Visible = True Then Exit Sub
        strx = .TextBox2.Text
        If Len(strx) = 0 Then Exit Sub
        If InStr(strx, vbCrLf) > 0 Then         'vbCrLf��ͬ��chr(10)���з���chr(13)�س���
            xi = Split(strx, vbCrLf)
            i = UBound(xi)
            For k = 0 To i
                strx1 = strx1 & k + 1 & ". " & xi(k) & vbCrLf
            Next
            .TextBox2.Text = strx1
        Else
            .TextBox2.Text = "1. " & strx
        End If
    End With
End Sub

Private Sub CommandButton126_Click() '�༭-��������-��ά��
    Dim strx As String

    With Me
        QRtextEN = .Label106.Caption
        If Len(QRtextEN) = 0 Then Exit Sub
        UserForm18.Show
        Exit Sub
        '----------------���沿�ֱ���
'        If Len(.Label106.Caption) = 0 Then Exit Sub
'        SearchFile .Label29.Caption
'        If rng Is Nothing Then .Label57.Caption = "�ļ�������": Set rng = Nothing: Exit Sub
'        strx = rng.Offset(0, 33).Value
'        If Len(strx) = 0 Or fso.FileExists(strx) = False Then
'            If IsNetConnectOnline = False Then .Label57.Caption = "�����޷�����": Exit Sub
'            UserForm14.Show '��ʾ����ͼƬ
'        Else
'            QRfilepath = strx
'            UserForm16.Show '��ʾ����ͼƬ
'        End If
    End With
    Set Rng = Nothing
End Sub

Private Sub CommandButton127_Click() '����-����-���IE����
    Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 1 "
    Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2 "
    Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8 "
    Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 16 "
    Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 32 "
    Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 255 "
    Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 4351 "
End Sub
'---------------------------------------------------�����
Private Sub CommandButton128_Click() '�������ҳ
    Dim strx As String
    
    strx = "https://www.baidu.com/"
    If browser1 = 1 Then
        With Me!web
            .Silent = True
            .Navigate (strx)
        End With
    ElseIf browser1 = 0 Then
        CreateWebBrowser (strx)
    End If
End Sub

Private Sub CommandButton143_Click() '�ⲿ���Ӵ�
    Dim strx As String
    With Me!web
        strx = .LocationURL
    End With
    If Len(strx) = 0 Then Exit Sub
    If Len(ThisWorkbook.Sheets("temp").Cells(45, "ab").Value) > 0 Then
        Turlx = strx
        UserForm15.Show
        Exit Sub
    End If
    Webbrowser strx
End Sub

Private Sub CommandButton144_Click() '��������
    Dim strx As String, sengine As String, Urlx As String
    
    strx = Me.TextBox3.Text
    If Len(strx) = 0 Then Exit Sub
    sengine = "https://book.douban.com/subject_search?cat=1003&search_text="
    Urlx = sengine & Replace(strx, " ", "+")  'douban
    With Me!web
        .Silent = True
        .Navigate (Urlx)
    End With
End Sub

Private Sub CommandButton131_Click() '�����-��ɽ
    Dim strx As String
    
    strx = "http://www.iciba.com/"
    If browser1 = 1 Then
        With Me!web
            .Silent = True
            .Navigate (strx)
        End With
    ElseIf browser1 = 0 Then
        CreateWebBrowser (strx)
    End If
End Sub

Private Sub CommandButton130_Click() '�����-����
    Dim strx As String
    
    strx = "https://book.douban.com/"
    If browser1 = 1 Then
        With Me!web
            .Silent = True
            .Navigate (strx)
        End With
    ElseIf browser1 = 0 Then
        CreateWebBrowser (strx)
    End If
End Sub

Function CreateWebBrowser(ByVal Urlx As String) '��������� '��Ҫע����multipage������webbrowser��ҳ���л������ʧ,����ʹ�����´����ؼ��ķ�ʽ��ԭ���� 'ҳ���ϵĴ���ֻ����������,ʵ���ϲ���������
    On Error Resume Next
    Me.Controls.Add "shell.explorer.2", "Web", True
    If Len(Urlx) < 5 Then Urlx = "https://www.baidu.com"
    With Me!web
        .Top = 78
        .Left = 98
        .Height = 286
        .Width = 780
        .Navigate (Urlx)
        .Silent = True
    End With
    browser1 = 1
End Function

Private Sub CommandButton129_Click() '�����-��ȡ������Ϣ
    Dim strx As String, strx1 As String, yesno As Variant, strx2 As String, strx3 As String, strx4 As String, strx5 As String
    Dim arr() As String, strx6 As String, strx7 As String
    
    If Me.Label55.Visible = True Then Exit Sub
    strx = Me.Label29.Caption
    If Len(strx) = 0 Then Exit Sub
    If browser1 = 0 Then Exit Sub '���������ؼ���δ����
    
    With Me!web
        strx1 = .LocationURL
    End With
    
    If InStr(strx1, "https://book.douban.com/subject/") = 0 Then Exit Sub '���ڶ��������
    SearchFile strx1
    If Rng Is Nothing Then Me.Label57.Caption = "�ļ�������": Exit Sub
    If Len(Rng.Offset(0, 25)) > 0 Then
        yesno = MsgBox("�Ƿ��滻�����е�����", vbYesNo, "��ʾ")
        If yesno = vbNo Then Set Rng = Nothing: Exit Sub
    End If
    Me!web.Stop
    With Me!web.Document
        strx2 = .getElementById("interest_sectl").InnerHtml '������ֲ��ֵ�Դ��
        strx3 = .getElementById("info").InnerHtml '����,
        strx4 = .getElementById("mainpic").InnerHtml '����+����
    End With
    '------------------------------------------------https://www.w3school.com.cn/jsref/met_doc_getelementbyid.asp
    ReDim arr(1 To 5)
    arr = DoubanTreat(strx2, strx3, strx4)
    With Rng '������д����
        .Offset(0, 23) = arr(3) '����
        .Offset(0, 24) = arr(1) '����
        .Offset(0, 25) = strx           '����
        .Offset(0, 14) = arr(2) '����
        strx5 = arr(4) '��������'CheckRname
        xi = Split(strx5, "/")
        strx5 = xi(UBound(xi)) '�ļ���
        strx5 = Right$(strx5, Len(strx5) - InStrRev(strx5, ".") + 1)
        strx5 = strx1 & strx5
        strx6 = ThisWorkbook.Path & "\" & "bookcover"
        If fso.folderexists(strx6) = False Then fso.CreateFolder strx6
        strx5 = strx6 & "\" & strx5 '����Ĵ洢·��
        strx7 = LCase(Right$(arr(4), 3))
        If strx7 = "jpg" Or strx7 = "png" Then '�ж����ӵ������Ƿ�����Ҫ��
            If DownloadFilex(arr(4), strx5) = True Then .Offset(0, 34) = arr(4) '��������
            .Offset(0, 36) = strx5 '����·��
        End If
        If Len(arr(5)) > 0 Then .Offset(0, 37) = arr(5) '���߹���
    End With
    
    With Me
        .Label106.Caption = strx1
        .TextBox3.Text = arr(3)
        .TextBox4.Text = arr(2)
        .Label69.Caption = arr(1)
    End With
    
    Set Rng = Nothing
End Sub
'---------------------------------------------------�����

Private Sub CommandButton132_Click() '�Ƚ�-�ı��Ƚ�
    Dim i As Byte

    With Me
    If .CommandButton132.Caption = "�ı��Ƚ�" Then
            With .TextBox30
                .Visible = True
                .Width = 254
                .Height = 230
                .Left = 10
                .Top = 12
            End With
            With .TextBox31
                .Visible = True
                .Width = 254
                .Height = 230
                .Left = 10
                .Top = 12
            End With
            For i = 204 To 228
                .Controls("label" & i).Visible = False
            Next
            For i = 98 To 101
                .Controls("commandbutton" & i).Visible = False
            Next
            .CommandButton104.Enabled = False
            .CommandButton132.Caption = "�ļ��Ƚ�"
        Else
            With .TextBox30
                .Visible = False
            End With
            With .TextBox31
                .Visible = False
            End With
            For i = 204 To 228
                .Controls("label" & i).Visible = True
            Next
            .CommandButton104.Enabled = True
            .CommandButton132.Caption = "�ı��Ƚ�"
            For i = 98 To 101
                .Controls("commandbutton" & i).Visible = True
            Next
        End If
    End With
End Sub

Private Sub CommandButton133_Click() '����-�ֵ�
    UserForm17.Show
End Sub

Private Sub CommandButton134_Click() '�༭-ժҪ-��ά��
    QRtextCN = Me.TextBox2.Text
    If Len(QRtextCN) = 0 Then Exit Sub
    UserForm1.Show
End Sub

Private Sub CommandButton135_Click() '�༭-����-������
    Barcodex = Me.Label29.Caption
    If Len(Barcodex) = 0 Or Me.Label55.Visible = True Then Exit Sub
    UserForm19.Show
End Sub

Private Sub CommandButton136_Click() '�༭-�ļ���-��ά��
    Dim strx As String
    
    If Me.Label55.Visible = True Then Exit Sub
    If Len(Filenamei) = 0 Then Exit Sub
    strx = Left$(Filenamei, InStrRev(Filenamei, ".") - 1) '�������ļ���չ�����ļ���
    If Filenamei Like "*[һ-��]*" Then
        QRtextCN = strx
        UserForm1.Show
    Else
        QRtextEN = strx
        UserForm18.Show
    End If
End Sub

Private Sub CommandButton138_Click() '�༭��ӵ�ftp
    If Len(Me.Label29.Caption) = 0 Then Exit Sub
    If Me.Label55.Visible = True Then Exit Sub
    UserForm20.Show
End Sub

Private Sub CommandButton139_Click() '�༭-�ļ�����
    Dim i As Byte, strx As String, k As Byte
    With Me
        strx = .Label29.Caption
        If Len(strx) = 0 Then Exit Sub
        If .Label55.Visible = True Then Exit Sub
        If .Label236.Caption = "Y" Then i = 1
        If Len(ThisWorkbook.Sheets("temp").Cells(50, "ab").Value) > 0 Then k = 2 '�������
        If FileDestroy(Filepathi, k, i) = True Then
            SearchFile strx
            If Not Rng Is Nothing Then ThisWorkbook.Sheets("���").rows(Rng.Row).Delete Shift:=xlShiftUp 'ɾ����� '�����ٵ��ļ������ᱻ��¼��������
            DeleFileOverx strx
            .Label57.Caption = "�����ɹ�"
            Set Rng = Nothing
        End If
    End With
End Sub

Private Sub CommandButton140_Click() '����-�����ļ��з���
    Dim strfolder As String
    Dim drx As Drive, i As Byte, k As Byte, yesno As Variant
    
    With ThisWorkbook.Application.FileDialog(msoFileDialogFolderPicker) '�ļ���ѡ�񴰿�(�ļ����Դ�ֻ��ѡһ��)
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub 'δѡ���ļ������˳�sub
        strfolder = .SelectedItems(1)
    End With
    For Each drx In fso.Drives
        If drx.DriveType = 2 Then '����ע���ڶ������ж�ʱ,������ִ�к����ɴ���,��,ĳֵ���жϱ��벻�ǿղ����ж�,����ǿվͻ����,��Ȼ��Ҫ��ȷ�����ֵ�Ƿ�Ϊ�ղŽ�����һ�����ж�
            i = i + 1
        End If
    Next
    If i > 1 Then '���������ϵĹ̶�Ӳ��
        If UCase(Left(strfolder, 2)) = Environ("SYSTEMDRIVE") Then MsgBox "��ֹѡ��ϵͳ��", vbOKOnly, "Tips": Exit Sub
        If UBound(Split(strfolder, "\")) > 2 Then MsgBox "����ѡ���Ŀ¼��1��Ŀ¼", vbOKOnly, "Tips": Exit Sub
    Else
        If CheckFileFrom(strfolder, 2) = True Then MsgBox "��ѡλ������", vbOKOnly, "Tips": Exit Sub
    End If
    yesno = MsgBox("Ĭ������,ѡ ""��""��ѡ��Ӣ��", vbYesNo, "Tips")
    If yesno = vbYes Then
        k = 0
    Else
        k = 1
    End If
    If CreateFolder(strfolder, k) = True Then
        Me.Label57.Caption = "�����ɹ�"
    Else
        Me.Label57.Caption = "����ʧ��"
    End If
End Sub

Private Sub CommandButton141_Click() '��ansi���
    Dim fdx As FileDialog, strfolder As String
    Dim selectfile As Variant

    If Me.CheckBox20.Value = True Then
    With ThisWorkbook.Application.FileDialog(msoFileDialogFolderPicker) '�ļ���ѡ�񴰿�(�ļ����Դ�ֻ��ѡһ��)
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub 'δѡ���ļ������˳�sub
        strfolder = .SelectedItems(1)
        If ErrCode(strfolder, 0) > 1 Then
            Me.Label238.Caption = errcodenx
        Else
            Me.Label238.Caption = "Clear"
        End If
        Exit Sub
    End With
    End If
    Set fdx = ThisWorkbook.Application.FileDialog(msoFileDialogFilePicker)
    With fdx
        .AllowMultiSelect = False '������ѡ�����ļ�(ע�ⲻ���ļ���,�ļ���ֻ��ѡһ��)
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub
        selectfile = .SelectedItems(1)
        If ErrCode(selectfile, 0) > 1 Then
            Me.Label238.Caption = errcodenx
        Else
            Me.Label238.Caption = "Clear"
        End If
    End With
    Set fdx = Nothing
End Sub

Private Sub CommandButton145_Click() 'windows������
    Shell "cmd /c" & "C:\Windows\System32\control.exe /name Microsoft.AdministrativeTools", vbHide
End Sub

Private Sub CommandButton146_Click() 'ע���
    Shell "cmd /c" & "regedit", vbHide
End Sub

Private Sub CommandButton147_Click() '�����
    UserForm15.Show
End Sub

Private Sub CommandButton148_Click() '�ļ����Ƶ����а�
    If Len(Filepathi) = 0 Then Exit Sub
    If fso.fileexists(Filepathi) = True Then
        CutOrCopyFiles Filepathi
        Me.Label57.Caption = "�ļ��Ѹ��Ƶ����а�"
    Else
        Me.Label57.Caption = "�ļ�������"
    End If
End Sub

Private Sub CommandButton149_Click() '�������
    UserForm2.Show
End Sub

Private Sub CommandButton150_Click() '���ɲ����ļ�
    TestFileGR
End Sub

Private Sub CommandButton152_Click() '�༭�鿴�����ļ�����
    Dim strx As String
    
    With Me
        strx = .Label24.Caption
        If Len(strx) = 0 Then Exit Sub
        If strx Like "doc*" Or strx = "xlsx" Or strx Like "ppt*" Then
            FileDetail.Show
        Else
            .Label57.Caption = "����֧�ָ�ʽ"
        End If
    End With
End Sub

Private Sub Image1_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '˫���򿪴�ͼƬ
    UserForm24.Show
End Sub

Private Sub Label23_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '�ļ���˫������
    If Me.Label55.Visible = True Then Exit Sub
    If Me.Label74.Caption = "Y" Then SetClipboard Filenamei: Exit Sub
    CopyToClipboard Filenamei, "�ļ����Ѹ���"  'ע�����ﲻ��ֱ�Ӹ��Ʊ�ǩ�ϵ�����, �����ļ���������ansi�ַ�, ��ǩ�ϵ����ݾ����������
End Sub

Private Function CopyToClipboard(ByVal strText As String, Optional ByVal strtips As String) '���Ƶ�ճ����
    Dim textb As Object, strx As String
    
    With Me
        If Len(strText) = 0 Then Exit Function
        Set textb = .Controls.Add("Forms.TextBox.1", "Text1", False) '�Դ�����ʱtextbox�ķ�ʽʵ�ָ�������;��Ҫע��������ַ�����Ȼ���Ա���ܶิ�Ƶĸ�ʽ����
        '-------------------------------------------------------------����Ҳ��һ�����������޷�����,��ansi������...�����������������
        With textb
            .Text = strText
            .SelStart = 0
            .SelLength = Len(.Text)
            .Copy
        End With
       If Len(strtips) > 0 Then .Label57.Caption = strtips
    End With
    Set textb = Nothing
End Function

Private Sub Label25_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.Label55.Visible = True Then Exit Sub
    If Me.Label74.Caption = "Y" Then SetClipboard Filepathi: Exit Sub
    CopyToClipboard Me.Label25.Caption, "�ļ�·���Ѹ���"
End Sub

Private Sub Label26_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.Label55.Visible = True Then Exit Sub
    OpenFileLocation Folderpathi
End Sub

Private Sub Label29_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.Label55.Visible = True Then Exit Sub
    CopyToClipboard Me.Label29.Caption, "����Ѹ���"
End Sub

Private Sub Label71_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.Label55.Visible = True Then Exit Sub
    CopyToClipboard Me.Label71.Caption, "Hash�Ѿ�����"
End Sub

Private Sub OptionButton10_Click() '�ַ���-����-У��
    If Me.OptionButton10.Value = True Then StringCH (1)
End Sub

Sub StringCH(ByVal cmCode As Byte) ''�ַ���-����
    Dim i As Byte
    
    Select Case cmCode
        Case 1: i = 1
        Case 2: i = 2
        Case 3: i = 3
        Case Else: Exit Sub
    End Select
    ThisWorkbook.Sheets("temp").Cells(38, "ab") = i
End Sub

Private Sub OptionButton11_Click() '�ַ���-����
    If Me.OptionButton11.Value = True Then StringCH (2)
End Sub

Private Sub OptionButton12_Click() '�ַ���-����
    If Me.OptionButton12.Value = True Then StringCH (3)
End Sub

Private Sub TextBox14_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single) 'Ϊtextbox����Ҽ��˵�
    On Error Resume Next
    If Button = 2 And Not NewM Then
        On Error Resume Next
        With ThisWorkbook.Application
            .CommandBars("NewMenu").Delete
            .CommandBars.Add "NewMenu", msoBarPopup, False, True
            With .CommandBars("NewMenu")
                .Controls.Add msoControlButton
                .Controls(1).Caption = "����"
                .Controls(1).FaceId = 21
                .Controls(1).OnAction = "Cutx"
                .Controls.Add msoControlButton
                .Controls(2).Caption = "����"
                .Controls(2).FaceId = 19
                .Controls(2).OnAction = "Copyx"
                .Controls.Add msoControlButton
                .Controls(3).Caption = "ճ��"
                .Controls(3).FaceId = 22
                .Controls(3).OnAction = "Pastex"
                .ShowPopup
            End With
        End With
    End If
    NewM = Not NewM
End Sub

Private Sub TextBox14_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With Me.TextBox14 '��������textbox�����������Ϊʵ�ָ�����Ϣ�����а�ļ��;��
        If Len(.Text) = 0 Then Exit Sub
        .Text = .Value
        .SelStart = 0
        .SelLength = Len(.Text)
        .Copy
    End With
End Sub

Private Sub TextBox15_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With Me.TextBox15 '��������textbox�����������Ϊʵ�ָ�����Ϣ�����а�ļ��;��
        If Len(.Text) = 0 Then Exit Sub
        .Text = .Value
        .SelStart = 0
        .SelLength = Len(.Text)
        .Copy
    End With
End Sub

Private Sub TextBox16_Change()
    With Me.TextBox16 '��������textbox�����������Ϊʵ�ָ�����Ϣ�����а�ļ��;��
        If Len(.Text) = 0 Then Exit Sub
        .Text = .Value
        .SelStart = 0
        .SelLength = Len(.Text)
        .Copy
    End With
End Sub

Private Sub TextBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '����/�༭-�ļ�ժҪ
    With Me.TextBox2 '��������textbox�����������Ϊʵ�ָ�����Ϣ�����а�ļ��;��
        If Len(.Text) = 0 Then Exit Sub
        .Text = .Value
        .SelStart = 0
        .SelLength = Len(.Text)
        .Copy
    End With
End Sub

Private Sub TextBox20_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '����-md5����
    Dim fdx As FileDialog
    Dim selectfile As Variant
    
    Set fdx = ThisWorkbook.Application.FileDialog(msoFileDialogFilePicker)
    With fdx
        .AllowMultiSelect = False '����ѡ�����ļ�(ע�ⲻ���ļ���,�ļ���ֻ��ѡһ��)
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub
        filepath5 = .SelectedItems(1)
        If ErrCode(filepath5, 1) > 1 Then MsgShow "�ļ�·��������ansi����,�����ֶ��޸����ݿ����Ϣ", "Tips", 1800
        Me.TextBox20.Text = filepath5
        If CheckFileFrom(filepath5, 1) = True Then Me.Label57.Caption = "�ļ���Դ����": Exit Sub
        Me.TextBox21.SetFocus
    End With
    Set fdx = Nothing
End Sub

Private Sub TextBox26_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '����-����ض��ļ�
    Dim strfolder As String
    
    With Application.FileDialog(msoFileDialogFolderPicker) '�ļ���ѡ�񴰿�(�ļ����Դ�ֻ��ѡһ��)
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub 'δѡ���ļ������˳�sub
        strfolder = .SelectedItems(1)
    End With
    folderoutput = strfolder
    If ErrCode(folderoutput, 1) > 1 Then MsgShow "�ļ�·��������ansi����,�����ֶ��޸����ݿ����Ϣ", "Tips", 1800
    Me.TextBox26.Text = folderoutput
    Me.TextBox27.SetFocus
End Sub

Private Sub CommandButton67_Click() 'ɸѡ-���ɸѡ����
    With Me
        .ComboBox1.Value = ""
        .ComboBox7.Value = ""
        .ComboBox8.Value = ""
    End With
End Sub

Private Sub ComboBox1_Click() 'ɸѡ
    With Me
        If .CheckBox3.Value = False Then
            .ComboBox7.Value = ""
            .ComboBox8.Value = ""
        End If
    End With
End Sub

Private Sub ComboBox7_Click() 'ɸѡ
    With Me
        If .CheckBox3.Value = False Then
            .ComboBox1.Value = ""
            .ComboBox8.Value = ""
        End If
    End With
End Sub

Private Sub ComboBox8_click() 'ɸѡ
    With Me
        If .CheckBox3.Value = False Then
            .ComboBox7.Value = ""
            .ComboBox1.Value = ""
        End If
    End With
End Sub

Private Sub CommandButton91_Click() '����-�ļ�����
    Dim i As Byte
    If filepath5 = ThisWorkbook.fullname Then MsgBox "�Ƿ�����", vbCritical, "Warning": Exit Sub
    With Me
    If Len(Trim(.TextBox20.Text)) = 0 Then Exit Sub
    If .OptionButton6.Value = True Then
    i = 2
    ElseIf .OptionButton5.Value = True Then
    i = 1
    ElseIf .OptionButton5.Value = False And .OptionButton6.Value = False Then
    i = 0
    End If
    FileDestroy filepath5, i
    filepath5 = "" 'ʹ�������ò���
    End With
End Sub

Private Function FileDestroy(ByVal strx As String, ByVal cmCode As Byte, Optional ByVal cmfrom As Byte) As Boolean '�ļ�����
    Dim fl As File
    Dim flop As Object
    Dim i As Long, k As Byte, j As Double, yesno As Variant, errx As Integer, strx1 As String, p As Integer, m As Double, n As Byte, c As Integer
    
    On Error GoTo 100
    FileDestroy = False
    With Me
        strx = Trim(strx)
        If Len(strx) = 0 Then Exit Function
        yesno = MsgBox("�ļ�������������,ȷ��?_", vbYesNo, "Warning!!!")
        If yesno = vbNo Then Exit Function
        If fso.fileexists(strx) = False Then .Label57.Caption = "�ļ�������": Exit Function
        If CheckFileFrom(strx, 1) = True Then .Label57.Caption = "���ļ�����Դ����": Exit Function '�����������ϵͳ�̵��ļ�
        
        Set fl = fso.GetFile(strx)
        j = fl.Size
        If j > 536870912 And cmCode = 2 Then
            .Label57.Caption = "�ļ�����,��ȴ�������512M"
            Set fl = Nothing
            Exit Function
        End If
        'Excel���ļ� '�޷�ͨ��������������鿴�ļ��Ƿ��ڴ򿪵�״̬
        If InStr(strx, ".") > 0 Then '��ֹû����չ�����ļ�
            strx1 = LCase(Right$(strx, Len(strx) - InStrRev(strx, "."))) '�ļ���չ��
            If strx1 Like "xl*" Then
                c = Workbooks.Count
                strx1 = fl.Name
                For n = 1 To c
                    If strx1 = Workbooks(n).Name Then .Label57.Caption = "�ļ����ڴ򿪵�״̬": Set fl = Nothing: Exit Function
                Next
                c = -1
            End If
        End If
        If j >= 1048576 And cmCode = 2 Then
            If j < 10485760 Then
                p = 32
            ElseIf j >= 10485760 And j < 52428800 Then
                p = 128
            ElseIf j >= 52428800 And j < 104857600 Then '�������д��������ƿ���������
                p = 512
            Else
                p = 1024
            End If
            j = j + 1048576               '��ȴ���-д���������ȫ����֮ǰ������-д������ݽ���Դԭ�ļ�����1024*1024
            .Label57.Caption = "������..."
        ElseIf cmCode = 1 Then
            j = 10240: p = 1
        ElseIf cmCode = 0 Then
            j = 1024: p = 1
        ElseIf cmCode = 2 Then
            j = 1048576: p = 16
        End If
        If cmfrom = 1 Then GoTo 101
        Set flop = fl.OpenAsTextStream(ForWriting, TristateMixed) 'Ĩ���ļ�����Ϣ
        m = Round((j / p), 0) 'round����,0,������С��
        With flop
            For i = 1 To m
                k = RandNumx(1) '���0/1
                strx1 = String(p, k) '����p����ͬ���ַ� 'https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/string-function
                .Write strx1
            Next
            .Close
        End With
        .Label57.Caption = "�������"
    End With
    fso.DeleteFile (strx) '��������վɾ���ļ�
    Set fl = Nothing
    Set flop = Nothing
    FileDestroy = True
    Exit Function
100
    errx = Err.Number
    If errx = 70 Then
        If c <> -1 Then
101
            If WmiCheckFileOpen(strx) = False Then '����������뱣�����ļ�,���淽���޷����ļ�д������,����powershellǿ��д��
                PowerSHForceW strx, j
                Err.Clear
                Set fl = Nothing
                fso.DeleteFile (strx)
                FileDestroy = True
                Me.Label57.Caption = "�������"
                Set flop = Nothing: Exit Function
            End If
            Me.Label57.Caption = "�ļ����ڴ򿪵�״̬"
        Else
            Me.Label57.Caption = "�ļ��������뱣��"
        End If
    Else
        Me.Label57.Caption = "�쳣,�����ļ�ʧ��"
    End If
    Err.Clear
    Set fl = Nothing
    Set flop = Nothing
End Function

Private Sub CommandButton93_Click() '�ر�����
    Dim i As Byte, k As Byte
    
    With Me.MultiPage1
        k = .Pages.Count
        k = k - 1
        .Pages(k).Visible = False
        k = k - 1
        For i = 0 To k
        .Pages(i).Visible = True
        Next
        .Value = 0 '������ҳ
    End With
End Sub

Private Sub CommandButton75_Click() '����-about me
    UserForm7.Show 1
End Sub

Private Sub CommandButton94_Click() '�༭-����-ˢ��'���»�ȡ�ĵ�����Ϣ '����Ϣд���б����ʾ����
    Dim strx As String
    
    strx = Me.Label29.Caption
    If Len(strx) = 0 Or Me.Label55.Visible = True Then Exit Sub
    Call UpdateFileIn(strx)
End Sub

Function UpdateFileIn(ByVal filecode As String) '�����ļ���Ϣ-�ʹ��ļ����Կ��Ǻϲ�
    Dim fl As File
    Dim fz As Long, filex As String
    With Me
    If FileStatus(filecode, 2) = 4 Then '�ļ�������Ŀ¼���ļ������ڴ���
        Set fl = fso.GetFile(Rng.Offset(0, 3).Value)
        If fl.DateLastModified <> Rng.Offset(0, 6).Value Then '�ļ����޸�ʱ�䷢���仯-ͨ������ļ����޸�ʱ�����ж��ļ��������Ƿ����仯
            Rng.Offset(0, 6) = fl.DateLastModified '�ļ��޸���Ϣ
            fz = fl.Size
            filex = UCase(Rng.Offset(0, 2).Value) '�ļ���չ��
            If fz < 1048576 Then
                Rng.Offset(0, 7) = Format(fz / 1024, "0.00") & "KB" '�ļ���С
            Else
                Rng.Offset(0, 7) = Format(fz / 1048576, "0.00") & "MB"    '���µ���Ϣ��д������
            End If
            Rng.Offset(0, 5) = fz '�ļ���ʼ��С
            If filex Like "EPUB" Or filex Like "MOBI" Or filex Like "PDF" Then Rng.Offset(0, 9) = GetFileHashMD5(fl.Path) '�ļ���hashֵ,���ļ����޸ĺ��ļ���hashֵ�ͷ����ı�
            Call FileChange '��������
        Else
            .Label57.Caption = "��Ϣ��������"
        End If
    Else
        .Label57.Caption = "�ļ���ɾ��"
    End If
    .Label57.Caption = "�������"
    End With
    Set Rng = Nothing
    Set fl = Nothing
End Function

Private Sub CommandButton102_Click() '��ӱȽ�a
    Dim str As String
    
    With Me
        str = .Label29.Caption
        If Len(str) = 0 Or .Label55.Visible = True Then Exit Sub
        If Len(.Label179.Caption) Or Len(.Label153.Caption) > 0 Then
             If str = .Label153.Caption Or str = .Label179.Caption Then .Label57.Caption = "�����": Exit Sub
        End If
        Call FileCompA(str)
    End With
End Sub

Private Sub CommandButton95_Click() '�ļ���-�Ƚ�a
    Dim i As Integer, k As Integer, p As Byte
    
    With Me.ListView2
        k = .ListItems.Count
        If k = 0 Then Exit Sub
        For i = 0 To k
        If .ListItems(i).Checked = True Then p = 1: Set Rng = Nothing: Exit For
        Next
        If p = 0 Then Exit Sub
        Call FileCompA(.SelectedItem.Text)
    End With
End Sub

Function FileCompA(ByVal strx As String) '�Ƚ�a
    With Me
        If strx = .Label179.Caption Then Exit Function
        SearchFile (strx)
        If Rng Is Nothing Then .Label57.Caption = "�ļ�Ŀ¼��ʧ": Exit Function
        .Label179.Caption = Rng.Offset(0, 0)
        .Label180.Caption = Rng.Offset(0, 1)
        .Label181.Caption = Rng.Offset(0, 5)
        .Label182.Caption = Rng.Offset(0, 8)
        .Label183.Caption = Rng.Offset(0, 6)
        .Label184.Caption = Rng.Offset(0, 12)
        .Label185.Caption = Rng.Offset(0, 14)
        .Label186.Caption = Rng.Offset(0, 17)
        .Label187.Caption = Rng.Offset(0, 18)
        .Label188.Caption = Rng.Offset(0, 19)
        .Label189.Caption = Rng.Offset(0, 20)
        .Label190.Caption = Rng.Offset(0, 15)
        .Label191.Caption = Rng.Offset(0, 16)
    End With
    Set Rng = Nothing
End Function

Private Sub CommandButton103_Click() '��ӱȽ�b
    Dim strx As String
    
    With Me
        strx = .Label29.Caption
        If Len(strx) = 0 Or .Label55.Visible = True Then Exit Sub
        If Len(.Label153.Caption) Or Len(.Label179.Caption) > 0 Then
            If strx = .Label153.Caption Or strx = .Label179.Caption Then .Label57.Caption = "�����": Exit Sub
        End If
        Call FileCompB(strx)
    End With
End Sub

Private Sub CommandButton96_Click() '�ļ���-��ӱȽ�b
    Dim i As Integer, k As Integer, p As Byte
    
    With Me.ListView2
        k = .ListItems.Count
        If k = 0 Then Exit Sub
        For i = 0 To k
        If .ListItems(i).Checked = True Then p = 1: Exit For
        Next
        If p = 0 Then Exit Sub
        Call FileCompB(.SelectedItem.Text)
    End With
End Sub

Function FileCompB(ByVal strx As String)
    With Me
        If strx = .Label153.Caption Then Exit Function
        SearchFile (strx)
        If Rng Is Nothing Then .Label57.Caption = "�ļ�Ŀ¼��ʧ": Set Rng = Nothing: Exit Function
        .Label153.Caption = Rng.Offset(0, 0)
        .Label154.Caption = Rng.Offset(0, 1)
        .Label155.Caption = Rng.Offset(0, 5)
        .Label156.Caption = Rng.Offset(0, 8)
        .Label157.Caption = Rng.Offset(0, 6)
        .Label158.Caption = Rng.Offset(0, 12)
        .Label159.Caption = Rng.Offset(0, 14)
        .Label160.Caption = Rng.Offset(0, 17)
        .Label161.Caption = Rng.Offset(0, 18)
        .Label162.Caption = Rng.Offset(0, 19)
        .Label163.Caption = Rng.Offset(0, 20)
        .Label164.Caption = Rng.Offset(0, 15)
        .Label165.Caption = Rng.Offset(0, 16)
    End With
    Set Rng = Nothing
End Function

Private Sub CommandButton104_Click() '�����ļ�����md5�Ƚ�
    Dim strx As String, strx1 As String, strx3 As String, strx4 As String
    Dim strx2 As String, strx5 As String, strx6 As String, strx7 As String
    
    With Me
        strx6 = .Label179.Caption
        strx7 = .Label153.Caption
        If Len(strx6) = 0 Or Len(strx7) = 0 Then Exit Sub
        If FileStatus(strx6, 2) = 4 Then
            strx = Rng.Offset(0, 3)
            strx2 = UCase(Rng.Offset(0, 2))
            Set Rng = Nothing
            If strx2 Like "EPUB" Or strx2 Like "MOBI" Or strx2 Like "PDF" Then .Label57.Caption = "�����ļ���֧�ִ˹���": Set Rng = Nothing: Exit Sub
            If FileStatus(strx7, 2) = 4 Then
                strx1 = Rng.Offset(0, 3)
                strx5 = UCase(Rng.Offset(0, 2))
                If strx5 Like "EPUB" Or strx5 Like "MOBI" Or strx5 Like "PDF" Then .Label57.Caption = "�����ļ���֧�ִ˹���": Set Rng = Nothing: Exit Sub
            Else
                Set Rng = Nothing
                .Label57.Caption = "�ļ���ʧ"
                Exit Sub
            End If
        Else
            Set Rng = Nothing
            .Label57.Caption = "�ļ���ʧ"
            Exit Sub
        End If
        If strx2 = strx5 Then
            strx3 = GetFileHashMD5(strx)
            strx4 = GetFileHashMD5(strx1)
            If Len(strx3) > 2 And Len(strx4) > 2 Then
                If strx3 = strx4 Then
                    .Label216.Caption = "Match"
                Else
                    .Label216.Caption = "MisMacth"
                End If
                .Label214.Caption = strx3
                .Label215.Caption = strx4
            Else
                .Label57.Caption = "δ��ȡ����Чֵ"
            End If
        Else
            .Label57.Caption = "���Ͳ�ƥ��"
        End If
    End With
    Set Rng = Nothing
End Sub

Private Sub CommandButton97_Click() '�Ƚ�-�Ƚ�
    Dim i As Byte, k As Byte, strLen As Byte, strlen1 As Byte
    Dim fz As Long, fz1 As Long
    Dim date1 As Date, date2 As Date, date3 As Date, date4 As Date
    Dim strx As String, strx1 As String, strx2 As String
    
    With Me
        If .CommandButton132.Caption = "�ı��Ƚ�" Then
            If Len(.Label179.Caption) > 0 And Len(.Label153.Caption) > 0 Then
                strx = .Label180.Caption
                strx1 = .Label154.Caption
                strx = Left(strx, Len(strx) - Len(Split(strx, Chr(46))(UBound(Split(strx, Chr(46))))) - 1)   '��ȡ��������չ�����ļ���
                strx1 = Left(strx, Len(strx1) - Len(Split(strx1, Chr(46))(UBound(Split(strx1, Chr(46))))) - 1)
                strLen = Len(strx)
                strlen1 = Len(strx1)
                If strLen >= strlen1 Then
                    For i = 1 To strLen
                        If InStr(strx, Mid(strx1, i, 1)) > 0 Then k = k + 1
                    Next
                Else
                    For i = 1 To strlen1
                        If InStr(strx1, Mid(strx, i, 1)) > 0 Then k = k + 1
                    Next
                End If
                .Label204.Caption = Format(k / i - 1, "0.0000") '�ļ����ƶ�
                fz = CLng(.Label181.Caption)
                fz1 = CLng(.Label155.Caption)
                .Label205.Caption = Format((fz - fz1) / fz, "0.0000") '�ļ���Сƫ��
                date1 = .Label182.Caption
                date2 = .Label156.Caption
                date3 = .Label183.Caption
                date4 = .Label157.Caption
                .Label206.Caption = DateDiff("s", date1, date2) & "s" '����ʱ��ƫ��
                .Label207.Caption = DateDiff("s", date3, date4) & "s" '�޸�ʱ��ƫ��
                If Len(.Label184.Caption) > 0 And Len(.Label158.Caption) > 0 Then '����
                    If .Label184.Caption = .Label160.Caption Then
                        .Label208.Caption = "Y"
                    Else
                        .Label208.Caption = "N"
                    End If
                End If
                If Len(.Label185.Caption) > 0 And Len(.Label159.Caption) > 0 Then .Label209.Caption = .Label185.Caption & " / " & .Label161.Caption '�򿪴���
                If Len(.Label186.Caption) > 0 And Len(.Label160.Caption) > 0 Then .Label210.Caption = .Label186.Caption & " / " & .Label162.Caption '��������
                If Len(.Label187.Caption) > 0 And Len(.Label161.Caption) > 0 Then .Label211.Caption = .Label187.Caption & " / " & .Label163.Caption '�Ƽ�ָ��
                If Len(.Label190.Caption) > 0 And Len(.Label164.Caption) > 0 Then .Label212.Caption = .Label188.Caption & " / " & .Label164.Caption 'pdf
                If Len(.Label191.Caption) > 0 And Len(.Label165.Caption) > 0 Then .Label213.Caption = .Label188.Caption & " / " & .Label164.Caption '�ı�
            End If
        Else
            strx1 = .TextBox30.Text
            strx2 = .TextBox31.Text
            If Len(strx1) = 0 Or Len(strx2) = 0 Then Exit Sub
            .Label216.Visible = True
            If CRC32API(strx1) = CRC32API(strx2) Then
                .Label216.Caption = "Mactch"
            Else
                .Label216.Caption = "MisMactch"
            End If
        End If
    End With
End Sub

Private Sub CommandButton98_Click() '�Ƚ�a-ɾ��
    Dim p As Byte, i As Byte, strx As String
    
    With Me
        strx = .Label179.Caption
        If Len(strx) = 0 Then Exit Sub
        i = FileStatus(strx)
        If i >= 5 Then .Label57.Caption = "�ļ����ڴ򿪵�״̬": Set Rng = Nothing: Exit Sub
        If i = 0 Then
            With Rng
                If Len(.Offset(0, 26).Value) > 0 Then p = 1
                Call FileDeleExc(.Offset(0, 3).Value, .Offset(0, 2).Value, .Row, p, 1, 1)
            End With
        End If
        Call ClearComp(1)
    End With
    Set Rng = Nothing
End Sub

Private Sub CommandButton101_Click() '�Ƚ�b-ɾ��
 Dim p As Byte, i As Byte, strx As String
    
    With Me
        strx = .Label153.Caption
        If Len(strx) = 0 Then Exit Sub
        i = FileStatus(strx)
        If i >= 5 Then .Label57.Caption = "�ļ����ڴ򿪵�״̬": Set Rng = Nothing: Exit Sub
        If i = 0 Then
            With Rng
                If Len(.Offset(0, 26).Value) > 0 Then p = 1
                Call FileDeleExc(.Offset(0, 3).Value, .Offset(0, 2).Value, .Row, p, 1, 1)
            End With
        End If
        Call ClearComp(2)
    End With
    Set Rng = Nothing
End Sub

Private Sub CommandButton105_Click() '�Ƚ�-���
    With Me
        If .CommandButton132.Caption = "�ı��Ƚ�" Then
            If Len(.Label179.Caption) = 0 And Len(.Label153.Caption) = 0 Then Exit Sub
            Call ClearComp(3)
        Else
            .TextBox30.Text = ""
            .TextBox31.Text = ""
            .Label216.Caption = ""
        End If
    End With
End Sub
'------------------------------------------------------�Ƚ�

Private Sub CommandButton99_Click() '�Ƚ�-��
    Dim i As Byte, strx As String, strx1 As String, strx2 As String
    
    With Me
        strx = .Label179.Caption
        strx1 = .Label180.Caption
        strx2 = LCase(Right$(strx1, Len(strx1) - InStrRev(strx1, ".")))
        If strx2 Like "xl*" Then
            If strx2 <> "xls" Or "xlsx" Then MsgBox "��ֹ�򿪴����ļ�", vbCritical, "Warning": Exit Sub
        End If
        i = FileStatus(strx)
        If i >= 5 Then .Label57.Caption = "�ļ����ڴ򿪵�״̬": Set Rng = Nothing: Exit Sub
        If i = 0 Then
        With Rng
            Call OpenFile(strx, .Offset(0, 1).Value, .Offset(0, 2).Value, .Offset(0, 3).Value, 1, .Offset(0, 26).Value, 1)
        End With
        End If
    End With
    Set Rng = Nothing
End Sub

Private Sub CommandButton100_Click() '�Ƚ�-��
    Dim i As Byte, strx As String, strx1 As String, strx2 As String
    
    With Me
        strx = .Label153.Caption
        strx1 = .Label154.Caption
        strx2 = LCase(Right$(strx1, Len(strx1) - InStrRev(strx1, ".")))
        If strx2 Like "xl*" Then
            If strx2 <> "xls" Or "xlsx" Then MsgBox "��ֹ�򿪴����ļ�", vbCritical, "Warning": Exit Sub
        End If
        i = FileStatus(strx)
        If i >= 5 Then .Label57.Caption = "�ļ����ڴ򿪵�״̬": Set Rng = Nothing: Exit Sub
        If i = 0 Then
        With Rng
            Call OpenFile(strx, .Offset(0, 1).Value, .Offset(0, 2).Value, .Offset(0, 3).Value, 1, .Offset(0, 26).Value, 1)
        End With
        End If
    End With
    Set Rng = Nothing
End Sub

Private Sub TextBox24_Change()
    With Me
        If Len(.TextBox24.Text) = 0 Then .Label109.Caption = ""
    End With
End Sub

Private Sub TextBox24_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer) '����-�������-���»س���,ִ�м������
    If KeyCode = 13 Then Call ChPN
End Sub

Private Sub CommandButton92_Click() '����-�������
    Call ChPN
End Sub

Sub ChPN() '����-�������
    Dim strx As String, strLen As Byte
    
    With Me
        strx = Trim(.TextBox24.Text)
        strLen = Len(strx)
        .Label109.Caption = ""
        If strLen > 5 Or IsNumeric(strx) = False Then Exit Sub
        If strx = 0 Then Exit Sub
        
        If CheckPN(strx) = True Then
            .Label109.ForeColor = &HC000&
            .Label109.Caption = "Y"
        Else
            .Label109.ForeColor = &HFF&
            .Label109.Caption = "N"
        End If
        .TextBox24.SetFocus
    End With
End Sub

'-------------------------------------------����¼
Private Sub ListBox4_Click() '����¼
    Dim i As Byte
    
    With Me
        .ListBox4.Visible = False
        .TextBox10.Visible = True
        i = .ListBox4.ListIndex
        .TextBox10.Text = .ListBox4.Column(1, i)  'listbox�������ַ�������ʾֵ��λ��
    End With
End Sub

Private Sub ComboBox9_Change() '����¼-����
    Dim timea As Date
    
    timea = Date
    With Me
        If .ComboBox9.Text <> CStr(timea) Then
            .CommandButton34.Enabled = False
            .TextBox10.Locked = True
        Else
            .CommandButton34.Enabled = True    '�ǵ������־���޷��޸�
            .TextBox10.Locked = False
        End If
    End With
End Sub

Sub DateUpdate()                             '����¼����ʵʱ����
    Dim dic As New Dictionary
    Dim TableName As String
    Dim i As Byte, k As Byte
    
    TableName = "����¼"
    SQL = "select * from [" & TableName & "$]"
    Set rs = New ADODB.Recordset    '������¼������
    rs.Open SQL, Conn, adOpenKeyset, adLockOptimistic
    
    'ReDim arr(1 To rs.RecordCount)
    k = rs.RecordCount
    If k = 0 Then Exit Sub
    For i = 1 To k
        dic(CStr(rs.Fields(0))) = ""
        rs.MoveNext '��ȡ���ݼ�����һ��ֵ
    Next
    Me.ComboBox9.List = dic.Keys 'ֻ��Ҫ��ȡ���ڵĸ���
    rs.Close
    Set rs = Nothing
End Sub
'------------------------------------------------------------------------------����¼

'---------------------------------------------------------------------------------------------------------------------------------------------------------������
Private Sub CommandButton17_Click() '������-����ļ���
    AddFx = 0
    If ListAllFiles(0, "NU") = False Then Me.Label57.Caption = "�ļ�������": Exit Sub
    DataUpdate
End Sub

Private Sub CommandButton1_Click() '����ļ�
    AddFx = 0
    If ListAllFiles(0, "NU") = False Then Me.Label57.Caption = "�ļ�������": Exit Sub
    DataUpdate
End Sub

Private Sub CommandButton51_Click() '����
    Dim strx As String, strx1 As String
    On Error GoTo 100
    'win10,���windows media player������,���ο����� bass, http://www.un4seen.com/,֧��vb
    With ThisWorkbook
        strx = .Path & "\whitenoise.mp4"
        strx1 = .Sheets("temp").Range("ab8").Value
    End With
    With Me
        If fso.fileexists(strx) = False Or Len(strx1) = 0 Then '��鲥���ļ��Ƿ����
            .Label57.Caption = "��Ƶ�ļ��Ѷ�ʧ"
            Exit Sub
        End If
        If .Label66.Caption = "play" Then Exit Sub '�Ѿ����ڲ���״̬
        If .Label66.Caption = "stop" Then
            wm.Controls.Play              '����ֹͣ����״̬
            .Label66.Caption = "play"
            .Label57.Caption = "���ֲ�����...."
        Exit Sub
        End If
        If Len(.Label66.Caption) = 0 Then
'            If fso.FileExists(Environ("ProgramW6432") & "\Windows Media Player\wmplayer.exe") = False Then '����Ƿ����windows media player 'Environ("ProgramW6432") programfiles
'               .Label57.Caption = "Windows media player�����ڣ��˹��ܲ�֧��"
'               Exit Sub
'            End If
            Set wm = .Controls.Add("WMPlayer.OCX.7") '����windows���ſؼ�
            If wm Is Nothing Then
                .Label57.Caption = "wm�ؼ�����ʧ��"
                Set wm = Nothing
                Exit Sub
            End If
            wm.Visible = False '����Ϊ����
            wm.url = strx1
            .Label66.Caption = "play" 'label66������ʱ�洢���ŵ�״ֵ̬�����ڿ��ư�ť
            .Label57.Caption = "���ֲ�����..."
        End If
    End With
Exit Sub
100
Me.Label57.Caption = "�����쳣"
End Sub

Private Sub CommandButton52_Click() '���ֲ���ֹͣ
    With Me
        If Len(.Label66.Caption) = 0 Or .Label66.Caption = "stop" Then Exit Sub
        wm.Controls.Stop
        .Label66.Caption = "stop"
        .Label57.Caption = "���ֲ�������ͣ"
    End With
End Sub

Private Sub CommandButton10_Click() '��ҳ
    Call BackSheet("��ҳ")
End Sub

Private Sub CommandButton11_Click() '���
    Call BackSheet("���")
End Sub

Sub BackSheet(ByVal shtn As String) '����Excel�����
    Unload Me
    With ThisWorkbook
        If .Application.Visible = False Then .Application.Visible = True
        .Windows(1).WindowState = xlMaximized
        .Sheets(shtn).Activate
    End With
End Sub

Private Sub CommandButton118_Click() '������-����
    Dim strx As String, strx1 As String, i As Integer

    With ThisWorkbook.Sheets("temp")
        strx = .Cells(39, "ab").Value
        strx1 = .Cells(39, "ac").Value
    End With
    With Me
        If Len(strx) = 0 Then
            .Label57.Caption = "�ļ��Ѷ�ʧ"
        Else
            If Len(strx1) = 0 Then
                i = ErrCode(strx, 1)
                If i < 0 Then MsgBox "�����ļ�·������쳣", vbOKOnly, "Warning": Exit Sub
                If i > 1 Then ThisWorkbook.Sheets("temp").Cells(39, "ac").Value = "ERC": strx1 = "ERC"
            End If
            If fso.fileexists(strx) = False Then
                .Label57.Caption = "�ļ��Ѷ�ʧ"
            Else
                Call OpenFile("N", "help.pdf", "pdf", strx, 1, strx1, 1)
            End If
        End If
    End With
End Sub

Private Sub CommandButton42_Click() '���ð�ť-������ҳ��һֱ�������
    Dim i As Byte, k As Byte
    
    With Me.MultiPage1
        k = .Pages.Count
        k = k - 2
        For i = 0 To k
            .Pages(i).Visible = False
        Next
        k = k + 1
        .Pages(k).Visible = True  '��ʾ����ҳ��
        .Value = k
    End With
End Sub

Private Sub CommandButton7_Click() '������-Excelģʽ
    Dim strx As String
    Dim i As Byte, k As Byte, p As Byte
    Dim wd As Object, yesno As Variant
'    ThisWorkbook.Application.ScreenUpdating = False
'    Me.Hide
'    Me.Show 0
'    Call Rewds
'    CopyToClipboard ThisWorkbook.fullname '�����ļ���·�������а�
'-------------------------------------------------------------------------
    yesno = MsgBox("�������˳�,ת��word,�Ƿ����", vbYesNo, "Tips")
    If yesno = vbNo Then Exit Sub
    yesno = MsgBox("�����رձ������е�word�ĵ�,�Ƿ����", vbYesNo, "Tips")
    If yesno = vbNo Then Exit Sub
    On Error Resume Next
    If CreateDB = False Then MsgBox "�������ݿ�ʧ��", vbCritical, "Warning": Exit Sub '------------�����µ����ݿ�
    strx = ThisWorkbook.Path
    strx = strx & "\LB.docm"
    If fso.fileexists(strx) = True Then
        If Err.Number > 0 Then Err.Clear
        Set wd = GetObject(, "word.application") '���word������������״̬
        If Err.Number > 0 And wd Is Nothing Then
            Err.Clear
            Set wd = CreateObject("word.application")
        End If
        With wd
            i = .documents.Count 'get���word�����ǿ���,δ���κε��ĵ�, createobject������������е��ĵ������
            If i > 0 Then
                For k = 1 To i
                    If strx = .documents(k).fullname Then
                        p = p + 1
                    Else
                        .documents(k).Close savechanges:=True
                    End If
                Next
                If p > 0 Then GoTo 100
            End If
            .Visible = True
            .Activate
            .documents.Open (strx)
        End With
100
        Set wd = Nothing
    Else
        Me.Label57.Caption = "�ļ���ʧ"
    End If
    MeQuit
'    ThisWorkbook.Application.ScreenUpdating = True
End Sub

Private Sub MeQuit() '�رճ���
    Unload Me
    If UF4Show > 0 Then Unload UserForm4
    MsgShow "�����Զ�����", "Tips", 1200
    With ThisWorkbook
        Call ResetMenu '�ر��Ҽ��˵�
        With .Sheets("���")
            .Label1.Caption = ""
            If .CommandButton21.Caption = "�˳�����" Then
                .CommandButton21.Caption = "����ģʽ"
                .CommandButton21.ForeColor = &H80000012
                .CommandButton1.Enabled = True
                .CommandButton11.Enabled = True
            End If
        End With
        If Workbooks.Count > 1 Then
            EnEvents '--------------������еĸ�����,��ֹ������©���ʹ��Excel������
            .Close savechanges:=True
        Else
            .Save
            .Application.EnableEvents = False 'ע������,��������workbook.close�¼�,����������¼�,excel��Ȼ�������� '�����¼���,�����´�excel�¼����Զ��ָ�
            .Application.Quit
        End If
    End With
End Sub

Private Sub CommandButton39_Click() '�رճ���
    MeQuit
End Sub

Private Sub CommandButton19_Click() '����ģʽ
    Me.Hide
    If Workbooks.Count = 1 Then
        If ThisWorkbook.Application.Visible = True Then ThisWorkbook.Application.Visible = False '��Է�excel�ļ���ҵ�Ļ���
        UserForm4.Hide
        UserForm4.Caption = "����"
        UserForm4.Show 1
    Else
        ThisWorkbook.Windows(1).WindowState = xlMinimized
        UserForm4.Show
    End If
End Sub
'----------------------------------------------------------------------------------------------������

Function CheckFileOpen(ByVal filecode As String) As Boolean '�����ؼ���ļ���״̬����Ϣ
    Dim i As Byte, strx5 As String
    
    CheckFileOpen = False
    i = FileStatus(filecode) '����ļ���״̬
    Select Case i
        Case 1: strx5 = "Ŀ¼������"
        Case 3: strx5 = "�ļ�������"
        Case 5: strx5 = "Excel���ܴ�ͬ���ļ�"
        Case 6: strx5 = "�ļ����ڴ򿪵�״̬"
        Case 7: strx5 = "�ļ����ڴ򿪵�״̬"
        Case 8: strx5 = "�쳣"
    End Select
    If i = 1 Or i = 3 Then
        DeleFileOverx filecode
    End If
    If i <> 0 Then
        Me.Label57.Caption = strx5
        Set Rng = Nothing
        CheckFileOpen = True
    End If
End Function

Private Sub CommandButton13_Click() '�ļ�ִ��-��Ҫ�޸�
    Dim i As Integer, p As Byte, xi As Byte, k As Byte, n As Single
    Dim strx3 As String, strx1 As String, strx2 As String, strx4 As String, filez As Long, wt As Integer, _
    Folderpath As String, dics As String, strx5 As String
    
    With Me
        If .Label55.Visible = True Then Exit Sub
        strx1 = .ComboBox6.Text
        strx3 = .Label29.Caption
        strx2 = Filepathi
        strx4 = Filenamei
        strx5 = .Label24.Caption '�ļ���չ��
        If Len(strx2) = 0 Or Len(strx1) = 0 Then Exit Sub '���·��Ϊ��,�ļ��ѱ�ɾ��,��ֵ
        filez = .Label76.Caption
        '�ڴ��ļ���ͬʱ,����txt�ļ�,Ȼ��ִ��vbs�ű�,ÿ��60s(����),ѭ���鿴���ļ���commandline�Ƿ񱻹ر�
        '�����׽���ر� , �ͽ���Ϣд�뵽txt�ļ�����ȥ, �ر�vbsִ��
        '��ô���Լ��ʵ���ļ��򿪺͹رյ�׷��
        If strx1 = "��" Then
            If strx5 Like "xl*" Then
                If strx5 <> "xls" Or "xlsx" Then MsgBox "��ֹ�򿪴����ļ�", vbCritical, "Warning": Exit Sub
            End If
            If .Label74 = "Y" Then strx3 = "ERC"
            If .CheckBox13.Value = True Then
                xi = 1                    '------�ļ���ѡ����
                Folderpath = ThisWorkbook.Sheets("temp").Cells(41, "ab").Value
                If Len(Folderpath) = 0 Or fso.folderexists(Folderpath) = False Then
                    Folderpath = ThisWorkbook.Path & "\protect"
                    fso.CreateFolder (Folderpath)
                    Folderpath = Folderpath & "\"
                    ThisWorkbook.Sheets("temp").Cells(41, "ab") = Folderpath
                End If
                dics = Left$(Folderpath, 1)
                Folderpath = Folderpath & "\"
                i = FileTest(Folderpath & strx4, strx5, strx4)
                If i >= 4 Then MsgBox "���ļ��Ѿ����ڱ����ļ�,�Ҵ��ڴ򿪵�״̬": Exit Sub
                If i = 2 Then '�ļ�������
                    If filez > fso.GetDrive(dics).AvailableSpace Then MsgBox "���̿ռ䲻��!", vbCritical, "Warning": Exit Sub '�жϴ����Ƿ����㹻�Ŀռ�
                    fso.CopyFile (strx2), Folderpath, True '����
                End If
                strx2 = Folderpath & strx4 '�µ��ļ�·��
            Else
                If CheckFileOpen(strx3) = True Then Exit Sub
            End If
            
            If OpenFile(strx3, strx4, .Label24.Caption, strx2, 1, strx3, xi) = False Then .Label57.Caption = "�ļ���ʧ��": Exit Sub
            
            If xi = 1 Then Set Rng = Nothing: Exit Sub
            Call OpenFileOver(strx3) '�����ƺ�
        
        ElseIf strx1 = "ɾ��" Then 'ɾ���漰,�ļ��Ƿ������Ŀ¼-�Ƿ�����ڱ���-�Ƿ��ڴ򿪵�״̬
            If CheckFileOpen(strx3) = True Then Exit Sub '����ļ��Ƿ��ڴ򿪵�״̬
            If .Label74.Caption = "Y" Then i = 1 '�Ƿ���ڷ�ansi
            If Len(ThisWorkbook.Sheets("temp").Range("ab37").Value) = 0 Then UserForm12.Show '��ɾ��ʱ����ɾ��ԭ�򴰿�
            '��userform ��ģʽ(modal)1��ʾ��ʱ��,�������뽫��ִͣ��
            If FileDeleExc(Rng.Offset(0, 3).Value, Rng.Row, i, 1) = True Then DeleFileOverx strx3 'ִ��ɾ������ɹ�-ִ��ɾ�����ƺ�����
            
        ElseIf strx1 = "��λ��" Then '�ж��ļ��Ƿ�����ڱ���
            Call OpenFileLocation(Folderpathi)
            
        ElseIf strx1 = "�����ļ�" Then
            i = FileStatus(strx3, 2)
            Select Case i
                Case 1: strx5 = "Ŀ¼������"
                Case 3: strx5 = "�ļ�������"
            End Select
            If i = 1 Or i = 3 Then
'                .Label55.Visible = ture
'                .Label56.Caption = strx3
                DeleFileOverx strx3
                .Label57.Caption = strx5
                Set Rng = Nothing
                Exit Sub
            End If
            If FileCopy(strx2, strx4, Rng.Row, 1) = True Then
                Set Rng = Nothing
                .Label57.Caption = "�����ļ���.."
                If filez < 104857600 Then    '100M
                    wt = 200
                Else
                    k = Int(filez / 104857600)
                    n = CSng(k)
                    wt = 200 * (n + n / 20 + n / 50)
                    If wt > 500 Then wt = 500
                End If
                Sleep wt
                .Label57.Caption = "�����ɹ�"
            Else
                .Label57.Caption = "����ʧ��"
            End If

        ElseIf strx1 = "��ӵ������б�" Then
            If CheckLb6(strx3) = False Then
                With .ListBox6
                    If .ListCount > 30 Then MsgBox "��������ﵽ����": Exit Sub
                    .AddItem
                     p = .ListCount - 1
                    .List(p, 0) = strx3 '���
                    .List(p, 1) = Filenamei '�ļ���
'                    .List(p, 2) = Filepathi '�ļ�·��
                End With
                .Label57.Caption = "�����ɹ�"
            Else
                .Label57.Caption = "�ļ������"
            End If
        End If
    End With
    Set Rng = Nothing
End Sub

Function CheckLb6(ByVal filecode As String) As Boolean '�жϵ����б����Ƿ����ֵ
    Dim i As Byte, k As Byte
    'Ҳ���Դ���һ��ģ�鼶��ʱ��������ʱ�洢����,���ԱȽ�
    CheckLb6 = False
    With Me.ListBox6
        k = .ListCount
        If k = 0 Then CheckLb6 = False: Exit Function
        k = k - 1
        For i = 0 To k
            If .List(i, 0) = filecode Then CheckLb6 = True: Exit Function
        Next
    End With
End Function

Function DeleFileOverx(ByVal filecodex As String) 'ɾ���ļ���ִ���ƺ�
    Dim i As Byte, itemf As ListItem
    Dim k As Byte, rnga As Range, p As Byte
    
    With Me.ListView1                  '����������е��������
        If .ListItems.Count <> 0 Then
            Set itemf = .FindItem(filecodex, lvwText, , lvwPartial)
            If itemf Is Nothing Then
                GoTo 1001
            Else
                .ListItems.Remove (itemf.Index) '�Ƴ��������
            End If
1001
            Set itemf = Nothing
        End If
    End With
    
    With Me '����Ƚ��е�����
        .Label56.Caption = filecodex
        If filecodex = .Label29.Caption Then
            .Label55.Visible = True
            DisablEdit '���ɾ���������Ǳ༭������ʾ���ļ�,��ô�ͽ��ñ༭
        End If
        If Len(.Label179.Caption) > 0 Then
            If .Label179.Caption = filecodex Then p = 1       '����Ƚ��������
        End If
        If Len(.Label153.Caption) > 0 Then
            If .Label153.Caption = filecodex Then p = p + 2
        End If
        If p > 0 Then ClearComp (p)
    End With
'    Call CwUpdate '���´��������
    With ThisWorkbook.Sheets("������")       '������,�����Ķ��޸�
        Set rnga = .Range("i27:i33").Find(filecodex, lookat:=xlWhole)
        If rnga Is Nothing Then Exit Function
        k = rnga.Row
        If k = 33 Then '��������һ��,��ôֱ�ӽ��������
            .Range("d33:l33").ClearContents
        Else
            If k = 27 Then prfile = .Range("i27").Value
            .Range("d" & k & ":" & "l" & k).ClearContents
            For i = k To 33
                If .Range("d" & i) = "" Then Exit For
                .Range("d" & i) = .Range("d" & i + 1)
                .Range("i" & i) = .Range("i" & i + 1)
                .Range("k" & i) = .Range("k" & i + 1)
            Next
        End If
        Me.ListBox2.RemoveItem (k - 27)
    End With
    Set rnga = Nothing
    DataUpdate '���´��������
End Function

Function ClearComp(ByVal cmCode As Byte) '����Ƚ�����
    Dim i As Byte, p As Byte
    p = cmCode
    With Me
        If p > 1 Then
            For i = 153 To 165
                .Controls("label" & i).Caption = ""
            Next
        End If
        If p = 1 Or p > 2 Then
            For i = 179 To 192
                .Controls("label" & i).Caption = ""
            Next
        End If
        If p = 3 Then
            For i = 204 To 216
                If i <> 215 Then .Controls("label" & i).Caption = ""
            Next
        End If
    End With
End Function

Private Sub CommandButton2_Click() '������-�Ƴ������Ķ�
    Dim i As Byte, k As Byte, m As Byte
    Dim j As Integer, p As Byte
    
    j = Me.ListBox2.ListIndex
    p = Me.ListBox2.ListCount
    If p = 0 Or j = -1 Then Exit Sub '-1��ʾû��ѡ���ļ� '������Ҫע��
    i = 27 + j
    With ThisWorkbook.Sheets("������")
        .Range("d" & i & ":" & "l" & i).ClearContents    'ע��ϲ�����Ӱ��
        Me.ListBox2.RemoveItem (j)
        If Len(.Range("d" & i + 1).Value) = 0 Then Exit Sub '�����һ��û�����ݾ��˳�
            m = 6 - j
            For k = 1 To m
                .Range("d" & i + k - 1) = .Range("d" & i + k)
                .Range("i" & i + k - 1) = .Range("i" & i + k)
                .Range("k" & i + k - 1) = .Range("k" & i + k)
            Next
            .Range("d" & i + k) = "" '���һ�������
            .Range("i" & i + k) = ""
            .Range("k" & i + k) = ""
    End With
End Sub

Private Sub CommandButton23_Click() '������-�������Ķ�
    If Me.ListBox1.ListCount = 0 Then Exit Sub
    With ThisWorkbook.Sheets("������")
        .Range("p27:x33").ClearContents    'ע��ϲ�����Ӱ��
    End With
    Me.ListBox1.Clear                   'ȫ�����
End Sub

Private Sub CommandButton24_Click() 'ִ�ж�����ɸѡ
    Dim strx1 As String, strx2 As String, strx3 As String
    Dim strx As String
    
    With Me
        If Len(.ComboBox1.Text) = 0 And Len(.ComboBox7.Text) = 0 And Len(.ComboBox8.Text) = 0 Then Exit Sub
        If .MultiPage1.Value <> 2 Then .MultiPage1.Value = 2 '�ص��ļ���ҳ�� 'listview�����ڿ�����״̬�½��и�ֵ��������λ��ƫ�Ƶ�����(�����ص�״̬�¸�ֵֻ��ʹ��list)
        strx1 = CStr(.ComboBox1.Value)
        strx2 = CStr(.ComboBox7.Value)
        strx3 = CStr(.ComboBox8.Value)
        strx = strx1 & strx2 & strx3
'        If strx1 & strx2 & strx3 = .Label81.Caption Then Exit Sub 'ͬ����ɸѡ����
'        Call FileListv(strx1, strx2, strx3, 3, 18, 19, 1)
    If Len(storagex) > 0 Then
        If storagex = strx Then Exit Sub
    End If
    FileFilterTemp strx1, , strx2, , strx3
    End With
    storagex = strx
End Sub

Private Function FileFilterTemp(ByVal targetx As String, Optional ByVal cmCode As Byte = 3, Optional ByVal targetx1 As String, _
Optional ByVal cmcode1 As Byte = 1, Optional ByVal targetx2 As String, Optional ByVal cmcode2 As Byte = 2, Optional ByVal cmCodex As Byte = 0) As Byte '��������ɸѡ/���ļ���
                                        '-------------------------------------�����ɸѡ�ķ��������ٶȸ���
    Dim i As Integer, blow As Integer, c As Integer
    Dim x1 As Byte, x2 As Byte, x3 As Byte
    
    ArrayLoad '��������
    blow = docmx - 5
    If cmCodex = 1 Then targetx = targetx & "\" 'ɸѡ���ļ���
    With Me.ListView2.ListItems
        If cmCodex = 0 Then .Clear
        If cmCodex > 0 Then '���ļ���
            For i = 1 To blow
                If cmCodex = 1 Then GoTo 98
                If arrax(i, cmCode) = targetx Then
                    If cmCodex = 2 Then GoTo 99
98
                    If InStr(arrax(i, cmCode), targetx) > 0 Then
99
                        With .Add
                            .Text = arrax(i, 1) '���
                            .SubItems(1) = arrax(i, 2) '�ļ���
                            .SubItems(2) = arrax(i, 3) '��չ��
                            .SubItems(3) = arrax(i, 5) 'λ��
                            c = c + 1
                        End With
                    End If
                End If
            Next
        Else 'ɸѡ
            x1 = Len(targetx)
            x2 = Len(targetx1)
            x3 = Len(targetx2)
            For i = 1 To blow
                If x1 = 0 Then GoTo 100
                If arrax(i, cmCode) = targetx Then
100
                    If x2 = 0 Then GoTo 101
                    If arrsx(i, cmcode1) = targetx1 Then
101
                        If x3 = 0 Then GoTo 102
                        If arrsx(i, cmcode2) = targetx2 Then
102
                            With .Add
                                .Text = arrax(i, 1) '���
                                .SubItems(1) = arrax(i, 2) '�ļ���
                                .SubItems(2) = arrax(i, 3) '��չ��
                                .SubItems(3) = arrax(i, 5) 'λ��
                                c = c + 1
                            End With
                        End If
                    End If
                End If
            Next
        End If
    End With
    If c = 0 And cmCodex = 0 Then
        Me.Label57.Caption = "δ�ҵ���Ӧ��Ϣ"
    ElseIf c > 0 And cmCodex > 0 Then
        FileFilterTemp = 1
    End If
End Function

Private Sub CommandButton27_Click() '��ӵ������Ķ�-�޸�
    Dim strx As String
    With Me
        If .Label55.Visible = True Then Exit Sub
        strx = .Label29.Caption
        If Len(strx) = 0 Then Exit Sub
        If AddPList(strx, .Label23.Caption, 1) = True Then
            Call PrReadList
'            .Label57.Caption = "��ӳɹ�"
        Else
            .Label57.Caption = "�����"
        End If
    End With
End Sub

Private Sub CommandButton28_Click() '�༭��Ϣ,��ֹ�����
  Call EnablEdit
End Sub

Private Sub CommandButton29_Click() '������-ת����ϸҳ -��������ý��������������е���,��������޸ĳ�ģ����������Ҳ��Ҫ�䶯
    Dim strx As String, k As Integer, xi As Byte
    Dim i As Integer
    With Me
        k = .ListView1.ListItems.Count
        If Len(Trim(.TextBox8.Text)) = 0 Or k = 0 Then Exit Sub
        With .ListView1
            For i = 1 To k
                If .ListItems(i).Selected = True Then xi = .SelectedItem.Index: Exit For
            Next
        End With
        If xi = 0 Then Exit Sub
        strx = .ListView1.SelectedItem.Text
        If strx = .Label29.Caption And .Label55.Visible = False Then .MultiPage1.Value = 1: Exit Sub '�Ѿ���ѯ���
        If strx = .Label56.Caption Then
            .Label57.Caption = "�ļ��ѱ�ɾ��"
            .ListView1.ListItems.Remove (xi)     '���Խ�һ����չ����ɾ�����������Ϣ'���Խ�һ����ɾ���ļ�֮ǰ��ȡ�ļ���md5,�����жϺ������ļ���ӽ����Ƿ��ظ�
            Exit Sub
        End If
        SearchFile strx      '�����ļ�
        If Rng Is Nothing Then DeleFileOverx strx: Exit Sub
    End With
    Call ShowDetail(strx)
    Set Rng = Nothing
End Sub

Private Sub CheckBox8_Click() '�༭ģʽ�Ƿ�����-��Ӧ�ı༭��ť��ȫ������,����༭��Ϣ,����Ҫ�ֶ�������ť
    If clickx = 1 Then clickx = 0: Exit Sub
    With ThisWorkbook.Sheets("temp")
        If Me.CheckBox8.Value = True Then
            .Cells(31, "ab") = 1
        Else
            .Cells(31, "ab") = ""
        End If
    End With
End Sub

'-------------------------------------------------����/�༭-����
Private Sub CommandButton3_Click() '�ٶ�����
    SearchCheck 1
End Sub

Private Sub CommandButton5_Click() '��������
   SearchCheck 2
End Sub

Private Sub CommandButton6_Click() 'ά���ٿ�
    Dim yesno As Variant
    yesno = MsgBox("��վ���Ѿ�404, �Ƿ������", vbYesNo, "Warning")
    If yesno = vbNo Then Exit Sub
    SearchCheck 4
End Sub

Function MultiSearch(ByVal engine As Byte, Keyword As String) As String
    Dim SearchEngine As String
    Dim Urlx As String
    
    Select Case engine
    Case 1
        SearchEngine = "https://www.baidu.com/s?wd="  '�ٶ�����
        Keyword = Replace(Keyword, " ", "%20")
        Urlx = SearchEngine & Keyword
    Case 2
        SearchEngine = "https://www.douban.com/search?q="  'douban����
        Keyword = Replace(Keyword, " ", "+")
        Urlx = SearchEngine & Keyword
    Case 3
        SearchEngine = "https://www.bing.com/search?q="  'bing
        Keyword = Replace(Keyword, " ", "%20")
        Urlx = SearchEngine & Keyword
    Case 4
        SearchEngine = "https://en.wikipedia.org/w/index.php?search="   'wiki
        Keyword = Replace(Keyword, " ", "%20")
        Urlx = SearchEngine & Keyword
    Case 5
        SearchEngine = "" 'Ԥ��
        Keyword = Replace(Keyword, " ", "%20")
        Urlx = SearchEngine & Keyword
    End Select
    MultiSearch = Urlx
End Function

Function SearchCheck(ByVal cmCode As Byte)  '�������������
    Dim strx1 As String, strx As String
    Dim strLen As Byte
    Dim i As Byte, k As Byte, j As Byte, m As Byte
    
    With Me
        If .CheckBox21.Value = True Then i = 1 '�����ѡ������
        If .CheckBox21.Value = True Then k = 1
        If i = 1 Or k = 1 Then
            m = 2
            If i = 1 Then               '�����������ļ���
                strx1 = Trim(.TextBox3.Text)
            ElseIf k = 1 Then
                strx1 = Trim$(Left$(Filenamei, InStrRev(Filenamei, ".") - 1)) 'ȥ����չ��
            End If
        Else
            strx1 = Trim(.TextBox1.Text)
            If strx1 Like "HLA*" Then Exit Function
        End If
        strLen = Len(strx1)
        If strLen = 0 Then Exit Function
        strx = MultiSearch(cmCode, strx1)
        If cmCode = 4 Then
            For j = 1 To strLen
                If Not Mid(strx1, j, 1) Like "[a-zA-Z]" Then .Label57.Caption = "���������ݰ�����Ӣ���ַ�,λ��:" & j: Exit Function '��������wikipedia/ֻ����Ӣ��
            Next
        End If
        If Len(ThisWorkbook.Sheets("temp").Cells(45, "ab").Value) > 0 Then m = 2 'ʹ�����������
        Webbrowser strx, m
    End With
End Function
''----------------------------------------------------------------------------------------------����/�༭-����

'-----------------------------------------------------------������-����¼
Private Sub CommandButton137_Click() '����¼-�����Ϣ
    Dim timea As Date, timeb As Date, str As String
    
    timea = Date
    timeb = Format(time, "hh:mm:ss")
    str = Me.TextBox10.Text
    If Me.TextBox10.Text <> "" And Me.ComboBox9.Text <> "" Then
        SQL = "Insert into [����¼$] (����,ʱ��,����) Values (#" & timea & "#,'" & CStr(timeb) & "', '" & str & "')"
        Conn.Execute (SQL)
        Call DateUpdate
    Else
'        Call Warning(6)
    End If
End Sub

Private Sub CommandButton32_Click() '����¼-�½�
    Dim timea As Date
    timea = Date      'dateΪ���ں���
    With Me
        .TextBox10.Enabled = True
        .TextBox10.Visible = True
        .ListBox4.Visible = False
        If .ComboBox9.Text <> CStr(timea) Then .ComboBox9.Text = CStr(timea) 'cstrΪת���ı�����
        .TextBox10.Text = ""    '��֤�µ��ı����ǿհ׵�
    End With
End Sub

Private Sub CommandButton33_Click() '����¼-�鿴
    Dim timea As Date, i As Byte, k As Byte
    Dim arr()
    Dim TableName As String
    
    TableName = "����¼"
    timea = Me.ComboBox9.Text
    With Me.ListBox4
        If RecData = True Then
            .Clear
            SQL = "select * from [" & TableName & "$] where ����=#" & timea & "#"
            Set rs = New ADODB.Recordset    '������¼������
            rs.Open SQL, Conn, adOpenKeyset, adLockOptimistic
            If rs.BOF And rs.EOF Then   '�����ж������ҵ�����
'                Call Warning(6)
                GoTo 100
            End If
            .Visible = True
            Me.TextBox10.Visible = False
            k = rs.RecordCount
            ReDim arr(1 To k, 1 To 2)
            For i = 1 To k
                arr(i, 1) = rs(1)
                arr(i, 2) = rs(2)
                rs.MoveNext
            Next
            
            For i = 1 To k
                .AddItem
                .List(.ListCount - 1, 0) = arr(i, 1) '������ʾ
                .List(.ListCount - 1, 1) = arr(i, 2)
            Next
100
            rs.Close
            Set rs = Nothing
        Else
        Me.Label57.Caption = "�쳣"
        End If
    End With
End Sub

Private Sub CommandButton34_Click() '����¼-�����Ϣ
    Dim timea As Date, timeb As String, str As String
    
    timea = Date
    timeb = Format(time, "h:mm:ss")
    str = Trim(Me.TextBox10.Text)
    If Len(str) > 0 And RecData = True Then
        SQL = "Insert into [����¼$] (����,ʱ��,����) Values (#" & timea & "#,'" & timeb & "', '" & str & "')"
        Conn.Execute (SQL)
'        Call Warning(1)
        Call DateUpdate
    Else
'        Call Warning(6)
    End If
End Sub
'-----------------------------------------------------------������-����¼

Private Sub CommandButton43_Click() '���ļ���
    Dim arrt() As String, k As Byte, i As Byte, xi As Byte, strx1 As String, strx2 As String
    Dim p As Byte, m As Byte, chk As Byte
    Dim strx As String
    
    With Me.ListBox3
        k = .ListCount
        If k = 0 Then Exit Sub  'Or .ListIndex < 1
        k = k - 1
        If Me.CheckBox9.Value = True Then '��ѡ���ļ���
            m = 4: p = 1
        Else
            ReDim arrt(0 To k)
            m = 5: p = 2
        End If
        For i = 0 To k
            If .Selected(i) = True Then
                If InStr(.List(i, 1), "�Ƴ�") = 0 Then '�Ƴ��������ڵ��ļ��е�����
                    If m = 4 Then
                        strx = arraddfolder(i)           '.List(i, 0) '��ȡ��ѡ�е��ļ��� 'arraddfolder(i)
                        Exit For
                    Else
                        xi = xi + 1
                        arrt(i) = arraddfolder(i)
                    End If
                End If
            End If
        Next
    End With
    If xi = 0 And Len(strx) = 0 Then Exit Sub
    With Me
        .MultiPage1.Value = 2          ',listview '�޷�������״̬�½��и�ֵ,��������λ��ƫ�Ƶ�����
        If pgx2 = 1 Then .ListView2.ListItems.Clear
        If m = 4 Then
            chk = chk + FileFilterTemp(strx, m, , , , , p)
        Else
            For i = 0 To k
                If Len(arrt(i)) > 0 Then chk = chk + FileFilterTemp(arrt(i), m, , , , , p)
            Next
        End If
    End With
    If chk = 0 Then Me.Label57.Caption = "�ļ���Ϊ��"
'        If .CheckBox9.Value = True Then
'            m = 4: p = 2
'        Else
'            m = 5: p = 1
'        End If
'        If m = 4 And xi > 1 Then .Label57.Caption = "������ѡ�����ļ���": Exit Sub
'        xi = xi - 1
'        .ListView2.ListItems.Clear '��ҳ��������
'        For i = 0 To xi
'            Call FileListv(arrt(i), strx1, strx2, m, 0, 0, 0, p)
'        Next
'    EnEvents
End Sub

Function FileListv(ByVal str As String, ByVal str1 As String, ByVal str2 As String, ByVal filtercode As Byte, _
ByVal filtercode1 As Byte, ByVal filtercode2 As Byte, ByVal methcode As Byte, Optional ByVal cmCode As Byte)        'ɸѡ�б���Ϣ
    Dim rngf As Range
    Dim k As Integer, spl As Integer, j As Byte, p As Byte, m As Byte, flow As Integer, alow As Integer
    Dim rngx As Range, rngxs As Range
    Dim arrlist() As Variant, strx3 As String, strx As String, str3 As String, blow As Integer
    
    On Error GoTo ErrHandle
    With ThisWorkbook.Sheets("���")       '���������ر���ʱ��,���ַ���̫��
'        flow = .[f65536].End(xlUp).Row
        flow = docmx
        If Len(str) > 0 Then
            If methcode = 0 Then
                Set rngf = .Range("f6:f" & flow).Find(str, lookat:=xlWhole) 'ɸѡ��Ŀ��Ϊ�� '�ļ���
                If rngf Is Nothing Then GoTo 101
                    j = 1
                Else
                    Set rngf = .Range("d6:d" & flow).Find(str, lookat:=xlWhole) 'ɸѡ��Ŀ��Ϊ�� '�ļ���׺
                    If rngf Is Nothing Then
                        GoTo 101
                    Else
                        j = 1
                    End If
            End If
        End If
        
        If methcode = 1 Then
            If Len(str1) > 0 Then
                Set rngf = .Range("s6:s" & .[s65536].End(xlUp).Row).Find(str1, lookat:=xlWhole) '��������
                If rngf Is Nothing Then
                    GoTo 101
                Else
                    p = 2
                End If
            End If
            If Len(str2) > 0 Then
            Set rngf = .Range("t6:t" & .[t65536].End(xlUp).Row).Find(str2, lookat:=xlWhole) '�Ƽ�ָ��
                If rngf Is Nothing Then
                    GoTo 101
                Else
                    m = 3
                End If
            End If
        End If
        '--------------------------------------------------------------------------------------------�Ƚ����ļ�����,����Ŀ¼�Ƿ��ж�Ӧ��ֵ
        If j = 0 And p = 0 And m = 0 Then GoTo 101
        DisEvents '------------------------------selection�¼������ױ�����
        If .AutoFilterMode = True Then .AutoFilterMode = False 'ɸѡ������ڿ���״̬��ر�
        Set rngx = .Range("b5:v" & flow)
        If cmCode = 0 Or cmCode = 1 Then 'ɸѡ����������
            If j = 1 Then rngx.AutoFilter Field:=filtercode, Criteria1:=str
            If p = 2 Then rngx.AutoFilter Field:=filtercode1, Criteria1:=str1
            If m = 3 Then rngx.AutoFilter Field:=filtercode2, Criteria1:=str2
        Else
            strx3 = str & "\" '-----�������Դ��ļ���
            rngx.AutoFilter Field:=filtercode, Criteria1:="=" & strx3 & "*", Operator:=xlOr 'ɸѡ���ļ��� 'Excel��ɸѡ���ԶԶ���������ɸѡ
        End If
        '-----------�����ݴ��ڵ�һ��, ɸѡ�����Ľ��Ҳ�ǵ�һ��ʱ, ������������ɸѡ���.[b65536].End(xlUp).Row����ԭ�����к�, ���6�о���ɸѡ�����Ľ��
        blow = .[b65536].End(xlUp).Row
        If blow = 6 Then
            spl = 1
        Else
            spl = .Range("b6:b" & blow).SpecialCells(xlCellTypeVisible).Count
        End If
        Set rngxs = rngx.SpecialCells(xlCellTypeVisible)
        With ThisWorkbook.Sheets("temp")
            rngxs.Copy .Range("a1") '-------------------��ɸѡ������ֵ���Ƶ���ʱ�ı����(��Ϊɸѡ�����Ľ��ͨ���ǲ�������������,�޷�һ���Ը�ֵ������)
            arrlist = ThisWorkbook.Application.Transpose(.Range("a2:d" & spl + 1).Value) '��ӵ�������
        End With
        '---------------------------------------------��ȡɸѡ�Ľ��
        If cmCode = 0 Then Me.ListView2.ListItems.Clear 'ʹ��ǰ������е�����' ����Ǵ򿪶���ļ��оͲ����
        With Me.ListView2.ListItems
            For k = 1 To spl
                With .Add
                    .Text = arrlist(1, k)
                    .SubItems(1) = arrlist(2, k)
                    .SubItems(2) = arrlist(3, k)
                    .SubItems(3) = arrlist(4, k)
                End With
            Next
        End With
        .Range("f5:f" & flow).AutoFilter
101
    End With
    ThisWorkbook.Sheets("temp").Range("a1:z" & spl + 1).ClearContents '������������
    Set rngx = Nothing
    Set rngxs = Nothing
    Set rngf = Nothing
    Me.Label81.Caption = str & str1 & str2 '��ʱ�洢��ֵ���ڿ��ƴ����ִ��
    If cmCode = 0 Then EnEvents
    Exit Function
ErrHandle:
    Me.Label57.Caption = Err.Number
    Err.Clear
    EnEvents
End Function

Private Sub CommandButton36_Click() '�����ļ���
    Dim arrt() As String, k As Byte, i As Byte, xi As Byte, p As Byte
    
    With Me.ListBox3
        k = .ListCount
        If k = 0 Or .ListIndex = -1 Then Exit Sub
        k = k - 1
        ReDim arrt(0 To k)
        For i = 0 To k
            If .Selected(i) = True Then
                If InStr(.List(i, 1), "�Ƴ�") = 0 Then xi = xi + 1: arrt(i) = .List(i, 0) '��ȡ��ѡ�е��ļ���
            End If
        Next
    End With
    If xi = 0 Then Exit Sub
    With Me
        .MultiPage1.Value = 2          ',listview '�޷�������״̬�½��и�ֵ,��������λ��ƫ�Ƶ�����
        If .CheckBox9.Value = True Then
            p = 1
        Else
            p = 2
        End If
        xi = xi - 1
        AddFx = 0
        For i = 0 To xi
            Call ListAllFiles(p, arrt(i))
        Next
    End With
    DataUpdate
End Sub

Private Sub CommandButton38_Click() '�Ƴ��ļ���
    Dim yn As Variant
    Dim i As Integer, j As Integer, k As Byte, p As Byte, xi As Byte, n As Byte
    Dim rngf As Range
    Dim flow As Integer, strx As String, strx2 As String
    
    k = Me.ListBox3.ListCount
    j = Me.ListBox3.ListIndex
    If k = 0 Or j = -1 Then Exit Sub
    yn = MsgBox("�˲�����ͬʱ�Ƴ�����Ŀ¼(����ɾ�������ļ�)?_", vbYesNo) 'msgbox��ѡyes or no
    If yn = vbNo Then Exit Sub
    
    With ThisWorkbook
        With .Sheets("���")
            flow = .[f65536].End(xlUp).Row
            p = k - 1
            For i = p To 0 Step -1 'ɾ��һ����õ�ɾ������ɾ�������λ��ƫ�Ƶ�����
                If Me.ListBox3.Selected(i) = True Then '�����ѡ��-֧�ֶ��ļ���ѡ��
                    n = n + 1
                    strx = Me.ListBox3.Column(0, i)
                    strx2 = strx & "\"
                    Set rngf = .Range("e6:e" & flow).Find(strx2, lookat:=xlPart) 'ɸѡ��Ŀ��Ϊ��
                    If rngf Is Nothing Then GoTo 100
                    If .AutoFilterMode = True Then .AutoFilterMode = False 'ɸѡ������ڿ���״̬��ر�
                    strx2 = "=" & strx2 & "*" '------------------------------ʹ��Excel��ɸѡ,-����
                    .Range("e5:e" & flow).AutoFilter Field:=1, Criteria1:=strx2, Operator:=xlAnd  '----ģ��ɸѡ����,ע������������Ҫѡ��e��, ���ļ�·�����ڵ���
                    .Range("e5").Offset(1).Resize(flow - 5).SpecialCells(xlCellTypeVisible).Delete Shift:=xlShiftUp 'ɾ����ɸѡ����ʱɸ�Ľ��
                    .Range("e5").AutoFilter
100
                    With ThisWorkbook.Sheets("������")
                        xi = 37 + i
                        .Range("e" & xi).Delete Shift:=xlUp
                        .Range("i" & xi).Delete Shift:=xlUp
                    End With
                    Call DeleFileOver(strx, 1) '�Ƴ�Ŀ¼
                    Me.ListBox3.RemoveItem (i)
                End If
            Next
        End With
    End With
    Set rngf = Nothing
    addfilec = flow - n
    Me.Label57.Caption = "�����ɹ�"
    ThisWorkbook.Save
End Sub

Private Sub CommandButton4_Click() '��ѯ���/����-ִ�����������ݻ�ȡ
    Dim strx As String, strx1 As String
    Dim strLen As Byte, keyworda As String
    
    With Me
        strx = Trim(.TextBox1.Text)
        strLen = Len(strx)
        If strLen = 0 Then Exit Sub
        If Len(.Label29.Caption) > 0 Then
            If .Label29.Caption = strx1 Then Exit Sub '�������Ѿ���ѯ֮��
        End If
        If InStr(strx, "HLA") = 0 Then Exit Sub
        If InStr(strx, "&") > 0 Then strx1 = Trim(Split(strx, "&")(0))
        If strx1 = .Label56.Caption Then Exit Sub
        If strx Like "HLA-000*&*" Then keyworda = strx1
        SearchFile keyworda
        If Rng Is Nothing Then .Label57.Caption = "�ļ�������": Exit Sub
        ShowDetail (keyworda)
    End With
    Set Rng = Nothing
End Sub

Private Sub CommandButton4_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '����/����л�-δ���
Exit Sub
    With Me
    If Len(.TextBox1.Text) = 0 Then
        .CommandButton4.Caption = "��ѯ���"
        .CommandButton44.Visible = False '����
    End If
    End With
End Sub

Private Sub TextBox11_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '����-˫��ѡ��
    Dim strfolder As String, str As String, strx As Byte
    Dim fdx As FileDialog
    Dim selectfile As String

    With Me
        str = .ComboBox11.Text
        Select Case str
            Case "�����": strx = 1
            Case "Axure": strx = 1
            Case "Mind": strx = 1
            Case "Note": strx = 1
            Case "PDF": strx = 1
            Case "��ͼ": strx = 1
            Case "Spy++": strx = 1
            Case "����": strx = 2
            Case "��ѹĿ¼": strx = 2
            Case Else: strx = 0
        End Select
        
        If strx = 1 Then
            Set fdx = Application.FileDialog(msoFileDialogFilePicker) '�ļ�ѡ�񴰿�
            With fdx
                .AllowMultiSelect = False
                .Show
                .Filters.Clear '������˹���
                .Filters.Add "Application", "*.exe" '����exe�ļ�
                If .SelectedItems.Count = 0 Then Exit Sub
                selectfile = .SelectedItems(1)
            End With
            filepathset = selectfile
            .TextBox11.Text = selectfile
            Set fdx = Nothing
        ElseIf strx = 2 Then
            With Application.FileDialog(msoFileDialogFolderPicker) '�ļ���ѡ�񴰿�(�ļ����Դ�ֻ��ѡһ��)
                .Show
                If .SelectedItems.Count = 0 Then Exit Sub 'δѡ���ļ������˳�sub
                strfolder = .SelectedItems(1)
            End With
            folderpathset = strfolder
            .TextBox11.Text = strfolder
        ElseIf strx = 0 Then
            .Label57.Caption = "��������"
        End If
    End With
End Sub

Private Sub ComboBox11_Click() '����-��������
    Dim str As String, strx As String

    str = Me.ComboBox11.Value
    With ThisWorkbook.Sheets("temp")
        Select Case str
            Case "�����": strx = .Range("ab10").Value
            Case "Axure": strx = .Range("ab11").Value
            Case "Mind": strx = .Range("ab12").Value
            Case "Note": strx = .Range("ab13").Value
            Case "PDF": strx = .Range("ab14").Value
            Case "��ͼ": strx = .Range("a15").Value
            Case "����": strx = .Range("ab16").Value
            Case "��ѹĿ¼": strx = .Range("ab42")
            Case "Spy++": strx = .Range("ab51")
        End Select
        If Len(strx) > 0 Then
            Me.TextBox11.Text = strx
        Else
            Me.TextBox11.Text = "δ����"
        End If
        If str = "����" Then
            Me.TextBox22.Enabled = True
            If Len(.Range("ab27").Value) > 0 Then
            If IsNumeric(.Range("ab27").Value) = True Then Me.TextBox22.Text = .Range("ab27").Value
            End If
        Else
            Me.TextBox22.Text = ""
            Me.TextBox22.Enabled = False
        End If
    End With
End Sub

Private Sub CommandButton41_Click() '����-�޸�����
    Dim str As String, strx As String, limitnum As Byte, strx2 As String

    On Error GoTo 100
    With Me
        strx = Trim(.TextBox11.Text)
        str = Trim(.ComboBox11.Value)
        If Len(strx) = 0 Or strx = "��������" Or strx = "δ����" Then Exit Sub  '��Ч����
        If str = "����" Or str = "��ѹĿ¼" Then
            strx = folderpathset
            If fso.folderexists(strx) = False Then '�����ļ�������
                .Label57.Caption = "���ļ��в�����"
                .TextBox11.SetFocus
                Exit Sub
            End If
            If strx = ThisWorkbook.Path Then
                .Label57.Caption = "���������úͳ���ͬһ���ļ���"
                .TextBox11.Text = ""
                .TextBox11.SetFocus
                Exit Sub
            End If
            If Right(strx, 1) = "\" Then
                .Label57.Caption = "�ļ������ú��治��Ҫ\����"
                .TextBox11.SetFocus
                Exit Sub
            End If
            If CheckFileFrom(strx) = True Then
                .Label57.Caption = "ϵͳ������λ��"
                .TextBox11.Text = ""
                .TextBox11.SetFocus
                Exit Sub
            End If
            If str = "����" Then
                If fso.GetFolder(strx).Files.Count > 0 Then
                    .Label57.Caption = "���ļ����Ѵ����ļ�"
                    .TextBox11.Text = ""
                    .TextBox11.SetFocus
                End If
                If Len(.TextBox22.Text) > 0 Then
                    If IsNumeric(.TextBox22.Text) = True Then
                       limitnum = Int(.TextBox22.Text)
                       If limitnum < 6 Then
                            .Label57.Caption = "���ù���,����ֵӦ����5" 'Ĭ������Ϊ10
                            GoTo 100
                        ElseIf limitnum > 30 Then
                            .Label57.Caption = "���ù���,����ֵӦС��30"
                            GoTo 100
                        End If
                        ThisWorkbook.Sheets("temp").Range("ab27") = limitnum
                   End If
                End If
            End If
            folderpathset = ""
        Else
            strx = filepathset
            If InStr(strx, "exe") = 0 Or InStr(strx, "\") = 0 Or fso.fileexists(strx) = False Then '��������
                .Label57.Caption = "���򲻴���"
                .TextBox11.SetFocus
                Exit Sub
            End If
        End If
        filepathset = ""
         .Label57.Caption = "���óɹ�"
    End With
    With ThisWorkbook.Sheets("temp") '������д����
        Select Case str
            Case "�����": .Range("ab10") = strx
            Case "Axure": .Range("ab11") = strx
            Case "Mind": .Range("ab12") = strx
            Case "Note": .Range("ab13") = strx
            Case "PDF": .Range("ab14") = strx
            Case "��ͼ": .Range("a15") = strx
            Case "����": .Range("ab16") = strx
            Case "��ѹĿ¼": .Range("ab42") = strx
            Case "Spy++": .Range("ab51") = strx
        End Select
    End With
    Exit Sub
100
    Me.Label57.Caption = "�����쳣,����: " & Err.Number
    Err.Clear
End Sub

Private Sub CheckBox1_Click() '����-��ѡ����-����
    If clickx = 1 Then clickx = 0: Exit Sub
    With ThisWorkbook.Sheets("temp")
        If Me.CheckBox1.Value = True Then
            .Range("ab18") = 1
            voicex = 1
        Else
            .Range("ab18") = ""
            voicex = 0
        End If
    End With
End Sub

Private Sub CheckBox7_Click() '�ļ�md5�Զ�д������
    If clickx = 1 Then clickx = 0: Exit Sub
    With ThisWorkbook.Sheets("temp")
        If Me.CheckBox7.Value = True Then
            .Range("ab29") = 1
        Else
            .Range("ab29") = ""
        End If
    End With
End Sub

Private Sub CommandButton44_Click() 'δ���-Ӣ�ķ���
Exit Sub
'    If Len(Me.TextBox1.text) = 0 Or Me.CheckBox1.Value = False Then Exit Sub '����ջ��߲���ѡ���˳�
'    If IsNumeric(Me.TextBox1.text) Or Me.TextBox1.text Like "*[һ-��]*" Then Exit Sub '�����ֻ��߰�������
'    If Me.CommandButton4.Caption = "���ʲ�ѯ" Then Application.Speech.Speak (Me.TextBox1.text)
End Sub

Private Sub CommandButton45_Click() '����-����
    Dim strx As String
    
    strx = Me.TextBox14.Text
    If Len(strx) = 0 Then Exit Sub '����ջ��߲���ѡ���˳�
    If strx Like "*#*" Or strx Like "*[һ-��]*" Then Exit Sub '�������ֻ�������/��ͬ��ϵͳ��װ�����Է���ģ�鲻һ��
    Call Speakvs(strx)
End Sub

Private Sub CommandButton46_Click() '��ӵ���-δ���

Exit Sub '��δ���
If Len(Me.TextBox13.Text) > 0 And Len(Me.TextBox14.Text) > 0 Then

'Call Warning(1)
End If
End Sub
'------------------------------------------------------------------------------------------���ع���
Function SendTools(ByVal toolx As Byte) 'ѡ����Ҫִ�еĳ���
    Dim xtool As String, exepath As String
    
    With ThisWorkbook.Sheets("temp")
        Select Case toolx
            Case 1: exepath = .Range("ab11").Value
            Case 2: exepath = .Range("ab12").Value
            Case 3: exepath = .Range("ab13").Value
            Case 6: exepath = .Range("ab14").Value
            Case 7: exepath = .Range("ab15").Value
            Case 8: exepath = .Range("ab51").Value
        End Select
    End With
    If Len(exepath) = 0 Or InStr(exepath, "exe") = 0 Or fso.fileexists(exepath) = False Then
        Me.Label57.Caption = "δ���ó���"
        Exit Function
    End If
    xtool = exepath & Chr(32)
    Shell xtool, vbNormalFocus
End Function

Private Sub CommandButton142_Click()
    Call SendTools(8)
End Sub
Private Sub CommandButton81_Click() '����-powershell ISE
    If Len(ThisWorkbook.Sheets("temp").Range("ab5").Value) = 0 Then
        Me.Label57.Caption = "��֧�ִ˹���"
        Exit Sub
    End If
    Shell ("PowerShell_ISE "), vbNormalFocus
End Sub

Private Sub CommandButton78_Click() '����-powershell
    If Len(ThisWorkbook.Sheets("temp").Range("ab4").Value) = 0 Then
        Me.Label57.Caption = "��֧�ִ˹���"
        Exit Sub
    End If
    Shell ("powershell "), vbNormalFocus '��powershell
End Sub

Private Sub CommandButton82_Click() '����-vbe
    Unload Me
    With ThisWorkbook
        If .Application.Visible = False Then .Application.Visible = True
        If .Windows(1).Visible = False Then .Windows(1).Visible = True
        .Application.SendKeys ("%{F11}")
    End With
End Sub

Private Sub CommandButton83_Click() '����-˼ά��ͼ
    Call SendTools(2)
End Sub

Private Sub CommandButton84_Click() '����-axure
    Call SendTools(1)
End Sub

Private Sub CommandButton31_Click() '����-��ͼ����
    Call SendTools(7)
End Sub

Private Sub CommandButton85_Click() '����-onenote
    Call SendTools(3)
End Sub

Private Sub CommandButton70_Click() '����-pdf�༭
    Call SendTools(6)
End Sub

Private Sub CommandButton79_Click() '����-����
    Dim strx As String, limitnum As Integer, i As Integer, k As Integer
    Dim fd As Folder, fl As File
    Dim timea As Date, Oldestfile As String
    
    With ThisWorkbook.Sheets("temp")
        strx = .Range("ab26").Value
        
        If Len(strx) > 0 And fso.folderexists(strx) = True Then
            timea = Now '�Աȵĳ�ʼֵ
            ThisWorkbook.Save   '�����ļ�
            fso.CopyFile (ThisWorkbook.fullname), strx & "\", overwritefiles:=True  '�����ļ����µ��ļ���
            fso.GetFile(strx & "\" & ThisWorkbook.Name).Name = CStr(Format(timea, "yyyymmddhmmss")) & ".xlsm" '���ļ��������ڽ���������
            .Range("ab19") = Now '���ݵ�ʱ��
        Else
            Me.Label57.Caption = "�ļ�����������"
            Exit Sub
        End If
        
        If Len(.Range("ab27").Value) > 0 And IsNumeric(.Range("ab27").Value) = True Then '�����ļ����ڵı����ļ�������
            limitnum = .Range("ab28").Value
            If limitnum < 6 Or limitnum > 30 Then limitnum = 10 '�����ļ�������5,����30
        Else
            limitnum = 10 'Ĭ��Ϊ10
        End If
        
        Set fd = fso.GetFolder(strx)
        k = fd.Files.Count
        If k > limitnum Then '�ļ�������������
            For Each fl In fd.Files '�ҳ���ɵ��ļ�
                If fl.DateCreated < timea Then
                    timea = fl.DateCreated
                    Oldestfile = fl.Path
                End If
            Next
            fso.DeleteFile (Oldestfile) 'ɾ������ɵ��ļ�
        End If
        Me.Label57.Caption = "���ݳɹ�"
    End With
    Set fd = Nothing
End Sub           '----------------------------------------------------------------------------����-���ع���

Private Sub Label106_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '��������
    Dim strx As String

    strx = Me.Label106.Caption
    If Len(strx) = 0 Or InStr(strx, "http") = 0 Then Exit Sub
    browserkey = strx
    Me.MultiPage1.Value = 6
    CopyToClipboard strx
    '---------------------�򿪶������վ����(�������վ����,������chrome/Firefox,IE,�����Ƿ���������,�򿪶��궼����ܿ�(�������վ���������edge(�ɿ�)����ģʽ,�ȵ����ݲ�������,��������ʾ����)
'    Me.Label57.Caption = "���ڴ���վ��..�Ժ�"
    '--------------------���������ȫ����ҳ���ڵ��������
'    Call WebBrowser(strx)
End Sub

Private Sub TextBox17_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '˫���򿪶��������
    Dim strx As String

    strx = Me.TextBox17.Text
    If Len(strx) = 0 Or InStr(strx, "http") = 0 Then Exit Sub
    browserkey = strx
    Me.MultiPage1.Value = 6
    CopyToClipboard strx
'    Call WebBrowser(strx)
End Sub
'---------------------------------------------------------����-���繤��

Function SendUrl(ByVal httpx As Byte, Optional ByVal cmCode As Byte) '�򿪵���վѡ��
    Dim Urlx As String
    
    Select Case httpx
        Case 1: Urlx = "http://www.iciba.com/"
        Case 2: Urlx = "https://note.youdao.com/web"
        Case 3: Urlx = "https://shouqu.me/my.html"
        Case 4: Urlx = "https://docs.microsoft.com/zh-cn/office/vba/api/overview/"
        Case 5: Urlx = "https://www.pstips.net/"
        Case 6: Urlx = "https://stackoverflow.com/"
        Case 7: Urlx = "http://club.excelhome.net/forum-2-1.html"
    End Select
    Call Webbrowser(Urlx, cmCode)
End Function

Private Sub CommandButton80_Click() '��ɽ�ʵ�
    Call SendUrl(1)
End Sub

Private Sub CommandButton25_Click() '�е��ʼ�
    Call SendUrl(2, 1)
End Sub

Private Sub CommandButton26_Click() '��Ȥ��ǩ
    Call SendUrl(3, 1)
End Sub

Private Sub CommandButton86_Click() 'VBA docs
    Call SendUrl(4)
End Sub

Private Sub CommandButton87_Click() 'ps tips
    Call SendUrl(5)
End Sub

Private Sub CommandButton89_Click() 'stack
    Call SendUrl(6)
End Sub

Private Sub CommandButton88_Click() 'excel
    Call SendUrl(7)
End Sub
'------------------------------------------------------------------------------����-���繤��

Private Sub CommandButton90_Click() '����-����
    With ThisWorkbook
        If .Application.Visible = False Then .Application.Visible = True
        If .Windows(1).Visible = False Then .Windows(1).Visible = True
    End With
    End 'end�����ִ���൱��vbe��������ð�ť,���е�sub������,������ȫ�ͷ�
End Sub

Private Sub ListView2_Click() '�����ʱ��,ȫ��ѡ��
    With Me.ListView2
        If .ListItems.Count = 0 Then Exit Sub
        .FullRowSelect = True
    End With
End Sub

Private Sub ListView2_DblClick() '�ļ�-˫�����ļ���
    Dim k As Byte, n As Byte, p As Byte, i As Byte
    Dim strx As String, strx1 As String, strx2 As String
    
    With Me.ListView2
        If .ListItems.Count = 0 Then Exit Sub
        strx = .SelectedItem.Text
        strx1 = .SelectedItem.ListSubItems(1).Text
        strx2 = LCase(Right$(strx1, Len(strx1) - InStrRev(strx1, ".")))
        If strx2 Like "xl*" Then
            If strx2 <> "xls" Or "xlsx" Then MsgBox "��ֹ�򿪴����ļ�", vbCritical, "Warning": Exit Sub
        End If
        If CheckFileOpen(strx) = True Then Exit Sub '�ж��ļ��Ƿ������Ŀ¼���߱��ش���/�ļ��Ƿ��ڴ򿪵�״̬
    End With
    With Rng
        If OpenFile(strx, .Offset(0, 1).Value, .Offset(0, 2).Value, .Offset(0, 3).Value, 1, .Offset(0, 27).Value) = False Then
            Me.Label57.Caption = "�쳣"
            Set Rng = Nothing
            Exit Sub
        End If
    End With
    Call OpenFileOver(strx)
    Set Rng = Nothing
End Sub

Private Sub TextBox14_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) '���������, ������������
    With Me
        If KeyAscii > asc(0) And KeyAscii < asc(9) Then '�������������־���ʾ
            .Label57.Caption = "��ֹ��������"
            KeyAscii = 0
            .TextBox14.Text = ""
        Else
            .Label57.Caption = "" '��������������ݵ�ʱ����վ�����Ϣ
        End If
    End With
End Sub

Private Sub CommandButton47_Click() '���ʲ�ѯ '��Ҫ���� 'ͳһʹ�ý�ɽ�ʰ�
    Dim strx As String, strx1 As String, strx2 As String
    Dim strLen As Byte
    Dim i As Byte, k As Byte, j As Byte, bl As Byte, c As Byte
    
    On Error GoTo 100
    With Me
        strx = Trim(.TextBox14.Text) 'ȥ���������Ǳ�ڵĿո�
        strLen = Len(strx)
        If strLen = 0 Or strLen > 16 Or strx Like "*#*" Then Exit Sub '���������null or ̫��/�ǰ������־�ֱ���˳� -���ڻ�ȡ�������ݵ��ٶȲ���,һ��Ĳ���Ҫ�����󶼹��˵�
        bl = UBound(Split(strx, Chr(32))) '��ȡ�ո������
        If bl > 1 Then
            .TextBox13.Text = Replace(strx, Chr(32), "", 1, 1) '���ƿո������,��ȥ�����е�һ���ո�
            .Label57.Caption = "�������������"
            Exit Sub
        End If
        
        If TestURL("https://www.baidu.com") Then '������������
            For i = 1 To strLen '�ж����������ȫ���Ƿ�Ϊ���Ļ���Ӣ�� '�м�ո������
                strx1 = Mid$(strx, i, 1)
                If strx1 Like "[a-zA-Z]" Then
                    k = k + 1
                ElseIf strx1 Like "[һ-��]" Then
                    j = j + 1
                End If
            Next
            If bl = 1 Then j = j + 1: k = k + 1 '������ڿո�,��+1 '����Ӣ�Ĵ���һ���ո�
            If k = strLen Then
                If k < 2 Or k > 16 Then                 'Ӣ������,��������ݹ���,���������˵�
                    .Label57.Caption = "��������ݿ�������"
                    Exit Sub
                End If
                If bl = 1 Then '������ڿո�,�͸ĵ��ý�ɽ
                    c = 3
                Else
                    c = 1
                End If
            ElseIf j = strLen Then            'ֻ�������������ȷִ��-��������
                If j > 6 Then                                                        '�����������ݵĳ���
                    .Label57.Caption = "��������ݿ�������"
                    Exit Sub
                End If
                If bl = 1 Then strx = Replace(strx, Chr(32), "") '��ֹ���Ĵ��ڿո�
                c = 2
            Else
                .Label57.Caption = "��������ݴ��ڷ���/��Ӣ�Ļ������"
                Exit Sub
            End If
        Else
            .Label57.Caption = "���������쳣"
            Exit Sub
        End If
        strx2 = GetdicMeaning(strx, c)
        If c = 3 Or c = 2 Then strx2 = Replace(strx2, "����", "����: ", 1, 1) '�Է��صĽ�������Ż�
        If Right$(strx2, 1) = Chr(59) Then strx2 = Left$(strx2, Len(strx2) - 1) 'chr(59)=";",�ұ����һ������ȥ��
        .TextBox13.Text = strx2
        If c = 1 Or c = 3 Then '��ѡ����
            If voicex = 1 Then Speakvs (strx)
        End If
    End With
    Exit Sub
100
    If Err.Number <> 0 Then
    Me.Label57.Caption = "�쳣: " & Err.Number
    Err.Clear
    End If
End Sub

Private Sub CommandButton53_Click() '��ȡ��������
    Dim strx As String
    
    With Me
        If .CommandButton53.Caption = "�������ֻ�ȡ" Then
            strx = Trim(.TextBox3.Text)
            If Len(strx) < 2 Or .Label56.Caption = "δ�ҵ��鼮��Ϣ" Then Exit Sub '����Ǳ�Ҫ��ִ��.label56��ʱ�洢ִ�е�״̬,������ģ�鼶�����滻��
            Call DoubanBook(strx)
        Else
            SearchFile (.Label29.Caption) '�༭������Ϣ
            If Rng Is Nothing Then .Label57.Caption = "�ļ���ʧ": Exit Sub
            .TextBox15.Text = Rng.Offset(0, 23).Value '����
            .TextBox16.Text = Rng.Offset(0, 24).Value '����
            .TextBox17.Text = Rng.Offset(0, 25).Value '����
            .TextBox15.Visible = True
            .TextBox16.Visible = True
            .TextBox17.Visible = True
            CommandButton54.Visible = True
            Set Rng = Nothing
        End If
    End With
End Sub

Private Sub CommandButton54_Click() '��Ӷ�������
    Dim strx As String, strx1 As String, strx2 As String, strx3 As String, strx4 As String, strx5 As String
    Dim arr() As String, strx7 As String, strx6 As String, strx8 As String
    With Me
        strx1 = .TextBox17.Text
        strx8 = .Label29.Caption
        If Len(strx1) = 0 Or Len(strx8) = 0 Then Exit Sub
        If FileStatus(strx8, 2) = 4 Then '����ļ��Ƿ����Ŀ¼/���ش���
            strx = .TextBox16.Text
            Rng.Offset(0, 23) = .TextBox15.Text '����
            Rng.Offset(0, 24) = strx '
            Rng.Offset(0, 25) = strx1 '�鼮����
            .Label106.Caption = strx1
            .Label69.Caption = strx '����
            '---------------------------------------------��һ����ȡ������Ϣ
            If Len(ThisWorkbook.Sheets("temp").Cells(53, "ab").Value) = 0 Then
                If TestURL(BdUrl) = False Then .Label57.Caption = "���粻����": Exit Sub
                ReDim arr(2)
                arr = ObtainDoubanPicture(strx1) '����ͼƬ����/����,����
                strx2 = arr(0)
                strx3 = .Label29.Caption
                strx4 = ThisWorkbook.Sheets("temp").Cells(46, "ab").Value
                If Len(strx4) = 0 Then
                    strx5 = ThisWorkbook.Path & "\bookcover"
                    ThisWorkbook.Sheets("temp").Cells(46, "ab").Value = strx5
                    strx4 = strx5
                Else
                    If fso.folderexists(strx4) = False Then
                        strx5 = ThisWorkbook.Path & "\bookcover"
                        fso.CreateFolder (strx5)
                        ThisWorkbook.Sheets("temp").Cells(46, "ab").Value = strx5
                        strx4 = strx5
                    End If
                End If
                strx4 = strx4 & "\" '�洢λ��
                strx7 = LCase(Right(strx2, 3))
                If strx7 = "jpg" Or strx7 = "png" Then '����Ҫ��
                    strx6 = strx4 & strx3 & "." & strx7
                    If DownloadFilex(strx2, strx6) = True Then
                        Rng.Offset(0, 34) = strx2  '����
                        Rng.Offset(0, 36) = strx6 '�����ļ�·��
                        Imgurl = strx6
                        .Frame2.Width = 158
                        .TextBox2.Width = 143
                        .CommandButton134.Left = 610
                        .CommandButton125.Left = 654
                        With .Label239
                            .Visible = True
                            .Left = 728
                            .Top = 94
                            .Caption = "�������"
                        End With
                        With .Image1
                            .Left = 708
                            .Top = 108
                            .Width = 84
                            .Height = 122
                            .Visible = True
                            .Picture = LoadPicture(strx6)
                            .PictureSizeMode = fmPictureSizeModeStretch '����ͼƬ
                        End With
                        imgx = 1
                    Else
                        If imgx = 1 Then
                            Rng.Offset(0, 34) = strx2
                            If strx2 <> Rng.Offset(0, 36).Value Then Rng.Offset(0, 36) = ""
                            '-----------------���ͼƬ�ļ�û�����سɹ�,��ԭ�е�ͼƬ·����һ��, ��ԭ����ͼƬ��ô�������ԭ�е�·��
                            With Me
                                .Image1.Visible = False
                                .Label239.Visible = False
                                .Frame2.Width = 246
                                .TextBox2.Width = 231
                                .CommandButton134.Left = 698
                                .CommandButton125.Left = 742
                            End With
                            imgx = 0
                            Imgurl = ""
                        End If
                    End If
                End If
                Rng.Offset(0, 37) = arr(2) '����
                If Len(arr(1)) > 0 Then
                    Rng.Offset(0, 14) = arr(1) '����
                    .TextBox4.Text = arr(2) & arr(1)
                End If
            End If
            Set Rng = Nothing
            .Label57.Caption = "��ӳɹ�"
        Else
            .Label57.Caption = "�ļ���ʧ"
        End If
    End With
End Sub

Private Sub CommandButton55_Click() '�ļ�md5����
    Dim p As Integer, strx As String
    
    With Me
        If Len(.Label71.Caption) > 0 Or Len(.Label25.Caption) = 0 Or Me.Label55.Visible = True Then Exit Sub
        If FileStatus(.Label29.Caption, 2) <> 4 Then Set Rng = Nothing: Exit Sub '�ж��ļ��Ƿ����
        If .Label74.Caption = "Y" Then '·�����Ƿ���ڷ�ansi�����ַ�
            p = 2
        ElseIf .Label74.Caption = "N" Then
            p = 1
        End If
        strx = GetFileHashMD5(.Label25.Caption, p)
        If Len(strx) = 2 Then .Label57.Caption = "δ�����Чֵ": Exit Sub
        .Label71.Caption = strx
        If Len(ThisWorkbook.Sheets("temp").Range("ab29").Value) = 0 Then '�����ѡ�Զ�д��md5
            .CommandButton56.Enabled = True
            .Label57.Caption = "�������"
        Else
            Call WriteMd5(1)
        End If
'        Warning (1) '�������
    End With
    Set Rng = Nothing
End Sub

Function WriteMd5(ByVal xi As Byte) 'д��md5
    With Me
        If xi = 0 Then SearchFile (.Label29.Caption)
        If Rng Is Nothing Then .Label57.Caption = "���ʧ��": Set Rng = Nothing: Exit Function
        Rng.Offset(0, 9) = .Label71.Caption
        Set Rng = Nothing
        Me.Label57.Caption = "��ӳɹ�"
'        Me.CommandButton55.Enabled = False
    End With
End Function

Private Sub CommandButton56_Click() '�༭-��¼�ļ�hash/������Ҫ���,md���㲿���Ѿ�����
    Call WriteMd5(0)
End Sub

Private Sub CommandButton57_Click() '����-��-��ת��
    Dim strx As String, strx1 As String, strLen As Byte, i As Byte, k As Byte
    
    With Me
        strx = Trim(.TextBox18.Text)
        strLen = Len(strx)
        If strLen = 0 Then Exit Sub
        
        If IsNumeric(strx) Then
            .TextBox18.Text = ""
            .Label57.Caption = "���������봿����"
            Exit Sub ''���ȫ�������־��˳�
        ElseIf strLen > 30 Then
            .Label57.Caption = "������ַ�������,����30"
            Exit Sub
        End If
        
        For i = 1 To strLen
            strx1 = Mid(strx, i, 1)
            If strx1 Like "[һ-��]" Then k = k + 1 '�ж��Ƿ��������
        Next
        
        If k = 0 Then
            .TextBox18.Text = ""
            .Label57.Caption = "�����Ϊ��Ч��Ϣ"
            Exit Sub
        End If
        
        .TextBox19.Text = SC2TC(strx)
    End With
End Sub

Private Sub CommandButton58_Click() '����-��-��ת��
    Dim strx As String, strx1 As String, strLen As Byte, i As Byte, k As Byte
    
    With Me
        strx = Trim(.TextBox18.Text)
        strLen = Len(strx)
        If strLen = 0 Then Exit Sub
        
        If IsNumeric(strx) Then
            .TextBox18.Text = ""
            .Label57.Caption = "���������봿����"
            Exit Sub ''���ȫ�������־��˳�
        ElseIf strLen > 30 Then
            .Label57.Caption = "������ַ�������,����30"
            Exit Sub
        End If
        
        For i = 1 To strLen
            strx1 = Mid(strx, i, 1)
            If strx1 Like "[һ-��]" Then k = k + 1
        Next
        
        If k = 0 Then
            .TextBox18.Text = ""
            .Label57.Caption = "�����Ϊ��Ч��Ϣ"
            Exit Sub
        End If
        
        .TextBox19.Text = TC2SC(.TextBox18.Text)
    End With
End Sub

Private Sub CommandButton59_Click() '����-�����ļ�����md5-��Ҫ�޸�
    Dim strx As String, strx1 As String

    With Me
        strx = .TextBox20.Text
        If Len(strx) > 4 Then
            If fso.fileexists(strx) = False Then Exit Sub '�������ļ�·������C:\a '4
            .Label57.Caption = "������..."
            strx1 = GetFileHashMD5(strx)
            If Len(strx1) = 2 Then .Label57.Caption = "δ�����Чֵ": Exit Sub
            .TextBox21.Text = UCase(strx1)
            .Label57.Caption = "�������"
        End If
    End With
End Sub

Private Sub CommandButton60_Click() '����md5
    With Me.TextBox21 '��������textbox�����������Ϊʵ�ָ�����Ϣ�����а�ļ��;��
        If Len(.Text) = 0 Then Exit Sub
        .Text = .Value
        .SelStart = 0
        .SelLength = Len(.Text)
        .Copy
    End With
End Sub

Private Sub CommandButton61_Click() 'ת���ַ���-md5-crc32-sha256
    Dim strx As String, strx1 As String, strx2 As String, i As Byte
    
    With Me
        strx = Trim(.TextBox18.Text)
        If Len(strx) = 0 Then Exit Sub
        strx1 = ThisWorkbook.Sheets("temp").Cells(38, "ab").Value
        If Len(strx1) = 0 Then
            i = 0
        Else
            If IsNumeric(strx1) = True Then
                i = Int(strx1)
            Else
                i = 0
            End If
        End If
        
        Select Case i
            Case 1: strx2 = GetMD5Hash_String(strx)
            Case 2: strx2 = CRC32API(strx1)
            Case 3: strx2 = SHA256Function(strx)
            Case Else
                strx2 = GetMD5Hash_String(strx)
        End Select
        .TextBox19.Text = strx2
    End With
End Sub

Private Sub CommandButton62_Click()
    With Me.TextBox19 '��������textbox�����������Ϊʵ�ָ�����Ϣ�����а�ļ��;��
        If Len(.Text) = 0 Then Exit Sub
        .Text = .Value
        .SelStart = 0
        .SelLength = Len(.Text)
        .Copy
    End With
End Sub

Private Sub CommandButton63_Click() '�ļ���-�۵����еĽڵ�
    Dim i As Integer, k As Integer
    With Me.TreeView1
        k = .Nodes.Count
        If k = 0 Then Exit Sub
        For i = 2 To k
            .Nodes(i).Expanded = False
        Next
    End With
    With Me
        If Len(.Label240.Caption) > 0 Then
            .Label240.Caption = ""
            .Label241.Caption = ""
        End If
    End With
End Sub

Private Sub CommandButton64_Click() '�ļ���-չ�����еĽڵ�
    Dim i As Integer, k As Integer
    
    With Me.TreeView1
        k = .Nodes.Count
        If k = 0 Then Exit Sub
        For i = 1 To k
            .Nodes(i).Expanded = True
        Next
    End With
End Sub

Private Sub TextBox28_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With Me.TextBox28 '��������textbox�����������Ϊʵ�ָ�����Ϣ�����а�ļ��;��
        If Len(.Text) = 0 Then Exit Sub
        .Text = .Value
        .SelStart = 0
        .SelLength = Len(.Text)
        .Copy
    End With
    Me.Label57.Caption = "�����Ѹ���"
End Sub

Private Sub TextBox29_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '����-�ļ���ѹ
    Dim fdx As FileDialog, strfolder As String
    Dim selectfile As Variant

    If Me.CheckBox14.Value = True Then
    With ThisWorkbook.Application.FileDialog(msoFileDialogFolderPicker) '�ļ���ѡ�񴰿�(�ļ����Դ�ֻ��ѡһ��)
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub 'δѡ���ļ������˳�sub
        strfolder = .SelectedItems(1)
        folderpathc = strfolder
        If ErrCode(folderpathc, 1) > 1 Then MsgShow "�ļ�·��������ansi����,�����ֶ��޸����ݿ����Ϣ", "Tips", 1800
        Me.TextBox29 = folderpathc
        Me.TextBox29.SetFocus
        Exit Sub
    End With
    End If
    Set fdx = ThisWorkbook.Application.FileDialog(msoFileDialogFilePicker)
    With fdx
        .AllowMultiSelect = False '������ѡ�����ļ�(ע�ⲻ���ļ���,�ļ���ֻ��ѡһ��)
        .Show
        .Filters.Clear                                  '������й���
        .Filters.Add "Zip", "*.7z; *.zip; *.rar", 1     'ɸѡ�ļ�
        .Filters.Add "All File", "*.*", 1
        If .SelectedItems.Count = 0 Then Exit Sub
        filepathc = .SelectedItems(1)
        If ErrCode(filepathc, 1) > 1 Then MsgShow "�ļ�·��������ansi����,�����ֶ��޸����ݿ����Ϣ", "Tips", 1800
        Me.TextBox29 = filepathc
    End With
    Me.TextBox29.SetFocus
    Set fdx = Nothing
End Sub

Private Sub TextBox3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With Me.TextBox3 '��������textbox�����������Ϊʵ�ָ�����Ϣ�����а�ļ��;��
        If Len(.Text) = 0 Then Exit Sub
        .Text = .Value
        .SelStart = 0
        .SelLength = Len(.Text)
        .Copy
    End With
End Sub

Private Sub TreeView1_DblClick() '��״ͼ�ڵ�˫��չ��
    With Me
        .TreeView1.Nodes(.TreeView1.SelectedItem.Index).Expanded = True
    End With
End Sub

Private Sub CommandButton65_Click() '�����ļ���
    Dim rnglistx As Range
    Dim filexpath As String
    Dim strx As String, strx1 As String
    Dim k As Integer, addcodex As Byte, j As Integer, i As Integer
    
    ReDim arrch(1 To 50)
    With Me
        If .TreeView1.Nodes.Count = 0 Then Exit Sub
        k = .TreeView1.SelectedItem.Index
        If k = 1 Then Exit Sub
        strx = .TreeView1.Nodes(k).Text           '-��ֹ�޷����µ��ļ�
        If InStr(strx, "�޷�����") > 0 Then .Label57.Caption = "���ļ�����": Exit Sub
        If .CheckBox2.Value = True Then '��ѡ
            addcodex = 1
        Else
            addcodex = 2
        End If
        filexpath = .TreeView1.Nodes(k).key
        If ListAllFiles(addcodex, filexpath) = False Then .Label57.Caption = "�����ļ���": Exit Sub
        If InStr(strx, "�����仯") > 0 Then
            .TreeView1.Nodes(k).Text = Left$(strx, Len(strx) - 6) 'ȥ�������仯����ʾ
        ElseIf InStr(strx, "�����") = 0 Then
            .TreeView1.Nodes(k).Text = strx & "(�����)" '�������
        End If
        If .TreeView1.Nodes(k).Children > 0 Then '���������ļ���
            If addcodex = 1 Then
                Call CheckTreeLists(.TreeView1, .TreeView1.Nodes(k)) '�������
                j = ich - 1
                For i = 1 To j                                '��������һ��ֵ����ѡ����
                    strx1 = .TreeView1.Nodes(arrch(i)).Text
                    If InStr(strx1, "�����仯") > 0 Then
                        .TreeView1.Nodes(arrch(i)).Text = Left$(strx1, Len(.TreeView1.Nodes(arrch(i)).Text) - 6) 'ȥ�������е�(�����仯)'.TreeView1.Nodes(arrch(arrch(i))).text
                    ElseIf InStr(strx, "�����") = 0 Then
                        .TreeView1.Nodes(arrch(i)).Text = strx & "(�����)" '�������,���������
                    End If
                Next
            End If
        End If
    End With
    DataUpdate '�������ݸ���
End Sub

Private Sub CommandButton66_Click() '�ļ���-������״ͼ
    With Me
        If Len(.Label80.Caption) = 1 Then
            .ListView2.ListItems.Clear
            With .TreeView1
                If Len(ThisWorkbook.Sheets("������").Range("e37").Value) = 0 Then '�����ݱ���յ�ʱ��
                    .Nodes.Clear
                Else
                    .Nodes.Clear
                    .Nodes.Add , , "Menus", "Menus" '��Ŀ¼
                    .Nodes(1).Expanded = True
                    .Appearance = cc3D
                    .HotTracking = True
                    .Nodes(1).Bold = True
                    .LabelEdit = tvwManual '���ýڵ㲻�ɱ༭
                    Call TreeLists(1)
                End If
            End With
            If Len(.Label240.Caption) > 0 Then
                .Label240.Caption = ""
                .Label241.Caption = ""
            End If
        End If
    End With
End Sub

Function TreeLists(ByVal exc As Byte) '��״չ��
    Dim dic As New Dictionary '�洢ȥ�ص����ļ���
    Dim arr() As String
    Dim arrx() As String
    Dim i As Integer, l As Integer, k As Integer, Elow As Integer
    Dim strx As String, strx1 As String
    
    strx = ThisWorkbook.Sheets("������").Range("e37").Value
    If Len(strx) = 0 Or InStr(strx, "\") = 0 Then Exit Function '�������ļ���
    With Me
        strx1 = .Label80.Caption
        If exc = 0 And strx1 = "temp" Then '���ڿ���
            With .TreeView1
                .Appearance = cc3D
                .HotTracking = True
                .Nodes.Add , , "Menus", "Menus" '��Ŀ¼
                .Nodes(1).Expanded = True '��һ���ڵ�չ����״̬
                .Nodes(1).Bold = True
                .LabelEdit = tvwManual '���ýڵ㲻�ɱ༭
            End With
        End If
        
        If strx1 = "temp" Or exc = 1 Then
            With ThisWorkbook.Sheets("������")
                Elow = .[e65536].End(xlUp).Row
                ReDim arr(1 To Elow - 36)
                For i = 37 To Elow
                    arr(i - 36) = Split(.Range("e" & i), "\")(0) & "\" & Split(.Range("e" & i), "\")(1) '���ϲ��Ŀ¼
                Next
            End With
        
            For k = 1 To UBound(arr)
                If fso.folderexists(arr(k)) = False Then GoTo 10 'У���ļ����Ƿ����
                dic(arr(k)) = ""
10
            Next
            ReDim arrx(0 To UBound(dic.Keys))
            For l = 0 To UBound(dic.Keys)
                arrx(l) = dic.Keys(l)
            Next
            ListFolderx arrx
            If exc = 1 Then Exit Function '����
            .Label80.Caption = 1 '���
        End If
    End With
End Function

Function ListFolderx(ByRef arrt() As String) '��״չ�� '���������byref
    Dim fd As Folder
    Dim i As Integer
    Dim showname As String
    Dim rnglistx As Range
    Dim strp As String
    Dim tracenum As Byte, k As Byte, blow As Integer, j As Byte
    
    k = UBound(arrt())
    ReDim arrlx(1 To 50)
    With ThisWorkbook.Sheets("Ŀ¼")
        blow = .[b65536].End(xlUp).Row
        j = .Cells.SpecialCells(xlCellTypeLastCell).Column
        For i = 0 To k
            s = 1                         'ע���������¹�1���� sΪģ�鼶����,�����������
            tracenum = 0
            Set fd = fso.GetFolder(arrt(i))
            If fd.ParentFolder.Path = Environ("SYSTEMDRIVE") & "\" Then tracenum = 1 'λ��ϵͳ�����ڵĴ���
            arrlx(s) = fd.Path
            showname = fd.Name
            If tracenum = 1 Then showname = showname & "(�޷�����)"
            If tracenum <> 1 Then
                strp = fd.Path & "\"
                Set rnglistx = .Cells(4, 3).Resize(blow, j).Find(strp, lookat:=xlWhole)
                If Not rnglistx Is Nothing Then
                    If fd.DateLastModified <> rnglistx.Offset(0, 2) Then '�ļ��е��޸�ʱ�䷢���仯(��ζ���ļ���(���������е����ļ���)��һ�㷢���仯,�޸�/ɾ��/�½��ļ���/�޸��ļ���)
                        showname = showname & "(�����)(�����仯)"
                    Else
                        showname = showname & "(�����)"
                    End If
                End If
            End If
            
            With Me.TreeView1.Nodes
                .Add "Menus", 4, arrlx(s), showname
            End With
            ListFolderxs fd, tracenum
        Next
    End With
    Erase arrlx
    Set fd = Nothing
End Function

Private Function ListFolderxs(ByVal fd As Folder, ByVal contx As Byte) '�ļ���-��ʾ�б�
    Dim sfd As Folder, i As Long
    Dim showname As String
    Dim rnglistx As Range
    Dim strp As String, strx As String
    
    On Error GoTo 110
10
    If fd.SubFolders.Count = 0 Then Exit Function '���ļ�����ĿΪ�����˳�sub
    For Each sfd In fd.SubFolders
       strx = sfd.Path
       If ErrCode(strx, 1) > 1 Then GoTo 100 '���ư�����ansi�ַ����ļ��� , �����Ҫ��Ҫ�����ansi�ַ���·��,��Ҫ����ʱ���������ڱ���·��
       If contx = 1 Then
          If strx <> Environ("UserProfile") Then GoTo 100    'ֻ��������û��ļ���
       End If
       If contx = 2 Then
'           If sfd.Path <> Environ("UserProfile") & "\Downloads" And sfd.Path <> Environ("UserProfile") & "\Documents" And Environ("UserProfile") & "Desktop" Then GoTo 100 'ֻ��������û��ļ����µ�download��document,desktop�����ļ���
            If CheckFileFrom(strx, 2) = True Then GoTo 100
       End If
20
       i = sfd.Attributes
       If i = 18 Or i = 1046 Then GoTo 100 '������Ҫע��ϵͳ�ļ��е�����,�ܽӷ��� ,����18��ʾ��������,�б����ļ���34��������\1046,�����ļ�����
       '-------------------------���߿��Խ�һ�������ļ�������Ϊ17
       showname = sfd.Name
       If contx = 1 Then showname = showname & "(�޷�����)"
       If contx <> 1 Then
          With ThisWorkbook.Sheets("Ŀ¼")
             strp = strx & "\"
             Set rnglistx = .Cells(4, 3).Resize(.[b65536].End(xlUp).Row, .Cells.SpecialCells(xlCellTypeLastCell).Column).Find(strp, lookat:=xlWhole)
                 If Not rnglistx Is Nothing Then
                     If sfd.DateLastModified <> rnglistx.Offset(0, 2) Then '�ļ��е��޸�ʱ�䷢���仯(��ζ���ļ���(���������е����ļ���)��һ�㷢���仯,�޸�/ɾ��/�½��ļ���/�޸��ļ���)
                         showname = showname & "(�����)(�����仯)"
                         
                     Else
                         showname = showname & "(�����)"
                     End If
                 End If
           End With
        End If
        With Me.TreeView1.Nodes
            If arrlx(s + 1) <> strx Then
                arrlx(s + 1) = strx
                .Add arrlx(s), 4, arrlx(s + 1), showname
            End If
        End With
30
        If sfd.SubFolders.Count > 0 Then s = s + 1
        If contx > 0 Then contx = 2 'ִ��ѭ����һ�׶�
        ListFolderxs sfd, contx
100
    Next
    s = s - 1 '����
    Exit Function
110
If Err.Number = 70 Then Err.Clear: GoTo 100 'ĳЩϵͳ�ļ��л���־ܾ�����Ȩ��70����
End Function

Function CheckTreeLists(ByRef treevw As TreeView, ByRef nodThis As node) '�ļ���-��ʾ�б��ӽڵ�,��ѡ
    Dim lngIndex As Integer

    If nodThis.Children > 0 Then
        lngIndex = nodThis.Child.Index
        Call CheckTreeLists(treevw, treevw.Nodes(lngIndex))
        While lngIndex <> nodThis.Child.LastSibling.Index
          lngIndex = treevw.Nodes(lngIndex).Next.Index
          Call CheckTreeLists(treevw, treevw.Nodes(lngIndex))
        Wend
    End If
    ich = ich + 1
    arrch(ich) = nodThis.Index
End Function

Private Sub TreeView1_Click() '��״ͼ�ڵ�ѡ��
    Dim i As Integer, k As Integer, m As Integer
    Dim fd As Folder, strx3 As String, strx4 As String, fdz As String, fdzx As Long
    Dim strx As String, strx1 As String, strx2 As String
    
    On Error Resume Next
    With Me.TreeView1
        k = .Nodes.Count
        If k = 0 Then Exit Sub
        m = .SelectedItem.Index
        strx = .Nodes(m).key
        strx3 = .Nodes(m).Text
        .SelectedItem.Bold = True
        For i = 2 To k
            If i = m Then GoTo 100
            .Nodes(i).Bold = False '��ѡ�нڵ�,����Ĵ���ȡ��
100
        Next
    End With
    If m = 1 Then Exit Sub
    If InStr(strx3, "�޷�����") > 0 Then Exit Sub
    If Len(storagex) > 0 Then
        If strx = storagex Then Exit Sub
    Else
        storagex = strx '�洢ֵ,����ͬ�������ݷ���������
    End If
    Set fd = fso.GetFolder(strx)
    strx4 = "Size: "
    '---------------------������ĳЩ�����ļ��е�ʱ�����ִ���(��Ȩ��)
    fdz = CStr(fd.Size) ' ת��Ϊ�ı�   http://www.360doc.com/content/18/0613/16/36550511_762117327.shtml
    If Len(fdz) = 0 Then
        fdz = "δ��ȡ������"
    Else
        fdzx = Int(fdz) 'ת��Ϊ����
        If fdzx < 1048576 Then
             fdz = Format(fdzx / 1024, "0.00") & "KB"    '�ļ��ֽڴ���1048576��ʾ"MB",������ʾ"KB"
        Else
            fdz = Format(fdzx / 1048576, "0.00") & "MB"
        End If
    End If
    With Me
        .ListView2.ListItems.Clear
        .Label240.Caption = strx4 & fdz
        .Label241.Caption = fd.DateLastModified
    End With
    Set fd = Nothing
    If FileFilterTemp(strx, 5, , , , , 2) = 0 Then Me.Label57.Caption = "���ļ���Ϊ��"
'   If strx = .Label81.Caption Then Exit Sub '�Ѿ�д������ '����
'   Call FileListv(strx, strx1, strx2, 5, 0, 0, 0)
End Sub

Private Sub CommandButton68_Click() '�ļ���-��ʾ�б�
    Call TreeLists(0)
End Sub

Private Sub CommandButton69_Click() '����Ƽ�-�ļ���
    Dim listnumt As Integer, listnumx As Integer, litm As Variant, randomx As Integer, i As Integer
    
    With ThisWorkbook.Sheets("���")
        listnumt = .[b65536].End(xlUp).Row - 5
        If listnumt < 10 Then
            Me.Label57.Caption = "����̫�ٲ��Ƽ�"
            Exit Sub '����̫�ٲ���ʾ
        End If
        Me.ListView2.ListItems.Clear
        If listnumt < 50 Then '��ʾ�Ƽ�������
            listnumx = 5
        ElseIf listnumt > 49 And listnumt < 150 Then
            listnumx = 10
        ElseIf listnumt > 149 Then
            listnumx = 15
        End If
        ReDim arrTemp(1 To listnumx)
        For listnum = 1 To listnumx             '�������ֹ����ʲ��ֵĲ���arrtemp,listnumΪģ�鼶�����
            randomx = RandomNumx(listnumt)
            arrTemp(listnum) = randomx
            Set litm = Me.ListView2.ListItems.Add()
            i = randomx + 5
            litm.Text = .Range("b" & i)
            litm.SubItems(1) = .Range("c" & i)
            litm.SubItems(2) = .Range("d" & i)
            litm.SubItems(3) = .Range("e" & i)
        Next
    End With
End Sub

Private Sub CommandButton76_Click() '����-�㷨1
    Call Matchx1
    Me.CommandButton76.Enabled = False
    Me.CommandButton77.Enabled = True
End Sub

Private Sub CommandButton77_Click() '����-�㷨2
    Call Matchx2
    MsgBox "Speed Defines The Winner", vbCritical, "Tips"
    Me.CommandButton76.Enabled = True
    Me.CommandButton77.Enabled = False
End Sub

'---------------------------------------------------------------------------����ѵ��
Private Sub Frame11_Click() '����ѵ�����
    Me.TextBox23.SetFocus '����;۽�
End Sub

Private Sub CommandButton71_Click() '����ѵ��-��ʼ
    Timeset = 2 '����һ��ʱ��sub����ֹͣ��״̬
    With Me
        If RecData = False Then .Label57.Caption = "�����쳣": Exit Sub '��鱾���ļ������Ƿ�����
        If Len(.ComboBox10.Value) = 0 Or IsNumeric(.ComboBox10.Value) = False Then Exit Sub
        If .ComboBox10.Value < 5 Or .ComboBox10.Value > 30 Then Exit Sub '��ֹ�����ֶ�������ɵĴ���
        If .Label66.Caption = "play" Then '���ֿؼ������е�״̬
            wm.Controls.Stop
            .Label66.Caption = "stop"
        End If
        .Frame1.Enabled = False
        .Frame3.Enabled = False
        .Frame10.Enabled = False
        .CommandButton71.Enabled = False 'ִ�к����
        .TextBox23.SetFocus
        .ComboBox10.Enabled = False
    End With
    Call Excesub
End Sub

Sub Excesub() '����ѵ��������
Dim sTest As String
Dim i As Byte, Alastrow As Integer, randomx As Integer, chlistnum As Byte, dicl As Byte
Dim time1 As Long, timex As Integer '��¼ʱ��
Dim spx As Integer, sp As Byte, spx1 As Byte, pausetime As Long, tipslen As Byte, tipsx As Byte, errc As Byte
Dim lastpath As String, TableName As String
Dim strx As String

Alastrow = 3041 '���ʱ������
listnum = 0

With Me
    chlistnum = Int(.ComboBox10.Value) 'ȡ����,��ֹ�ֶ��������
    .Label86.Caption = "����:" & chlistnum '���޸�
    .Label88.Caption = "״̬:������" '��ʾ״̬
    .Label94.Caption = "" '�����ȷ��
    .ListView3.ListItems.Clear 'ʹ��֮ǰ����б�����
    
    ReDim arrTemp(1 To chlistnum)
    ReDim arrtemp1(1 To chlistnum) '�������� '��������Ҳ�ܽ�ԭ�е�������������
    ReDim arrtemp2(1 To chlistnum)
    ReDim arrtemp3(1 To chlistnum)
    TableName = "����"
    Set rs = New ADODB.Recordset    '������¼������    '���߰����еĵ�����ȡ�����ŵ�������
    If rs Is Nothing Then MsgBox "�޷���������", vbCritical, "Waring"
    For listnum = 1 To chlistnum
        spx = 1 '��ʾʱ�� 'ע��Ҫ���ò���
        spx1 = 1 '3s����
        errc = 0
        pausetime = 0
        FlagStop = False
        Flagpause = False
        Flagnext = False 'ÿ��ִ��ǰ���ò���
200
        randomx = RandomNumx(Alastrow) '���������ʾ���ɵ�����������ֵ
        arrTemp(listnum) = randomx '��ʱ�洢���� '���ڱȽ������ɵ�������Ƿ�����ص�
        
        SQL = "select * from [" & TableName & "$] where ��� = " & randomx '�Ӵ洢�ļ���ȡ��Ϣ
        rs.Open SQL, Conn, adOpenKeyset, adLockOptimistic
        
        If rs.BOF And rs.EOF Then
            errc = errc + 1       '����û�л�ȡ�ĵ���,���½��л�ȡ,����5�ξ��˳�
            rs.Close
            If errc < 5 Then GoTo 200
            .Label57.Caption = "�쳣,�˳�����"
            Set rs = Nothing
            Exit Sub
        End If
        
        sTest = rs(1) '�洢���� '��-Ӣ��
        arrtemp3(listnum) = rs(2) '����
        arrtemp1(listnum) = sTest
        .Label90.Caption = rs(2)
        .Label97.Caption = rs(3)
        rs.Close
        
        dicl = Len(sTest)
        tipsx = 0
        If .CheckBox6.Value = True Then tipsx = 1
        If dicl < 6 Then                        '���ݵ��ʵĳ���������ʱ��'��ȻҲ�����޸�ʱ�������Ȩ��
            timex = 16000
            If tipsx = 1 Then .Label102.Caption = Left$(sTest, 2) '�����ѡ��ʾ
        ElseIf dicl > 5 And dicl < 8 Then
            timex = 19000
            If tipsx = 1 Then .Label102.Caption = Left$(sTest, 2)
        ElseIf dicl > 7 Then
            timex = 24000
            tipslen = Int(dicl * 0.4) '���ݳ�����������ʾ
            If tipsx = 1 Then .Label102.Caption = Left$(sTest, tipslen)
        End If
        .Label84.Caption = "��ʱ��:" & timex / 1000 & "s" '����ʱ��
        Sleep 100                                    '�ӳ�ִ��
        time1 = timeGetTime '��ʼ��ʱ��
        .Label89.Caption = "������:" & listnum '��������
        If .CheckBox4.Value = True Then
            Speakvs (sTest) '����
            timex = timex + 100 'ʱ�䲹��,���ڷ�����ɵ�ʱ���ӳٻᵼ�����һ�����ʾ�쳣(���ֱ��ʹ��application.speech,��ʱ�ͻ�ǳ�����)
        End If
        
        Do
            DoEvents
            If Flagpause Then
                GoSub 1
                time1 = timeGetTime '���µ�ʱ�俪ʼ���¼���
                pausetime = sp * 1000 '���¿�ʼ��ʱ��
            End If
            sp = Int((timeGetTime - time1 + pausetime) / 1000)
            'С����������� ����i as integer,k as integer �ڽ������,��˵Ȳ���ʱ,��i=100,k=400 msgbox i*k��������ݷ�Χ�����,��Ҫʹ��cnlg���߶�����������Ϊlong����
            If sp = spx Then
                .Label85.Caption = "��ʱ:" & spx & "s" '��ʱ
                spx = spx + 1
            End If
            If sp / 5 = spx1 Then 'ÿ��5s����һ��
            spx1 = spx1 + 1 'ѭ������
                If spx1 < 4 Then
                    timex = timex + 100 'ע�ⲹ���ʱ�䲻Ҫ����1000ms(1s)
                    If .CheckBox4.Value = True And .CheckBox5 = True Then Speakvs (sTest)
                    'ThisWorkbook.Application.Speech.Speak (sTest) 'ע�������applicationҪ��thisworkbook,����ڼ�����淢���л���ʱ�����ִ���
                    DoEvents
                End If
            End If
            If Flagnext Then GoTo 1000 'ע��ִ�е��Ⱥ�˳��
            If FlagStop Then
                Set rs = Nothing
                Conn.Close
                Set Conn = Nothing
                Call QuestionOver
                .Label94.Caption = "����δ���"
                Exit Sub
            End If
            Sleep 25 '�������ѭ���ļ�Ъ,��ֹcpu������ת
        Loop While timeGetTime - time1 + pausetime < timex
1000
        .Label85.Caption = "��ʱ:"
        .Label102.Caption = ""
        strx = Trim(.TextBox23.Text) '��ֹ���ڿո�
        If Len(strx) > 0 Then  '�Զ���ȡ����������
            arrtemp2(listnum) = strx '������
        Else
            arrtemp2(listnum) = "No ANSWER" '��ֵ���߿ո�
        End If
        .TextBox23.Text = "" '�����������
    Next 'ע��������next֮��listnum�����仯

End With

Set rs = Nothing
Call Answerx '����Ľ��
Call QuestionOver '������ɺ�Ĵ���

Exit Sub '�ڵ��ù��̼���������ģ�鼶��ı���ʱ,��Ҫ����ִ�в��ͷű����Ĳ��ǲ���exit sub ����end(����,������ȫ�ͷ��ڴ�)'end ���Խ������еĽ���,��ζ�ſ�������sub�н�����sub,ע�ⲻ���ڴ�����ʹ��
1:
    For i = 0 To 1 Step 0 '�������ͷ����תȦȦ��������Կ���ʹ���Ե��õ�application,now���滻ѭ���ķ���
        Sleep 50
        DoEvents
        If Flagpause = False Then Return '����ִ��
        If FlagStop Then
            Set rs = Nothing
            Call QuestionOver
            Me.Label94.Caption = "����δ���"
            Exit Sub
        End If
    Next
End Sub

Function Speakvs(ByVal strx As String) 'vbs��ʽִ��application.speech,���ֱ��ִ��,�������л���������������Լ1s�ӳ��޷���textbox����������,vbs�ķ�ʽ��ĳ�̶ֳ�ʵ����ν�Ķ��߳�
    Dim WshShell As Object
    Dim vbfilecm As String
    
    If vbsx = 0 Then '����
    With ThisWorkbook.Sheets("temp")
        vbfilecm = .Range("ab17").Value
        If Len(vbfilecm) = 0 Or fso.fileexists(vbfilecm) = False Then
            Me.Label57.Caption = "�����ļ���ʧ"
            Exit Function
        End If
        vbfilex = vbfilecm            '��һ��ʹ�þ͵��ü��
        vbsx = 1
    End With
    End If
    vbfilecm = vbfilex & """" & strx & """"
    Set WshShell = CreateObject("Wscript.Shell")
    WshShell.Run """" & vbfilecm & ""
    Set WshShell = Nothing
End Function

Function RandomNumx(ByVal randomnum As Integer) As Integer '�����
    Dim RndNumber, i As Byte
    
    Randomize (Timer) '��ʼ��rnd
100
    RandomNumx = Int(randomnum * Rnd) + 1
    If listnum > 1 Then
        For i = 1 To listnum
            If RandomNumx = arrTemp(i) Then GoTo 100 '�����ظ��ľ�����ִ��
        Next
    End If
End Function

Function Answerx() '����ѵ��-���Խ��
    Dim i As Byte, k As Byte, j As Byte
    Dim arrtemp4() As String '�洢�жϽ��
    Dim litm As Variant
    
    With Me
        j = .ComboBox10.Value
        ReDim arrtemp4(1 To j)
        .Label94.Caption = ""
        For i = 1 To j
            If arrtemp1(i) = arrtemp2(i) Then
                arrtemp4(i) = "Y"
                k = k + 1 '��ȷ������
            Else
                arrtemp4(i) = "N"
            End If
            Set litm = .ListView3.ListItems.Add()
            litm.Text = arrtemp3(i)
            litm.SubItems(1) = arrtemp1(i)
            litm.SubItems(2) = arrtemp2(i)
            litm.SubItems(3) = arrtemp4(i)
        Next
        If k = 0 Then
            .Label94.Caption = "Fail"
            Exit Function
        End If
        .Label94.Caption = Int(k * 10) & "%" 'int������ʾȡ����,����ȡ,��12.5,��ȡ12 ��������������ȡ13
    End With
    Set litm = Nothing
End Function

Function QuestionOver() '����ѵ��-������Ϻ�Ĵ���
    With Me
        .Frame1.Enabled = True
        .Frame3.Enabled = True
        .Frame10.Enabled = True
        .Label88.Caption = "״̬:����"
        .CommandButton71.Enabled = True
        .ComboBox10.Enabled = True
        .Label85.Caption = "��ʱ:"
        .Label86.Caption = "����:"
        .Label89.Caption = "������:"
        .Label84.Caption = "��ʱ��:"
        .Label90.Caption = ""   '����
        .Label97.Caption = "" '��ʾ
        .Label102.Caption = ""
    End With
    Erase arrTemp '�������
    Erase arrtemp1
    Erase arrtemp2
    Erase arrtemp3
End Function

Private Sub CommandButton72_Click() '����ѵ��-��ͣ/����
    With Me
        If .Label88.Caption = "״̬:����" Or .Label88.Caption = "״̬:" Then Exit Sub
        If .CommandButton72.Caption = "��ͣ" Then
            .CommandButton72.Caption = "����"
            Flagpause = True
        Else
            Flagpause = False
            .CommandButton72.Caption = "��ͣ"
        End If
        .TextBox23.SetFocus
    End With
End Sub

Private Sub CommandButton73_Click() '����ѵ��-ֹͣ
    With Me
        If .Label88.Caption = "״̬:����" Or .Label88.Caption = "״̬:" Then Exit Sub
        If .CommandButton73.Caption = "ֹͣ" Then FlagStop = True
        .TextBox23.SetFocus
        Call QuestionOver
        .Label94.Caption = "����δ���"
    End With
End Sub

Private Sub CommandButton74_Click() '����ѵ��-��һ��
    With Me
        If .Label88.Caption = "״̬:" Or .Label88.Caption = "״̬:����" Then Exit Sub
        Flagnext = True
        .TextBox23.SetFocus
    End With
End Sub
'----------------------------------------------------------------------------------����ѵ��

Private Sub CommandButton8_Click() '�༭-�����Ϣ
    Dim strx As String, strx1 As String, timea As Date
    Dim str As String, TableName As String, str1 As String, str2 As String, str4 As String, str3 As String
    
    With Me
        strx = Trim(.TextBox3.Text)          '��ִ���ַ��ж�   '�漰���޸��ļ�·���Ĳ�����Ҫ�����ַ����ж�
        strx1 = Trim(.TextBox4.Text)
        If Len(strx) > 0 Then  '������ '���ļ���
            If ErrCode(strx, 1) > 1 Then
                .Label57.Caption = "���ļ������ڷ�ANSI�ַ�"
                Exit Sub
            End If
        End If
        If Len(strx1) > 0 Then '����
            If ErrCode(strx1, 1) > 1 Then
                .Label57.Caption = "�������ƴ��ڷ�ANSI�ַ�"
                Exit Sub
            End If
        End If
        
        TableName = "ժҪ��¼"
        str = .Label29.Caption 'ͳһ����
        str1 = .Label23.Caption '�ļ���
        str2 = strx
        str3 = .Label33.Caption '��ʶ
        timea = Now 'ʱ��
        str4 = .TextBox2.Text '����
        If Len(str4) > 1024 Then MsgBox "���ݳ������ȷ�Χ1024", vbInformation, "Tips": Exit Sub
        If RecData = True Then
            SQL = "select * from [" & TableName & "$] where ͳһ����='" & str & "'"                                          '��ѯ����
            Set rs = New ADODB.Recordset
            rs.Open SQL, Conn, adOpenKeyset, adLockOptimistic
            If rs.BOF And rs.EOF Then '�����ж������ҵ�����
                SQL = "Insert into [" & TableName & "$] (ͳһ����,�ļ���,���ļ���,��ʶ����,ʱ��,����) Values ('" & str & "','" & str1 & "', '" & str2 & "','" & str3 & "',#" & timea & "#,'" & str4 & "')"
            Else
                SQL = "UPDATE [" & TableName & "$] SET ����='" & str4 & "',ʱ��=#" & timea & "# WHERE ͳһ����='" & str & "'"
            End If
            rs.Close
            Conn.Execute (SQL)
            Call SearchFile(str)
            If Rng Is Nothing Then .Label57.Caption = "�ļ�������": Exit Sub
            Rng.Offset(0, 19) = .TextBox5.Text '��ǩ1
            Rng.Offset(0, 20) = .TextBox6.Text '��ǩ2
            If Text3Ch <> Trim(.TextBox13.Text) Then Rng.Offset(0, 13) = strx '�����༭��д��
            If Text4Ch <> Trim(.TextBox4.Text) Then Rng.Offset(0, 14) = strx1 '����/���ļ���
            Rng.Offset(0, 16) = .ComboBox4.Text '�ı�����
            Rng.Offset(0, 17) = .ComboBox5.Text '��������
            Rng.Offset(0, 18) = .ComboBox2.Text '�Ƽ�ָ��
            Rng.Offset(0, 31) = .ComboBox12.Text '��������
            .Label57.Caption = "�����ɹ�"
        Else
            .Label57.Caption = "ʧ��,û�����Ӵ洢�ļ�"
        End If
    End With
    Set Rng = Nothing
    Set rs = Nothing
End Sub

Private Sub CommandButton9_Click()          '����-�ʼǹ���-����word�ʼ�
    Dim wdapp As Object
    Dim filen As String, strx As String, filex As String, strx1 As String, strx2 As String
    
    On Error GoTo ErrHandle
    Set wdapp = CreateObject("Word.Application")
    filex = Me.Label29.Caption
    If Len(filex) = 0 Then '����ǿ�,��򿪿հ�word
        With wdapp
            .documents.Add
            .Visible = True
            .Activate
        End With
        Set wdapp = Nothing
        Exit Sub
    End If
    
    With ThisWorkbook.Sheets("temp")
        strx = .Range("ab30").Value
        strx1 = Me.TextBox3.Text
'        strx2 = Format(Now, "yyyymmddhhmmss")
        If Len(strx) = 0 Then  '�����ļ���
           strx = ThisWorkbook.Path & "\note" '��Ŀ¼�´����ʼ��ļ���
           If fso.folderexists(strx) = False Then
              fso.CreateFolder (strx)
              .Range("ab30") = strx
           End If
        Else
           If fso.folderexists(strx) = False Then
              strx = ThisWorkbook.Path & "\note"
              fso.CreateFolder (strx)
              .Range("ab30") = strx
           End If
        End If
    End With
'    filen = strx & "\" & filex & "-" & strx2 & ".docx"   '���-ʱ��
     filen = strx & "\" & filex & ".docx"   '���
    With wdapp
        If fso.fileexists(filen) = False Then
            .documents.Add
            .ActiveDocument.Paragraphs(1).Range.InsertBefore (strx1) '���ļ��ĵ�һ�в������
            .ActiveDocument.SaveAs FileName:=filen
         End If
        .documents.Open (filen)
        .Visible = True
        .Activate                                   '��֮����ʾ�ɼ�/Ϊ��ǰ�Ļ����
    End With
    
ErrHandle:                           '���ִ����ʱ���˳�word,������ܳ���word�ڽ���û�б��˳�������
    If Err.Number <> 0 Then
        wdapp.Quit
        Me.Label57.Caption = Err.Number
        Err.Clear
    End If
    Set wdapp = Nothing
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '����Ķ�-˫��������Ķ��ļ�
    Dim i As Integer, k As Byte
    Dim strx As String
    
    With Me.ListBox1
        If .ListCount = 0 Then Exit Sub
        k = .ListIndex
        strx = .Column(0, k)
        If CheckFileOpen(strx) = True Then Set Rng = Nothing: Exit Sub '�ж��ļ��Ƿ������Ŀ¼���߱��ش���/�ļ��Ƿ��ڴ򿪵�״̬
    End With
    With Rng
    If OpenFile(strx, .Offset(0, 1).Value, .Offset(0, 2).Value, .Offset(0, 3).Value, 1, .Offset(0, 27).Value) = False Then
        Me.Label57.Caption = "�쳣"
        Set Rng = Nothing
        Exit Sub
    End If
    End With
    Call OpenFileOver(strx)
End Sub

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '�����Ķ�-˫�������б�,���ļ�
    Dim n As Byte, m As Byte
    Dim strx As String, strx1 As String
    
    With Me.ListBox2 '�����Ķ��б�
        m = .ListCount
        If m = 0 Then Exit Sub
        n = .ListIndex
        strx = .Column(0, n)
        strx1 = .Column(1, n)
        If CheckFileOpen(strx) = True Then Set Rng = Nothing: Exit Sub '�ж��ļ��Ƿ������Ŀ¼���߱��ش���/�ļ��Ƿ��ڴ򿪵�״̬
    End With
    With Rng
        If OpenFile(strx, .Offset(0, 1).Value, .Offset(0, 2).Value, .Offset(0, 3).Value, 1, .Offset(0, 27).Value) = False Then
            Me.Label57.Caption = "�쳣"
            Set Rng = Nothing
            Exit Sub
        End If
    End With
    Call OpenFileOver(strx)
End Sub

Private Function OpenFileOver(ByVal filecodex As String, Optional cx As Byte) '�ڴ����д��ļ��漰���ƺ����-���´��������
    Dim itemf As ListItem '����listview1��ֵ
    Dim p As Integer, i As Byte
    
    Call RecentUpdate '���»�ȡ��������-�����
    If cx = 1 Then GoTo 100
    With Me.ListView1 '--------------------------------���������������
        If .ListItems.Count = 0 Then GoTo 100
        Set itemf = .FindItem(filecodex, lvwText, , lvwPartial) '��ѯ -�����������е���Ϣ
        If itemf Is Nothing Then
        Set itemf = Nothing
        Else
        p = itemf.Index
        With .ListItems(p) '������������
            If .SubItems(4) = "" Then .SubItems(4) = 0 '��ֵ��"0"�����������
            .SubItems(4) = .SubItems(4) + 1 '�򿪴���+1
            i = Int(.SubItems(4))
        End With
        End If
    End With
100   '----------------------------------���±༭������
    With Me
        If Len(.Label29.Caption) > 0 Then
            If .Label29.Caption = filecodex Then '���±༭ҳ���ϵ���Ϣ
                If Reditx = 1 Then
                    Call FileChange '�������������
                Else
                    If Len(.Label32.Caption) > 0 Then
                        i = Int(.Label32.Caption) + 1
                    Else
                        i = 1
                    End If
                    .Label32.Caption = i
                    .Label31.Caption = Recentfile '���´򿪵�ʱ�� '�򿪴���+1
                End If
            End If
        End If
    End With
    Reditx = 0 '��������������
    Set Rng = Nothing
    Set itemf = Nothing
End Function

Private Sub ListBox3_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '˫��������ļ���,���ļ���
    With Me.ListBox3
        If .ListCount = 0 Then Exit Sub
        Call OpenFileLocation(.Column(0, .ListIndex))
    End With
End Sub

Private Sub ListBox5_Click() '�༭-����������-δ���
    Dim i As Integer
    i = Me.ListBox5.ListIndex
    If Me.CommandButton4.Caption = "��ѯ���" Then
        Me.TextBox1.Text = Me.ListBox5.Column(0, i) & " " & "&" & " " & Me.ListBox5.Column(1, i)
    ElseIf Me.CommandButton4.Caption = "���ʲ�ѯ" Then
        Me.TextBox1.Text = Me.ListBox5.Column(0, i)
    End If
End Sub

Private Sub ListView1_Click()
    With Me.ListView1
        If .ListItems.Count = 0 Then Exit Sub
        .FullRowSelect = True
        .ControlTipText = .SelectedItem.ListSubItems(1).Text
    End With
End Sub

Private Sub ListView1_DblClick() '�������-���б���ֱ��˫�����ļ�
    Dim strx As String, strx1 As String, strx2 As String
    
    With Me.ListView1
        If .ListItems.Count = 0 Then Exit Sub
        strx = .SelectedItem.Text
        strx1 = .SelectedItem.ListSubItems(1).Text
        strx2 = LCase(Right$(strx1, Len(strx1) - InStrRev(strx1, ".")))
        If strx2 Like "xl*" Then
            If strx2 <> "xlsx" Then MsgBox "��ֹ�򿪴����ļ�", vbCritical, "Warning": Exit Sub 'excel���ļ�ֻ����xlsx��ʽ�Ĵ�, ��ֹ�����������ĺ����/��ͻ��������
        End If
        '����excel���ļ��Ĵ�
        If CheckFileOpen(strx) = True Then Set Rng = Nothing: Exit Sub '�ж��ļ��Ƿ������Ŀ¼���߱��ش���/�ļ��Ƿ��ڴ򿪵�״̬
        With Rng
            If OpenFile(strx, .Offset(0, 1).Value, .Offset(0, 2).Value, .Offset(0, 3).Value, 1, .Offset(0, 27).Value) = False Then
                Me.Label57.Caption = "�쳣"
                Set Rng = Nothing
                Exit Sub
            End If
        End With
        '----------�����д��ļ�·��(����)����Ҫʹ�ñ���ϵ�����,��ֹ��ansi
        With .SelectedItem
            If Len(.SubItems(4)) = 0 Then .SubItems(4) = 0 '��ֵ��"0"�����������
            .SubItems(4) = .SubItems(4) + 1 '�򿪴���+1
            Call OpenFileOver(strx, 1)
            End If
        End With
    End With
End Sub

Function ReSetDic() 'ҳ���л�-����ѵ��-����
    With Me
        FlagStop = True 'ֹͣ�������еĽ���
        .Label88.Caption = "״̬:"
        .Label85.Caption = "��ʱ:"
        .Label84.Caption = "��ʱ��"
        .Label90.Caption = ""     '����
        .Label86.Caption = "����:"
        .Label89.Caption = "������:"
        .Label94.Caption = "" '��ȷ��
        .Label97.Caption = "" '��ʾ
        .Frame3.Enabled = True
        .Frame1.Enabled = True
        .Frame10.Enabled = True
        .ComboBox10.Enabled = True
        .ComboBox10.Value = "" '��������
        .CommandButton71.Enabled = True '��ʼ��ť
        .CommandButton72.Caption = "��ͣ" '��ͣ������ť
        .ListView3.ListItems.Clear '��������������
        .Label102.Caption = ""
        .CheckBox4.Value = False
        .CheckBox5.Value = False 'ѡ��ť
        .CheckBox6.Value = False
    End With
End Function

Private Sub MultiPage1_Change()                'ҳ���л�-����
    Dim pgindex As Integer
    Dim arrpa() As Variant, i As Byte, ablow As Integer
    Dim url As String
    
    With Me
        pgindex = .MultiPage1.Value
        If .Label88.Caption <> "״̬:" Then Call ReSetDic 'ֻҪ����ҳ�淢���仯,��ִ������
        If browser1 = 1 Then
            If pgindex <> 7 Then 'ҳ���л�,�Ƴ��������
            With Me!web
                browserkey = .LocationURL
            End With
            Me.Controls.Remove "web"
            browser1 = 0
            End If
        End If
        '----------------------------����
        If pgindex = 1 Then '����
            .TextBox1.SetFocus
            
        ElseIf pgindex = 0 Then '������
            .TextBox8.SetFocus
            
        ElseIf pgindex = 2 And pgx2 = 0 Then '�ļ��� '���ǵ��ļ��϶�ʱ���н���,ȡ���Զ���ʾ�б�
             With .ListView2
                .ColumnHeaders.Add , , "����", 66, lvwColumnLeft
                .ColumnHeaders.Add , , "�ļ���", 116, lvwColumnLeft
                .ColumnHeaders.Add , , "����", 32, lvwColumnLeft
                .ColumnHeaders.Add , , "�ļ�λ��", 221, lvwColumnLeft
                .View = lvwReport                            '�Ա���ĸ�ʽ��ʾ
                .LabelEdit = lvwManual                       'ʹ���ݲ��ɱ༭
                .Gridlines = True
            End With
            pgx2 = 1
            
        ElseIf pgindex = 4 Then '����
            .TextBox18.SetFocus
            .Label109.Caption = ""
            
        ElseIf pgindex = 3 Then '����
            .TextBox14.SetFocus
            If pgx3 = 0 Then 'ע��listview���������ص�״̬�������Ϣ
                With .ListView3
                    .ColumnHeaders.Add , , "����", 120, lvwColumnLeft
                    .ColumnHeaders.Add , , "��", 60, lvwColumnLeft
                    .ColumnHeaders.Add , , "����", 65, lvwColumnLeft
                    .ColumnHeaders.Add , , "Y/N", 40, lvwColumnLeft
                    .View = lvwReport                            '�Ա���ĸ�ʽ��ʾ
                    .LabelEdit = lvwManual                       'ʹ���ݲ��ɱ༭
                    .Gridlines = True
                End With
                '���ʲ�������
                .ComboBox10.List = Array(10, 15, 20)
                pgx3 = 1 '����
            End If
            
        ElseIf pgindex = 7 Then '����
            .TextBox11.SetFocus
            If pgx6 = 0 Then
                With ThisWorkbook.Sheets("temp")
                    ablow = .[aa65536].End(xlUp).Row
                    arrpa = .Range("ab1:ab" & ablow).Value
                End With
                If Len(arrpa(18, 1)) > 0 Then clickx = 1: .CheckBox1.Value = True '��ѯ���ʷ���
                If Len(arrpa(29, 1)) > 0 Then clickx = 1: .CheckBox7.Value = True
                If Len(arrpa(31, 1)) > 0 Then clickx = 1: .CheckBox8.Value = True
                If Len(arrpa(19, 1)) > 0 Then .Label232.Caption = arrpa(19, 1) '��ʾ�ļ��������µ�ʱ��
                If Len(arrpa(27, 1)) > 0 Then .TextBox22.Text = arrpa(27, 1) '�ļ��ı�������
                If Len(arrpa(36, 1)) > 0 Then
                    .CheckBox11.Value = True 'ɾ���ļ��Զ�����md5
                    If Len(arrpa(37, 1)) > 0 Then clickx = 1: .CheckBox12.Value = True
                End If
                If Len(arrpa(38, 1)) > 0 Then
                    If IsNumeric(arrpa(38, 1)) = True Then
                        i = arrpa(38, 1)
                        Select Case i
                            Case 1: clickx = 1: .CheckBox10.Value = True
                            Case 2: clickx = 1: .CheckBox11.Value = True
                            Case 3: clickx = 1: .CheckBox12.Value = True
                        End Select
                    End If
                End If
                If Len(arrpa(50, 1)) > 0 Then clickx = 1: .CheckBox19.Value = True
                If Len(arrpa(43, 1)) > 0 Then clickx = 1: .CheckBox17.Value = True 'pdfˮӡ
                If Len(arrpa(53, 1)) > 0 Then clickx = 1: .CheckBox23.Value = True
                pgx6 = 1
            End If
        ElseIf pgindex = 6 Then
            CreateWebBrowser (browserkey)
        End If
    End With
End Sub

'Private Sub OptionButton1_Click() 'youdao -����
'
'With ThisWorkbook.Sheets("��ҳ")
'If Me.OptionButton1.Value = True Then .Cells(3, 1) = 1
'End With
'
'End Sub

'Private Sub OptionButton2_Click() 'baidu
'
'With ThisWorkbook.Sheets("��ҳ")
'If Me.OptionButton2.Value = True Then .Cells(3, 1) = 2
'End With
'
'End Sub

'Private Sub OptionButton3_Click() 'bing
'With ThisWorkbook.Sheets("��ҳ")
'If Me.OptionButton3.Value = True Then .Cells(3, 1) = 3
'End With
'End Sub

Private Sub OptionButton4_Click() '��ɽ�ʰ�
    With ThisWorkbook.Sheets("��ҳ")
        If Me.OptionButton4.Value = True Then .Cells(3, 1) = 4
    End With
End Sub

Private Sub TextBox1_Change() '����/�༭-������
    Dim arra() As Variant
    Dim arrB() As Variant
    Dim i As Integer, j As Integer, k As Integer
    Dim dic As New Dictionary
    Dim dica As New Dictionary
    Dim dicb As New Dictionary, blow As Integer
    Dim strx As String, m As Integer, n As Integer, p As Integer, xi As Integer, mi As Byte

'If Me.CommandButton4.Caption = "���ʲ�ѯ" Then                                           '���Ȳ�ѯ���صĴʿ� -δ���
'   If Me.TextBox1.text Like "*#*" Then Exit Sub '����������а�������,�˳�
'      If LenB(StrConv(Me.TextBox1, vbFromUnicode)) >= 4 Then 'ת���ַ�(vba�޷�������Ӣ���ַ�)
'        Me.ListBox5.Clear
'        sql = "select * from [" & TableName & "$] Where Ӣ�� like '%" & Me.TextBox1.text & "%'or ���� like '%" & Me.TextBox1.text & "%'or �Զ��� like '%" & Me.TextBox1.text & "%'or ���� like '%" & Me.TextBox1.text & "%'" 'ģ������,�ٷֺű�ʾͨ���"*"
'        Set rs = New ADODB.Recordset    '������¼������
'        rs.Open sql, conn, adOpenKeyset, adLockOptimistic
'        If rs.BOF And rs.EOF Then                         '������صĴʿ�Ϊ��
'           Me.ListBox5.Visible = False
'        Else
'        Me.ListBox5.Visible = True
'        For m = 1 To rs.RecordCount
'           Me.ListBox5.AddItem
'           Me.ListBox5.List(Me.ListBox5.ListCount - 1, 0) = rs(1) 'Ӣ��
'           Me.ListBox5.List(Me.ListBox5.ListCount - 1, 1) = rs(3) '����
'           Me.ListBox5.List(Me.ListBox5.ListCount - 1, 2) = rs(4) '�Զ���
'           Me.ListBox5.List(Me.ListBox5.ListCount - 1, 3) = rs(5) '����
'           rs.MoveNext
'        Next
'        End If
'    Else
'    Me.ListBox5.Clear
'    End If
'End If

'    With ThisWorkbook.Sheets("���") 'ע�������sheet�����ص�����
'        blow = docmx
'        'If Me.CommandButton4.Caption = "��ѯ���" Then
'        arra = .Range("b6:c" & blow).Value
'        arrB = .Range("r6:s" & blow).Value
'    End With
    If docmx < 8 Then
        Me.Label57.Caption = "���ݿ���δ�洢����"
        Exit Sub
    End If
    ArrayLoad
    strx = Me.TextBox1.Value
    strx = Replace(strx, "/", " ") '�滻��"/"����
    With Me.ListBox5
        If Len(strx) >= 2 Then
            .Clear
            p = docmx - 5
            mi = 0
'            strx = strx & "*"
'            For j = 1 To 2 '����ѭ���Ĵ���
                '��ʱ����1, 4��, ���,�ļ�·��(��������Ҫ���ļ���Ҫ����Ϣ),���ڿ������ñ�ǩ1,��ǩ2������
                For k = 1 To p
                    If InStr(1, arrax(k, 1) & "/" & arrax(k, 4), strx, vbTextCompare) > 0 Then
'                    If arrax(k, j) Like strx Then
                    '-------------------------------������instr������д��,��ʼ�ַ���λ��,Դ,�Ƚϵ�ֵ,�Ƚϵķ���;vbtextcompare��ʾ�����ִ�Сд���бȽ�
                        dic(arrax(k, 1)) = arrax(k, 2)
                        dica(arrax(k, 1)) = arrsx(k, 1)
                        dicb(arrax(k, 1)) = arrsx(k, 2)
                        mi = mi + 1
                        If mi > 10 Then GoTo 100
                    End If
                Next
'            Next
100
            xi = mi - 1 'dic.Count
            If xi >= 0 Then 'ע��dict.keys��ĳ��ֵ����ȷд��Ӧ��dict.keys()(i), ������������,��ֱ��ʹ��new dict���г�ʼ��,�Ϳ���ʡ�Ե�ǰ�������
                For m = 0 To xi
                    .AddItem
'                    n = .ListCount - 1         '�����һ��д������
                    .List(m, 0) = dic.Keys(m)
                    .List(m, 1) = dic.Items(m)
                    .List(m, 2) = dica.Items(m)
                    .List(m, 3) = dicb.Items(m)
                Next '-----------------------��m=0����for��ʱ��,m��+1
            End If
            If xi >= 0 Then
                .Visible = True
            Else
                .Visible = False
            End If
        Else
            .Visible = False
        End If
    End With
End Sub

Private Sub TextBox3_Change() '���ļ����޸�
    Me.Label56.Caption = "" '�������Ļ�����Ϣ,����ִ�ж����鼮���ҵ��ж�
End Sub

Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) '��ֹ���ļ�����������ϵͳ��ֹ���ַ�,��Ϊ�������ļ������������޸��ļ���
    Select Case KeyAscii
        Case asc("/"), asc("\"), asc(":"), asc("*"), asc("?"), asc("<"), asc(">"), asc("|")
        Me.Label57.Caption = "��������Ƿ��ַ���""/\ : * ? <> |"
        Me.TextBox3.Text = ""
        KeyAscii = 0
    Case Else
        Me.Label57.Caption = ""
    End Select
End Sub

Private Sub TextBox4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) '����
    Select Case KeyAscii
        Case asc("/"), asc("\"), asc(":"), asc("*"), asc("?"), asc("<"), asc(">"), asc("|")
        Me.Label57.Caption = "��������Ƿ��ַ���""/\ : * ? <> |"
        Me.TextBox3.Text = ""
        KeyAscii = 0
    Case Else
        Me.Label57.Caption = ""
    End Select
End Sub

Private Sub CheckBox24_Click() '�ı�����
    With Me.CheckBox24
        If .Value = True Then
            searchx = 1
        Else
            searchx = 0
        End If
    End With
End Sub

Private Sub CommandButton151_Click() '��ѯ�ı�
    Dim strx As String, strLen As Byte, strx1 As String * 1
    Dim i As Byte, j As Byte, k As Integer, m As Byte, p As Byte, blow As Integer, mi As Byte
    
    With Me
        If searchx = 0 Then .Label57.Caption = "δ��ѡ�ı�����": Exit Sub
        strx = .TextBox8.Text
        strLen = Len(strx)
        If strLen < 2 Then Exit Sub
        If docmx < 8 Then
            .Label57.Caption = "���ݿ���δ�洢����"
            Exit Sub
        End If
        
        For i = 1 To strLen
            strx1 = Mid(strx, i, 1)
            If strx1 Like "[һ-��]" Then
                j = j + 1
            ElseIf strx1 Like "#" Then
                k = k + 1
            ElseIf strx1 Like "[a-zA-Z]" Then
                m = m + 1
            End If
        Next
        If j < 2 And k = 0 And m = 0 Then
            p = 1
        ElseIf j = 0 And k < 4 And m = 0 Then
            p = 1
        ElseIf j = 0 And k = 0 And m < 4 Then
            p = 1
        End If
        If p = 1 Then
            .Label57.Caption = "�������ݳ��Ȳ�����Ҫ��'"
            Exit Sub
        End If
        '-----------------------------------ǰ���ж�
        ArrayLoad
        mi = 0
        blow = docmx - 5
        If .CheckBox26.Value = True Then '��ѡ����ģʽ
            With .ListView1.ListItems
                .Clear
                For k = 1 To blow
                    If arrax(k, 3) = "txt" Then
                        If fso.fileexists(arrax(k, 4)) = True Then
                            If FindTextInFile(arrax(k, 4), strx) > 0 Then
                                With .Add
                                    .Text = arrax(k, 1)
                                    .SubItems(1) = arrax(k, 2)
                                    .SubItems(2) = arrax(k, 3)
                                    .SubItems(3) = arrax(k, 4)
                                    .SubItems(4) = arrbx(k, 1)
                                End With
                                mi = mi + 1
                                If mi > 50 Then Exit For
                            End If
                        End If
                    End If
                Next
            End With
        Else
            If .CheckBox25.Value = True Then '�����ִ�Сд,ִ��vbtext�Ƚ�
                With .ListView1.ListItems
                    .Clear
                    For k = 1 To blow
                        If arrax(k, 3) = "txt" Then
                            If fso.fileexists(arrax(k, 4)) = True Then
                                If CheckFileKeyWord(arrax(k, 4), strx, 1, 0) = True Then 'ִ���ı��Ƚ�,����ļ��ı����ʽ
                                    With .Add
                                        .Text = arrax(k, 1)
                                        .SubItems(1) = arrax(k, 2)
                                        .SubItems(2) = arrax(k, 3)
                                        .SubItems(3) = arrax(k, 4)
                                        .SubItems(4) = arrbx(k, 1)
                                    End With
                                    mi = mi + 1
                                    If mi > 50 Then Exit For
                                End If
                            End If
                        End If
                    Next
                End With
            Else
                With .ListView1.ListItems '���ִ�Сд,�����ƱȽ�,ִ�е��ٶȸ���
                    .Clear
                    For k = 1 To blow
                        If arrax(k, 3) = "txt" Then
                            If fso.fileexists(arrax(k, 4)) = True Then
                                If CheckFileKeyWordC(arrax(k, 4), strx) = True Then
                                    With .Add
                                        .Text = arrax(k, 1)
                                        .SubItems(1) = arrax(k, 2)
                                        .SubItems(2) = arrax(k, 3)
                                        .SubItems(3) = arrax(k, 4)
                                        .SubItems(4) = arrbx(k, 1)
                                    End With
                                    mi = mi + 1
                                    If mi > 50 Then Exit For
                                End If
                            End If
                        End If
                    Next
                End With
            End If
        End If
        If mi = 0 Then .Label57.Caption = "δ�ҵ�"
    End With
End Sub

Private Sub TextBox8_Change()   '�������ı����� '�ɿ��ǸĳɷǶ�̬������,�����ݵ����㹻���ʱ�� '���϶�������ģʽ
    Dim i As Byte, j As Byte, k As Integer, blow As Integer, n As Byte, m As Byte, p As Byte, t As Byte
    Dim dic As New Dictionary
    Dim dica As New Dictionary
    Dim dicb As New Dictionary
    Dim dicc As New Dictionary
    Dim strx As String, strLen As Byte
    Dim strx1 As String, strx2 As String
    Dim mi As Byte
    Dim xi As Variant
    Dim strTemp As String, strtempx As String, strtempx1 As String, chk As Byte
    '��Ҫ����Ӣ��,����,����,�������(���ո�)�����е���
    If searchx = 1 Then Exit Sub 'ִ���ı�����
    If docmx < 8 Then
        Me.Label57.Caption = "���ݿ���δ�洢����"
        Exit Sub
    End If
    strx = Me.TextBox8.Value
    strx2 = Replace(strx, "/", " ") '�滻��"/"����, ����ʹ��"/"��Ϊ���ӷ�
    strLen = Len(strx)        '�����������ĳ���Ϊ38
    With Me.ListView1.ListItems
        If strLen >= 2 Then             '���������ַ���������Ӧ
            .Clear                 '����ҵ����������
            blow = docmx - 5
            ArrayLoad
            mi = 0
            For k = 1 To blow 'ע������ɸѡ��Ŀ��֮������ͬ���ַ�,�����³��ֶ��н����bug,����ʹ���ֵ�ķ��������
                chk = 0
                p = 0: j = 0: n = 0
                strtempx = ""
                strtempx1 = ""
                strTemp = arrax(k, 1) & "/" & arrax(k, 4) '"/"�����
                If InStr(1, strTemp, strx2, vbTextCompare) > 0 Then '������ʽ�������кܴ�ĵ����ռ�'����һ�Ž��ƴʵı�,������Ӣ�ĵ�ĳЩ�������ͬ������'��ƴд����Ĵ��滻����������
                    chk = 1
                Else
                    If strLen >= 3 Then
                        If InStr(strx, Chr(32)) > 0 Then '��������ݴ��ڿո�,���ո�����ݲ𿪽��м��
                            xi = Split(strx, Chr(32)) '�Կո���Ϊ�ָ�
                            i = UBound(xi)
                            For t = 0 To i
                                strtempx = xi(t)
                                If strtempx Like "[һ-��]" Then '��������ֱ�ӽ����ж�
                                    If InStr(1, strTemp, strtempx, vbTextCompare) > 0 Then
                                        p = p + 1
                                        If p >= 2 Then chk = 1: Exit For
                                    End If
                                Else
                                    If Len(strtempx) >= 2 Then
                                        If InStr(1, strTemp, strtempx, vbTextCompare) > 0 Then chk = 1: Exit For  '��ʾ����Ҫ��
                                    End If
                                End If
                            Next
                        Else
                            '--------- like���÷�: https://analystcave.com/vba-like-operator/
                            For i = 1 To strLen '�ж�����������Ƿ�Ϊ��������
                                strx1 = Mid(strx, i, 1)
                                If strx1 Like "[һ-��]" Then       '���� '�ɵ�������,�ɼ��д��� "��˾����" �ɲ��ɵ������� ��/˾/��/��, ���Կ��Բ�����ո�
                                    If InStr(1, strTemp, strx1, vbTextCompare) > 0 Then
                                        p = p + 1
                                        If p >= 2 Then GoTo 98 '���ȴ�������
                                    End If
                                ElseIf strx1 Like "[a-zA-Z]" Then  'Ӣ����ĸ,����Сд '���д���,������ĸ�ĺ������½�
                                    strtempx = strtempx & strx1
                                ElseIf strx1 Like "[0-9]" Then     '���ֻ���ʹ�� "#"����ʾ���ⵥ��0-9���� '���д���,��������û�е��ʲ���Ӱ���
                                    strtempx1 = strtempx1 & strx1
                                End If
                            Next
                            j = Len(strtempx)
                            n = Len(strtempx1)
                            If p = 0 Then
                                If j = 0 Or n = 0 Then GoTo 99
                            Else
                                If strLen - p < 3 Then GoTo 99
                            End If
                            If j >= 3 Then
                                If InStr(1, strTemp, strtempx, vbTextCompare) > 0 Then chk = 1
                            End If
                            If n >= 3 Then
                                If InStr(1, strTemp, strtempx1, vbTextCompare) > 0 Then chk = 1
                            End If
                        End If
                    End If
                End If
                If chk = 1 Then
98
                    dic(arrax(k, 1)) = arrax(k, 2)
                    dica(arrax(k, 1)) = arrax(k, 3)
                    dicb(arrax(k, 1)) = arrax(k, 4)
                    dicc(arrax(k, 1)) = arrbx(k, 1)  '��һforѭ�����Բ���Ҫ�ֵ�,ֱ��ʹ�����鼴��/����ֱ�����ֵ, ���forѭ����ʱ�����Ҫ
                    mi = mi + 1
                    If mi > 50 Then GoTo 100 '�����������������
                End If
99
            Next
100
            '----------------------------------------------------------------����д��listview
            If mi > 0 Then '��ȡ����Ч��ֵ
                mi = mi - 1
                For m = 0 To mi
                    With .Add
                        .Text = dic.Keys(m)
                        .SubItems(1) = dic.Items(m)
                        .SubItems(2) = dica.Items(m)
                        .SubItems(3) = dicb.Items(m)
                        .SubItems(4) = dicc.Items(m)
                    End With
                Next
            Else
                .Clear
            End If
        Else
            .Clear                   '���������������2λʱ,��������
        End If
    End With
End Sub

Private Sub UserForm_Activate() '��Ծ����
    Dim blow As Integer
    Dim strx As String, strx2 As String
    
    If Statisticsx = 1 Then Exit Sub
    If Workbooks.Count = 1 Then 'ֻ��һ����������ʱ����ʾ��С������
        AddIcon    '���ͼ��
        AddMinimiseButton   '��Ӱ�ť
        AppTasklist Me    '���������
    End If
    If UF3Show = 3 Then ''�������ķ�ʽ,��������״̬, ���������ݱ仯,�ͷ��ͱ仯, Ȼ���ڴ�����״̬�ָ�ʱ��������
        UF3Show = 1
        Call PauseRm '����selection�¼�
        If DeleFilex = 1 Then '�ļ���ɾ��
            strx = Me.Label29.Caption
            If Len(strx) > 0 Then '��ֵ
                SearchFile strx
                If Rng Is Nothing Then
                    DeleFileOverx strx
                    Exit Sub
                End If
            End If
        End If
        DataUpdate
    End If
'    Call EveSpy '���¼�������������״̬
'    Call RecData '��֤ado��������״̬
'    Call LockSet '���ֱ�����������vba��д��״̬
End Sub

Sub DataUpdate() '���´��������, ʹ�����,��Ҫ����������
    If AddPlistx = 1 Then Call PrReadList: AddPlistx = 0 '�����ص�״̬��,�����Ķ��б����仯
    If OpenFilex = 1 Then Call RecentUpdate: OpenFilex = 0 '����Ķ�
    If MDeleFilex = 1 Then Call AddFileListx: MDeleFilex = 0 '��ӵ��ļ��� '��ӵ��ļ��б��Ƴ� '���ݿ��ڵ����ݷ����˱仯
    If DeleFilex = 1 Then Call CwUpdate: Call Choicex: DeleFilex = 0 '�������ݿ� '����� 'ɸѡ����
    docmx = ThisWorkbook.Sheets("���").[d65536].End(xlUp).Row '�ؼ�����
End Sub

Sub ArrayLoad() '�����ֵ����ݼ��ص�����,�ӿ���Ӧ���ٶ�
    If spyx <> docmx Then '����ģ�鼶�����������������ֵ���ڴ���,���ٷ��ʱ�����Ҫ,ֻ�е��������ݷ����仯�����»�ȡֵ,�ӿ���ʵ��ٶ�
        spyx = docmx '-----------��ʼ��ֵ/�仯�ڽ��и�ֵ ' '���ֵ�ɴ����Ծ���Զ���ȡ '��ִ�������ݸ�����Ҫ������θ������ֵ
        If SafeArrayGetDim(arrax) <> 0 Then '�ж������Ƿ񾭹���ʼ��, ��������˳�ʼ��,�ͽ�ԭ�е�����Ĩ��
            Erase arrax
            Erase arrbx                        '�����ݷ����仯��ʱ��,Ĩ���ɵ��������»�ȡ�µ�
            Erase arrsx
        End If
        With ThisWorkbook.Sheets("���")
            arrax = .Range("b6:f" & docmx).Value '���,�ļ���, ��չ��, �ļ�·��, �ļ�����λ��
            arrbx = .Range("n6:n" & docmx).Value '�򿪴���
            arrsx = .Range("s6:t" & docmx).Value '����/�Ƽ�ָ��
        End With
'            arrux = .Range("u6:v" & docmx).Value '��ǩ1/��ǩ2
    End If
End Sub

Private Sub UserForm_Initialize() '�����ʼ��
    Dim dic As New Dictionary
'    Dim arr() As String
    Dim strx As String
    Dim i As Byte, TableName As String, k As Integer
    
    If Statisticsx = 1 Then Exit Sub
    UF3Show = 1 '��Ǵ��崦�ڼ����״̬
    NewM = False '���ڿ���textbox�˵�
    With Me
        .MultiPage1.Value = 0 '�򿪴���ʱ��ʾ������
        '��ʼ��listview
        With .ListView1  'ע��listview�����ڿɼ���״̬����ɳ�ʼ��
            .ColumnHeaders.Add , , "����", 66, lvwColumnLeft
            .ColumnHeaders.Add , , "�ļ���", 122, lvwColumnLeft
            .ColumnHeaders.Add , , "����", 33, lvwColumnLeft
            .ColumnHeaders.Add , , "�ļ�·��", 225, lvwColumnLeft '��Ҫ����һ���Ŀռ�,���������ºỬ����
            .ColumnHeaders.Add , , "�򿪴���", 50, lvwColumnLeft
            .View = lvwReport                            '�Ա���ĸ�ʽ��ʾ
            .LabelEdit = lvwManual                       'ʹ���ݲ��ɱ༭
            .Gridlines = True
        End With
        
        With .ListBox3 '�ļ���
            .MultiSelect = fmMultiSelectMulti '��ѡ
            .ListStyle = fmListStyleOption
        End With
        '��������
        .ComboBox11.List = Array("�����", "Axure", "Mind", "Note", "PDF", "��ͼ", "Spy++", "����", "��ѹĿ¼")
        '��������
        .ComboBox5.List = Array(1, 2, 3, 4, 5)
        'pdf������
        .ComboBox3.List = Array(1, 2, 3)
        '�ı�����
        .ComboBox4.List = Array(1, 2, 3)
        '�ļ�����
        .ComboBox6.List = Array("��", "ɾ��", "��λ��", "�����ļ�", "��ӵ������б�")
        '�Ƽ�ָ��
        .ComboBox2.List = Array(1, 2, 3)
        '��������
        .ComboBox12.List = Array("CNS", "CNT", "EN", "JPN", "OTS", "MIX") '��������,����,Ӣ��,����,����,�������
        .ComboBox13.List = Array(6, 12, 18)
        .ComboBox14.List = Array("Start", "Reading", "Over") '�Ķ�״̬
    End With
    
    PauseRm '����selection�¼�,worksheetchange�¼������״���,�ڴ�����ʾʱ,���õ�
    
    With ThisWorkbook
        docmx = .Sheets("���").[d65536].End(xlUp).Row
        voicex = .Sheets("temp").Range("ab18").Value
        Choicex      'ɸѡ����
        RecentUpdate '����Ķ�
        PrReadList   '�����Ķ�
        AddFileListx '��ӵ��ļ���
        CwUpdate     '�������
'        If .Sheets("temp").Range("ab18") = 1 Then Me.CommandButton44.Visible = True '������ť-δ���
    End With
    Exit Sub
    '��ʼ������¼
'    If RecData = True Then
'        TableName = "����¼"
'        Sql = "select * from [" & TableName & "$]"
'        Set Rs = New ADODB.Recordset    '������¼������
'        Rs.Open Sql, Conn, adOpenKeyset, adLockOptimistic
'        If Rs.BOF And Rs.EOF Then             '�����ж������ҵ�����
'            Me.TextBox10.Enabled = False
'        Else
'            k = Rs.RecordCount
''            ReDim arr(1 To k)
'            For i = 1 To k
'                dic(CStr(Rs.Fields(0))) = ""      '�ֵ�ķ�ʽ�洢����
''                arr(i) = Rs(2)                   '����ķ�ʽ�洢����(ע��2,�Ǻ��������,������һ������)
'                If i = k Then strx = Rs(2)
'                Rs.MoveNext                       '���ݼ�ָ��,ָ����һ������
'            Next
'            With Me
''                .TextBox10.Text = arr(k)
'                .TextBox10.Text = strx
'                .ComboBox9.List = dic.Keys
'                .ComboBox9.Text = dic.Keys(UBound(dic.Keys)) 'ȡ����ֵ
'            End With
'        End If
'        Rs.Close
'        Set Rs = Nothing
'    End If
End Sub

Private Sub Choicex() 'ɸѡ����
    Dim dicop As New Dictionary
    Dim num As Integer, n As Integer
    Dim arrop As Variant
    
    With ThisWorkbook.Sheets("���")
        n = docmx
        If n > 6 Then
            arrop = .Range("d6:d" & n).Value
            dicop.CompareMode = TextCompare '�����ִ�Сд
            n = n - 5
            For num = 1 To n
                dicop(arrop(num, 1)) = "" '����ֵ��key(���ص�����)
            Next
        ElseIf n = 6 Then
            dicop(.Range("d6").Value) = ""
        ElseIf n < 6 Then
            Exit Sub
        End If
    End With
    With Me
        .ComboBox1.List = dicop.Keys
        .ComboBox7.List = Array(1, 2, 3, 4, 5)
        .ComboBox8.List = Array(1, 2, 3)
    End With
End Sub

Private Function RecentUpdate() ''����Ķ�(�޸�)
    Dim k As Byte, i As Byte
    
    With ThisWorkbook.Sheets("������")
        '����Ķ�
        If Len(.Range("p27").Value) > 0 Then
            Me.ListBox1.Clear
            For k = 27 To 33
                If Len(.Range("p" & k).Value) > 0 Then
                   Me.ListBox1.AddItem
                   i = Me.ListBox1.ListCount - 1
                   Me.ListBox1.List(i, 0) = .Range("u" & k)
                   Me.ListBox1.List(i, 1) = .Range("p" & k)
                   Me.ListBox1.List(i, 2) = .Range("w" & k)
                Else
                   Exit For
                End If
            Next
        End If
    End With
End Function

Private Function PrReadList()  '�����Ķ�
    Dim m As Byte, i As Byte
    
    With ThisWorkbook.Sheets("������")
        '�����Ķ�
        If Len(.Range("i27").Value) > 0 Then                      '����ӿհ׵�ֵ����
            Me.ListBox2.Clear
            For m = 27 To 33
               If Len(.Range("d" & m).Value) > 0 Then
                    Me.ListBox2.AddItem
                    i = Me.ListBox2.ListCount - 1
                    Me.ListBox2.List(i, 0) = .Range("i" & m).Value
                    Me.ListBox2.List(i, 1) = .Range("d" & m).Value
                    Me.ListBox2.List(i, 2) = .Range("k" & m).Value
                Else
                    Exit For
               End If
            Next
        End If
    End With
End Function
 
Private Function AddFileListx() '��ӵ��ļ���
    Dim j As Byte, Elow As Byte, i As Byte, m As Byte
    Dim strx As String
    
    With ThisWorkbook.Sheets("������")
        If Len(.Range("e37").Value) > 0 Then '����ӿհ׵�ֵ����
            Elow = .[e65536].End(xlUp).Row
            m = Elow - 37
            ReDim arraddfolder(m) '--------------������ʱ�洢,��ֹ��ansi�ַ�
            Me.ListBox3.Clear '���֮ǰ������
            For j = 37 To Elow
                If Len(.Range("e" & j).Value) > 0 Then
                    Me.ListBox3.AddItem
                    i = Me.ListBox3.ListCount - 1
                    strx = .Range("e" & j).Value
                    Me.ListBox3.List(i, 0) = strx
                    arraddfolder(i) = strx
                    If fso.folderexists(strx) = False Then
                        Me.ListBox3.List(i, 1) = "���ļ������Ƴ�" '����ļ����Ƿ��Ѿ����Ƴ�
                    Else
                        Me.ListBox3.List(i, 1) = .Range("i" & j).Value
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End With
End Function

Private Function CwUpdate() '�������ݸ���
    Dim str1 As Integer
    Dim str2 As String 'ע������
    Dim str3 As Integer
    Dim str4 As Integer
    Dim str5 As Integer
    Dim str6 As Integer
    Dim str7 As Integer
    Dim str8 As Integer

    With ThisWorkbook.Sheets("������")
        str1 = .Range("p37").Value '�ļ�����
        str2 = .Range("p38").Value '�����ļ���С
        str3 = .Range("p40").Value 'pdf
        str4 = .Range("s40").Value 'EPUB
        str5 = .Range("p42").Value '����
        str6 = .Range("p41").Value 'PPT
        str7 = .Range("v41").Value 'Word
        str8 = .Range("s41").Value 'Excel
    End With
    With Me
        .Label47.Caption = str1
        .Label48.Caption = str2
        .Label49.Caption = str3
        .Label50.Caption = str4
        .Label51.Caption = str5
        .Label52.Caption = str6
        .Label53.Caption = str7
        .Label54.Caption = str8
    End With
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) '���ش��ں���ʾuf4
    'https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/queryclose-constants
    If Statisticsx = 1 Then Exit Sub
    UF3Show = 3 '���ڱ�ʾ���崦�����ص�״̬
    If CloseMode = vbFormControlMenu Then                                    '�����޸Ľ�ֹʹ��"x"��ť���رմ���
            Cancel = True
            Me.Hide
            If Workbooks.Count = 1 And Application.Visible = False Then
                UserForm4.Hide
                UserForm4.Show 1
                UserForm4.Caption = "����"
            Else
                UserForm4.Caption = "Mini"
                UserForm4.Show
            End If
    End If
    Timeset = 2
    If FlagStop = False Then FlagStop = True 'ȷ������ѵ�����еļ�ʱ����ֹͣ��״̬,��ʹ��ʱ����صĴ������Ҫ����ʱ��ֹͣ������
End Sub

Private Sub UserForm_Terminate()
    If Statisticsx = 1 Then Exit Sub
    UF3Show = 0
End Sub
