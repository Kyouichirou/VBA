Attribute VB_Name = "��Ƶ��"
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal Index As Long) As Long
Private Const HORZRES = 8
Private Const VERTRES = 10
Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90

Function CreateWorksheet(ByVal dfilepath As String) '�������ڴ洢��Ϣ�ı��
    Dim wb As Workbook
    Dim dicpath As String
    Dim wbdc As Workbook
    
    Set wb = Workbooks.Add
    With wb.Worksheets
        .Add after:=wb.Worksheets(.Count), Count:=6 - .Count '����6�ű�
    End With
    With wb
        .Worksheets(1).Name = "�򿪼�¼"                                                                '���зֱ�д���ͷ
        .Worksheets(1).Range("a1:f1") = Array("ͳһ����", "�ļ���", "���ļ���", "��ʶ����", "ʱ��", "����")
        .Worksheets(2).Name = "ժҪ��¼"
        .Worksheets(2).Range("a1:f1") = Array("ͳһ����", "�ļ���", "���ļ���", "��ʶ����", "ʱ��", "����")
        .Worksheets(3).Name = "����¼"
        .Worksheets(3).Range("a1:c1") = Array("����", "ʱ��", "����")
        With .Worksheets(4)
            .Name = "ɾ������"
            ThisWorkbook.Sheets("���").Range("b5:ag5").Copy .Cells(1, 1)
            .Cells(1, 33) = "ɾ��ԭ��"
            .Cells(1, 34) = "ɾ����ע"
        End With
        .Worksheets(5).Name = "�ʿ�"
        .Worksheets(5).Range("a1:m1") = Array("���", "Ӣ��", "����", "����", "�Զ���", "����", "����", "��ѯ����", "��Ҫ�̶�", "���ʱ��", "��Դ", "�ο���ϢԴ", "���ʱ�")
        .Worksheets(6).Name = "����"
    End With
    wb.SaveAs dfilepath '·�����з�ansi�ַ����?
    dicpath = ThisWorkbook.Path & "\���ʱ�.xlsx"
    If fso.fileexists(dicpath) = True Then '���ʱ���
        Set wbdc = Workbooks.Open(dicpath)
        Workbooks("���ʱ�.xlsx").Sheets("�ʻ��").Cells.Copy Workbooks("lbrecord.xlsx").Sheets("����").Range("a1") '�����ʱ�����ݸ��Ƶ�����
        wbdc.Close True
        Set wbdc = Nothing
    End If
    wb.Close savechanges:=True
    Set wb = Nothing
End Function

Sub AtClock() '��̬ʱ�� '�������ڴ�����ִ�ж��ʱ��sub-����
'If timest = 1 Then
''DoEvents
'UserForm3.TextBox9.text = Format(Now, "yyyy-mm-dd HH:MM:SS")
'Application.OnTime Now + TimeValue("00:00:01"), "Atclock" '���ҵ��ÿ��Ա������doѭ����ɵ�cpuռ������
'End If
End Sub

Sub Deletempfile() '�Ƴ�����Ҫ�ļ�-����
'Dim arr()
'Dim i As Integer, k As Integer
'With ThisWorkbook.Sheets("���")
'If .Range("b6") = "" Then Exit Sub
'If .[b65536].End(xlUp).Row > 6 Then
'arr = .Range("c6:d" & .[b65536].End(xlUp).Row).Value
'Else
'ReDim arr(1 To 1, 1 To 2)
'arr(1, 1) = .Range("c6").Value
'arr(1, 2) = .Range("d6").Value
'End If
'k = .[b65536].End(xlUp).Row - 5
'For i = k To 1 Step -1
'If Not UCase(arr(i, 2)) Like "EPUB" And Not UCase(arr(i, 2)) Like "PDF" And Not UCase(arr(i, 2)) Like "MOBI" And Not UCase(arr(i, 2)) Like "DO*" And Not UCase(arr(i, 2)) Like "XL*" And Not UCase(arr(i, 2)) Like "PP*" And Not UCase(arr(i, 2)) Like "AC*" And Not UCase(arr(i, 2)) Like "TX*" Then
'.Rows(i + 5).Delete Shift:=xlShiftUp
'Else
'   If arr(i, 1) Like "~$*" Then .Rows(i + 5).Delete Shift:=xlShiftUp
'End If
'Next
'End With
End Sub

Sub CheckAllFile() '����ļ��Ĵ���                 'ȫ��ִ���ж�Ŀ¼�µ��ļ��Ƿ����
    Dim arre() As String
    Dim Elow As Integer, i As Integer
    
    With ThisWorkbook.Sheets("���") '�����漰�����鸳ֵʱ��ֻ��һ��ֵ������
        Elow = .[e65536].End(xlUp).Row
        If Elow < 6 Then
        .Label1.Caption = "������"
        Exit Sub      '����Ϊ��
        End If
        If Elow > 100 Then UserForm6.Show 0
        Call PauseRm '����selection�¼�
        If Elow > 6 Then
            arre = .Range("e6:e" & Elow).Value
            For i = 1 To Elow - 5
                If fso.fileexists(arre(i, 1)) = False Then
                    Call Delefile(arre(i, 1), i + 5, 2)
                End If
            Next
        ElseIf Elow = 6 Then
            If fso.fileexists(.Range("e6").Value) = False Then Call Delefile(.Range("e6").Value, 6, 2)
        End If
        Unload UserForm6
        .Label1.Caption = "ִ�����"
    End With
    Call EnableRm '�����Ҽ��¼�
    ThisWorkbook.Save
End Sub

Function PSexist() As Boolean '�ж�powershell �Ƿ���� '��չһ��
    If ShellxExist = 1 Then PSexist = True: Exit Function
    If Len(ThisWorkbook.Sheets("temp").Range("ab4").Value) > 0 Then
        PSexist = True
    Else
        PSexist = False
    End If
End Function

Function CreateFolder(ByVal Folderpath As String, ByVal cmCode As Byte) As Boolean '�����ļ���/�ĵ�Ŀ¼
    Dim i As Byte, xi As Variant, k As Byte, yesno As Variant, wt As Integer, strx As String, strx1 As String, m As Byte, j As Byte, n As Byte
    Dim strx2 As String
    '���������Ǹ��ݱ��е����ݽ��д����ļ��е�,����޸ı��е�����,���µ�����Ҳ��Ҫ�޸�
    CreateFolder = True
    xi = Split(Folderpath, "\")
    i = UBound(xi)
    If Len(xi(i)) > 0 Then Folderpath = Folderpath & "\" '�Ǹ�Ŀ¼
    If cmCode = 1 Then
        strx = "Library"
        strx1 = "*[a-zA-Z]*"
        m = 3
    Else
        strx = "����"
        strx1 = "*[һ-��]*"
        m = 2
    End If
    Folderpath = Folderpath & strx
    
    If fso.folderexists(Folderpath) = True Then
        yesno = MsgBox("�ļ����Ѵ���,�Ƿ����´���(ԭ�ļ�����ɾ��!)", vbYesNo, "Warning")
        If yesno = vbNo Then
            CreateFolder = False
            Exit Function
        Else
            wt = 200
            If fso.GetFolder(Folderpath).Size > 1048576000 Then wt = 350
            fso.DeleteFolder (Folderpath)
            Sleep wt
        End If
    End If
    fso.CreateFolder (Folderpath)
    Folderpath = Folderpath & "\"
    strx2 = Folderpath
    With ThisWorkbook.Sheets("����")
        k = .[a65536].End(xlUp).Row
        j = 3
        n = j
100
            strx = .Cells(n, m).Value
            If strx Like strx1 Then
                fso.CreateFolder (strx2 & strx)
                Folderpath = strx2 & strx & "\"
            End If
            If n = k Then Exit Function
        For j = n To k
            strx = .Cells(j + 1, m + 1).Value
            If strx Like strx1 Then
                fso.CreateFolder (Folderpath & strx)
            Else
                n = j + 1
                GoTo 100
            End If
        Next
    End With
End Function

Sub ScreenDetail()
    Dim hDC As Long
    Dim x As Long, y As Long
    Dim x1 As Long, Y1 As Long
    hDC = GetDC(0)
    x = GetDeviceCaps(hDC, HORZRES)
    y = GetDeviceCaps(hDC, VERTRES)
    MsgBox "��ǰϵͳ����Ļ�ֱ���Ϊ" & x & "X" & y
    x1 = GetDeviceCaps(hDC, LOGPIXELSX)
    Y1 = GetDeviceCaps(hDC, LOGPIXELSY)
    MsgBox "��ǰ��ʾ����PPIΪ" & x1 & "X" & Y1
    ReleaseDC 0, hDC
End Sub

Sub ScreenDetaila() 'wps֧��
    Dim WMIObject As Object
    Dim WMIResult As Object
    Dim WMIItem As Object
    Set WMIObject = GetObject("winmgmts:\\.\root\WMI")
    Set WMIResult = WMIObject.ExecQuery("Select * From WmiMonitorBasicDisplayParams")
    Dim Diagonal As Double
    Dim Width As Double
    Dim Height As Double
    Dim Counter As Byte
    Counter = 1
    For Each WMIItem In WMIResult
        Width = WMIItem.MaxHorizontalImageSize / 2.54
        Height = WMIItem.MaxVerticalImageSize / 2.54
        Diagonal = Sqr((Height ^ 2) + (Width ^ 2))
        MsgBox "Your monitor # " & Counter & " is approximiately " & Round(Diagonal, 2) & " inches diagonal"
        Counter = Counter + 1
    Next
    Set WMIObject = Nothing
    Set WMIResult = Nothing
End Sub

