Attribute VB_Name = "�ı�����"
Option Explicit
'��Ҫϸ��
' �޳������õĴʻ�
' �޳��ص�(aa)
' �޳�������----
' �޳�����һ�����ȷ�Χ,����û��"-"���ӵ��ַ���
' �޳���������(ѧ����רҵ��������, ���������, ��ѧ,���ݸ����������γɵĳ����ʻ�)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, _
ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
'-----------------------------------------https://fishc.com.cn/thread-70452-1-1.html
Private Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, ByVal lpFileMappigAttributes As Long, _
ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, _
ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
'---------------------------------------https://docs.microsoft.com/en-us/windows/win32/api/memoryapi/nf-memoryapi-mapviewoffile
'---------------------------------------https://docs.microsoft.com/en-us/windows/win32/api/memoryapi/nf-memoryapi-mapviewoffileex
Private Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress As Any) As Long
Private Const PAGE_READWRITE = &H4
Private Const FILE_MAP_READ = &H4
'--------------------------------------http://binaryworld.net/Main/ApiDetail.aspx?ApiId=5817
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
'''----------------------------------------------------
Private Type SAFEARRAY1D
    cDims As Integer      '�����ά��
    fFeatures As Integer  '��������������η������α��ͷŵı�־
    cbElements As Long    '����Ԫ�صĴ�С
    clocks As Long        'һ�����������������ٸ����鱻�����Ĵ���
    pvData As Long        'ָ�����ݻ����ָ��--------------------�ؼ�����
    rgsabound(0) As SAFEARRAYBOUND '������ÿά������ṹ��������Ĵ�С�ǿɱ��rgsabound��һ����Ȥ�ĳ�Ա�����Ľṹ��ֱ̫�ۡ�
                                   '�������ݷ�Χ�����顣������Ĵ�С��safearrayά���Ĳ�ͬ����������
                                   'rgsabound��Ա��һ��SAFEARRAYBOUND�ṹ������--ÿ��Ԫ�ش���SAFEARRAY��һ��ά��

End Type
'--------------------https://docs.microsoft.com/en-us/previous-versions/windows/embedded/ms912047(v=winembedded.10)
Public Enum LocalID
    zh_hk = 3076
    zh_ma = 5124
    zh_cn = 2052
    zh_sg = 4100
    zh_tw = 1028
    japan = 1041
    korea = 1042
    en_us = 1033
    en_uk = 2053
End Enum
'----------------codepage��localid�漰��vba��֧�ֵ��������
'vba�����ַ�����Unicode����ʽ����, 2���ֽ�
'vba���õĺ���������ansi�汾��api��ʵ�ֹ���
'lenb���ٶȱ�len��, ��Ϊlen������������ݻ���Ҫ /2
'---------------------------------------------------
Public Enum CodePage '����ҳ, ������URLCodePage����(winhttprequest)
    GB2312 = 20936   '��������
    GBK = 936        '������չ
    Big5 = 950       '��������/��̨
    GB18030 = 54936  '��չ��,�������󲿷ֺ���(��)
    Shift_Jis = 932  '����
    Ks_c_5601 = 949  '����
    IBM437_us = 437  'Ӣ��(US)
    UTF8 = 65001     'UTF-8
End Enum

'----------------------------------------------------------adodb.stream
Private Enum mReadText 'ado.stream��ȡ�ı��ķ�ʽ
    adReadAll = -1
    adReadLine = -2
End Enum
'ָʾֻ��Ȩ��
'ָʾ��/дȨ�ޡ�
'������ *ShareDeny* ֵ��adModeShareDenyNone��adModeShareDenyWrite �� adModeShareDenyRead��һ��ʹ�ã��Խ��������ƴ�������ǰ Record �������Ӽ�¼��
'��� Record û���Ӽ�¼����û��Ӱ�졣��������� adModeShareDenyNone һ��ʹ�ã�����������ʱ����
'���ǣ�������ֵ��Ϻ������Ժ� adModeShareDenyNone һ��ʹ�á����磬����ʹ�á�adModeRead Or adModeShareDenyNone Or adModeRecursive����
'�������������κ�Ȩ�޴�����?���ܾ������˵Ķ����ʻ�д����
'��ֹ�������Զ�Ȩ�޴�����
'��ֹ��������дȨ�޴�����
'��ֹ�����˴�����
'Ĭ��ֵ?ָʾ��δ���û���ȷ��Ȩ��
'ָʾֻдȨ��
Private Enum adConnectMode
    adModeRead = 1
    adModeReadWrite = 3        '����
    adModeRecursive = &H400000
    adModeShareDenyNone = 16
    adModeShareDenyRead = 4
    adModeShareDenyWrite = 8
    adModeShareExclusive = 12
    adModeUnknown = 0
    adModeWrite = 2
End Enum

Private Enum adStreamType 'ָ����������
    adTypeBinary = 1     '������
    adTypeText = 2       '�ı�
End Enum
'-------------------------------------------------------------------adodb.stream

'ʹ���ڴ�ӳ�䷽ʽ���Ҵ����ļ��а������ַ���
Function FindTextInFile(ByVal strFileName As String, ByVal strText As String) As Long
    Dim hFile As Long, hFileMap As Long
    Dim nFileSize As Long, lpszFileText As Long, pbFileText() As Byte
    Dim ppSA As Long, pSA As Long
    Dim tagNewSA As SAFEARRAY1D, tagOldSA As SAFEARRAY1D
  
    hFile = CreateFile(strFileName, _
            GENERIC_READ Or GENERIC_WRITE, _
            FILE_SHARE_READ Or FILE_SHARE_WRITE, _
            0, _
            OPEN_EXISTING, _
            FILE_ATTRIBUTE_NORMAL Or FILE_ATTRIBUTE_ARCHIVE Or FILE_ATTRIBUTE_READONLY Or _
            FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_SYSTEM, _
            0) '���ļ�
    If hFile <> 0 Then
        nFileSize = GetFileSize(hFile, ByVal 0&) '����ļ���С
        hFileMap = CreateFileMapping(hFile, 0, PAGE_READWRITE, 0, 0, vbNullString) '�����ļ�ӳ�����
        lpszFileText = MapViewOfFile(hFileMap, FILE_MAP_READ, 0, 0, 0) '��ӳ�����ӳ�䵽�����ڲ��ĵ�ַ�ռ�
          
        ReDim pbFileText(0) '��ʼ������
        ppSA = VarPtrArray(pbFileText) '���ָ��SAFEARRAY��ָ���ָ��
        CopyMemory pSA, ByVal ppSA, 4 '���ָ��SAFEARRAY��ָ��
        CopyMemory tagOldSA, ByVal pSA, Len(tagOldSA) '����ԭ����SAFEARRAY��Ա��Ϣ
        CopyMemory tagNewSA, tagOldSA, Len(tagNewSA) '����SAFEARRAY��Ա��Ϣ
        tagNewSA.rgsabound(0).cElements = nFileSize '�޸�����Ԫ�ظ���
        tagNewSA.pvData = lpszFileText '�޸��������ݵ�ַ
        CopyMemory ByVal pSA, tagNewSA, Len(tagNewSA) '��ӳ�������ݵ�ַ��������
        Dim m As Long, n As Long, k As Long, i As Long, p As Long
        Dim ibyte() As Byte
        ibyte = strText 'StrConv(strText, vbFromUnicode)
        m = UBound(ibyte)
        k = UBound(pbFileText)
        p = m - 1
        Dim x As Long
        x = 0
        For i = 0 To k
            If pbFileText(i) = ibyte(0) Then
                If pbFileText(i + m) = ibyte(m) Then
                    For n = 1 To p
                        If ibyte(n) = pbFileText(i + n) Then
                            x = x + 1
                            If x = p Then FindTextInFile = i: GoTo 100
                        Else
                            x = 0
                            Exit For
                        End If
                    Next
                End If
            End If
        Next
100
        CopyMemory ByVal pSA, tagOldSA, Len(tagOldSA) '�ָ������SAFEARRAY�ṹ��Ա��Ϣ
        Erase pbFileText 'ɾ������
          
        UnmapViewOfFile lpszFileText 'ȡ����ַӳ��
        CloseHandle hFileMap '�ر��ļ�ӳ�����ľ��
    End If
    CloseHandle hFile '�ر��ļ�
End Function

Function CheckFileKeyWordx(ByVal FilePath As String, ByVal Keyword As String) As Boolean '���word�ļ��Ƿ�����ؼ���
    Dim wd As Object
    Dim reg As Object, Matches As Object
    '----------------------����ʹ��word find������https://docs.microsoft.com/zh-cn/office/vba/api/word.find.found
    '.found����true,���������ҵ�ƥ����
    CheckFileKeyWordx = False
    Set wd = CreateObject(FilePath)   'ִ���ٶȽ���,�����ٶ�Ҫ1s, ��Ҫ���ڴ���word�ļ�������һ��
'---If InStr(1, wd.Content.Text, keyword, vbBinaryCompare) > 0 Then CheckFileKeyWordx = True 'instr�ڴ��ı��Ĵ�����, �ٶ�Զ��������
    Set reg = CreateObject("VBScript.RegExp")
    With reg
        .Pattern = Keyword
        .Global = True
        .IgnoreCase = True
        Set Matches = .Execute(wd.Content.Text)
        If Matches.Count > 0 Then CheckFileKeyWordx = True
'        CheckFileKeyWordx = .test(keyword)
    End With
    wd.Close savechanges:=False
    Set reg = Nothing
    Set Matches = Nothing
    Set wd = Nothing
End Function

Sub StopTimer() '��ʱ�� /�����Ͻϸ߾���
With New Stopwatch
    .Restart
    Debug.Print FindTextInFile("C:\Users\adobe\Desktop\x31.txt", "��ϲ��")
    .Pause
    Debug.Print Format(.Elapsed, "0.000000"); " seconds elapsed"
End With
End Sub

Private Function Words_Static(ByVal FilePath As String) As String() 'word�ĵ�����ֱ��ʹ��word����������ȡ����
    Dim objwd As Object
    Dim objwords As Object
    Dim i As Long, k As Long
    Dim dic As Dictionary
    Dim strTemp As String, strText As String
    Dim arr() As String
    Dim myreg As Object
    
    Set objwd = CreateObject(FilePath)
    Set objwords = objwd.Content.Words 'https://docs.microsoft.com/zh-cn/office/vba/api/word.words
    i = objwords.Count
    If i = 0 Then Exit Function
    Set myreg = CreateObject("VBScript.RegExp")
    With myreg
        .Pattern = "[a-z]+['|-|��]?[a-z]{1,}"
        .Global = True '�����ִ�Сд
        .IgnoreCase = True
    End With
    Set dic = New Dictionary
    dic.CompareMode = vbTextCompare
    For k = 1 To i
        strTemp = Trim(objwords(k).Text)
        If Len(strTemp) > 1 Then
            strText = myreg.Replace(strTemp, "")
            If Len(strText) = 0 Then
                If dic.Exists(strTemp) Then
                    dic(strTemp) = dic(strTemp) + 1 '��¼���ֵĴ���
                Else
                    dic.Add strTemp, 1  '�����δ����,�����ֵ/item=1
                End If
            End If
        End If
    Next
    i = 0
    i = dic.Count
    If i = 0 Then Set dic = Nothing: Set objwords = Nothing: Exit Function
    i = i - 1
    ReDim arr(i, 1)
    ReDim Words_Static(i, 1)
    For k = 0 To i
        arr(k, 0) = dic.Keys(k)
        arr(k, 1) = dic.Items(k)
    Next
    Words_Static = arr
    Erase arr
    objwd.Close
    Set objwd = Nothing
    Set objwords = Nothing
    Set dic = Nothing
End Function

Sub WordAnalysis(ByVal FilePath As String, Optional ByVal strLen As Integer = 35) '���ʷ���/֧��word/text�ļ����ı��ļ�, ���ʹ��txt�ĵ�, Խ�򵥵��ļ�����Խ��
    Dim textx As String, dic As Object
    Dim myreg As Object, Matches As Object, match As Object
    Dim arrTemp As Variant, i As Long, j As Long, temp As String, c As String, arr() As Integer, p As Long
    Dim wb As Workbook
    Dim wordapp As Object, ado As Object
    Dim newdc As Object, strx As String, strx1 As String
    Dim strTemp As String
    
    If fso.fileexists(FilePath) = False Then Exit Sub
    ThisWorkbook.Application.ScreenUpdating = False
    strx = Right$(FilePath, Len(FilePath) - InStrRev(FilePath, "."))
    strx = LCase(strx)
    If strx = "txt" Or strx Like "doc*" Then
        If strx <> "txt" Then
            Set wordapp = CreateObject("Word.Application") '��ȡword�����ݷǳ���(��Ҫ�Ǵ���word���ٶȷǳ���)
            Set newdc = wordapp.documents.Add(FilePath)
            '����Dim obj As Object
            'Set obj = CreateObject("C:\Users\*.docx")'����ֱ�Ӵ���doc�ļ��Ķ���
            wordapp.Visible = False
            textx = newdc.Content.Text
            newdc.Close
            wordapp.Quit
            Set newdc = Nothing
            Set wordapp = Nothing
        Else
            Set ado = New ADODB.Stream
            With ado
                .Mode = 3 '��дȨ��
                .type = 2 '��ȡ�ı�
                .CharSet = "us-ascii"  '�ؼ�, ��Ȼ��ȡ�����ݻ��������, �����������Ӣ��(��Ӣ��)��������ѡascii�ַ���(��������Ӣ���),
                '-----------------------https://docs.microsoft.com/en-us/previous-versions/exchange-server/exchange-10/ms526296(v=exchg.10)
                .Open
                .LoadFromFile (FilePath) '�����ļ�
                textx = .ReadText()
                .Close
                Set ado = Nothing
            End With
        End If
    Else
        Exit Sub
    End If
    If Len(textx) = 0 Then Exit Sub
    Set dic = New Dictionary
    dic.CompareMode = vbTextCompare '�����ִ�С
    Set myreg = CreateObject("VBScript.RegExp") '������ʽ
    With myreg
        .Pattern = "[a-z]+['|-|��]?[a-z]{1,}"   'ƥ��Ҫ��, ��ĸ,������" - ",���ȴ���2�ĵ���,�������'����-���� ' �������ַ�,���������, �� what's, (С��(ȱʡ)35�ĵ���)"
                                               ' ?��ʾǰ��ķ��ſ��Գ���һ�λ��߲����� |��ʾ����� "��" +ƥ����,{1,} ���ӵĳ��ȴ���1
        .Global = True '�����ִ�Сд           'ע��chr(39)��chrw(8217)�ķ��ŵ�����
        .IgnoreCase = True
        Set Matches = .Execute(textx)
    End With
    For Each match In Matches  'ͳ��Ƶ��
        strTemp = match.Value
        If dic.Exists(strTemp) Then dic(strTemp) = dic(strTemp) + 1 Else dic.Add strTemp, 1
    Next
    Set Matches = Nothing
    p = dic.Count - 1
    If p < 5 Then MsgBox "����̫�ٲ��߱�������ֵ", vbInformation, "Tips": Exit Sub
'    For i = 0 To p - 1 '--------------�������� ,��Excel�в���������ֱ��ʹ��Excel�Դ�������(���������ݵ�����������, ������)
'        For j = i + 1 To p
'            If arrtemp(i) > arrtemp(j) Then
'                temp = arrtemp(i)
'                arrtemp(i) = arrtemp(j)
'                arrtemp(j) = temp
'            End If
'        Next
'    Next
    ReDim arr(p)
    ReDim arrTemp(p)
    For i = 0 To p             '��ȡ���ֵĴ���
       arrTemp = dic.Keys(i)
       arr(i) = dic.Items(i)
    Next
    Set wb = Workbooks.Add
    p = p + 1
    With wb            '---------------------�����������������
        .Worksheets(1).Name = "�������"
        With .Worksheets(1)
            .Cells(1, 1) = "����:"
            .Cells(1, 2) = "���ִ���:"
            .Cells(2, 2).Resize(p) = Application.Transpose(arr)
            .Cells(2, 1).Resize(p) = Application.Transpose(arrTemp)
            '--------------------------------------------------------����д��
            .Cells(1, 5) = "�ۼƷ�����������:"
            .Cells(1, 6) = dic.Count
            .Cells(2, 5) = "����ʱ��:"
            .Cells(2, 6) = Format(Now, "yyyy-mm-dd")
            '------------------------------------------------------------Excel���õ�������
            With .sort
                .SortFields.Clear
                .SortFields.Add key:=Range("B1"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                .SetRange Range("A2:B" & p + 1)
                .Header = xlGuess
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            If p > 25 Then p = 25
            p = p + 1
            For i = 2 To p
                .Cells(i, 1) = WorksheetFunction.Proper(.Cells(i, 1).Value) '���ⲿ�ֵ����ݵ��׸���ĸתΪ��д,���ಿ��תΪСд
            Next
            CreatChart wb, p '��������ͼ��
            .Cells(1, 5).ColumnWidth = 18
            .Cells(1, 6).ColumnWidth = 18 '������ʾ�ĸ��ӵĴ�С
            .Cells(1, 5).HorizontalAlignment = xlRight '�������з�ʽ
            .Cells(2, 5).HorizontalAlignment = xlRight
            .Cells(1, 3).Select
        End With
        strx1 = Left$(FilePath, InStrRev(FilePath, ".") - 1) '�ļ�������ļ�
        strx1 = strx1 & Format(Now, "yyyymmddhhmmss")
        strx = strx1 & ".xlsx"
        If fso.fileexists(strx) = True Then fso.DeleteFile strx
        If Err.Number = 70 Then Err.Clear: strx = strx1 & CStr(RandNumx(1000)) & ".xlsx"
        .SaveAs strx                           '���ļ����浽ͬλ����
    End With
    Erase arr
    Erase arrTemp
    Set dic = Nothing
    Set myreg = Nothing
    ThisWorkbook.Application.ScreenUpdating = True
End Sub

Private Sub CreatChart(ByVal wb As Workbook, ByVal numx As Byte) '����ͼ��
    Dim Shx As Shape
    Dim Cha As Chart
    Dim dTextx As Shape, rTextx As Shape, aTextx As Shape
    
    Set Shx = wb.Sheets(1).Shapes.AddChart2(201, xlColumnClustered, 250, 60, 720, 350) 'top=60,height=350,width=720
    Set Cha = Shx.Chart
    With Cha
        .SetSourceData Source:=wb.Sheets(1).Range("A1:B" & numx)
        '------------------------------------------------------����Դ
        numx = numx - 1
        .ChartTitle.Text = "Word Top" & CStr(numx) '����
        '--------------------------------------����
        .ApplyLayout (9) '----------------ͼ������� 'ͨ��¼�ƻ�ȡ����ֵ��6(����Ŀ����Ҫ��ͼ������)
        '----------ͼ��ķ��(ע�ⲻ������)
        .PlotArea.Select
        .PlotArea.Width = 634
        .PlotArea.Left = 24
        .PlotArea.Top = 39
        .PlotArea.Height = 276 '-------��ͼ�����С����
        '-----------------------------------------ͼ����ͼ����
        .FullSeriesCollection(1).ApplyDataLabels '��״ͼ��ʾ����
        '-------------------------------------------------------��״ͼ��������ʾ����
        .Axes(xlValue).AxisTitle.Format.TextFrame2.TextRange.Font.Size = 9.5
        .Axes(xlValue, xlPrimary).AxisTitle.Top = 128
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "���ʳ��ִ���" 'y����Ϣ
        '------------------------------------------------------------------------Y�����
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = ChrW(9670) & "����"
        .Axes(xlCategory, xlPrimary).AxisTitle.Left = 658
        .Axes(xlCategory, xlPrimary).AxisTitle.Top = 298
        .Axes(xlCategory).TickLabels.Font.Size = 11 '����x�������ֵĴ�С(¼�Ƶĺ��Ǵ��)
        '--------------------------------------------------------------------------------------------X�����
        Set dTextx = .Shapes.AddTextbox(msoTextOrientationHorizontal, 612, 0, 108, 16) '����ı���
        Set rTextx = .Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 332, 255, 16) '����ı���
        Set aTextx = .Shapes.AddTextbox(msoTextOrientationHorizontal, 612, 332, 108, 16) '����ı���
        '------------��������ı���
    End With
    With dTextx 'ʱ��
        .TextFrame.Characters.Text = "Date: " & Format(Now, "yyyy-mm-dd") '�ı���д����Ϣ
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignRight '�ı��������Ҷ���
    End With
    With rTextx '��Դ
        .TextFrame.Characters.Text = "Resource: File Analysis" '�ı���д����ϢWallstreet Journal
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft '�ı�������������
    End With
    With aTextx '����
        .TextFrame.Characters.Text = "Drawing By: HLA" '�ı���д����Ϣ
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignRight '�ı��������Ҷ���
    End With
    '-------------------------�ı������
    wb.Sheets(1).ChartObjects(1).Placement = xlFreeFloating 'ͼ���λ�ò�����Ϊ���������仯
    With Shx.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
    End With
   '-----------------------------�߿����
    Set dTextx = Nothing
    Set aTextx = Nothing
    Set rTextx = Nothing
    Set Shx = Nothing
    Set Cha = Nothing
End Sub

Private Function Min(ByVal one As Integer, ByVal two As Integer, ByVal three As Integer) As Integer
    Min = one
    If (two < Min) Then Min = two
    If (three < Min) Then Min = three
End Function
 
Private Function CompString(ByVal str1 As String, ByVal str2 As String, ByVal n As Integer, m As Integer) As Integer
    Dim i, j As Integer, p As Integer, q As Integer
    Dim ch1, ch2 As String
    Dim arr() As Integer
    Dim temp As Byte
    
    If (n = 0) Then CompString = m
    If (m = 0) Then CompString = n
    ReDim arr(n + 1, m + 1)
    arr(0, 0) = 0
    For i = 1 To n
        arr(i, 0) = i
        ch1 = Mid(str1, i, 1)
        For j = 1 To m
            arr(0, j) = j
            ch2 = Mid(str2, j, 1)
            If (ch1 = ch2) Then
                temp = 0
            Else
                temp = 1
            End If
            p = i - 1
            q = j - 1
            arr(i, j) = Min(arr(p, j) + 1, arr(i, q) + 1, arr(p, q) + temp)
        Next
    Next
    CompString = arr(n, m)
End Function
 
Function Similar(ByVal str1 As String, ByVal str2 As String) As Double '˳���ַ������ƶȱȽ�
    Dim ldint As Integer
    Dim i As Integer, k As Integer
    Dim strLen As Integer
    
    i = Len(str1)
    k = Len(str2)
    ldint = CompString(str1, str2, i, k)
    If (i >= k) Then
        strLen = i
    Else
        strLen = k
    End If
    Similar = 1 - ldint / strLen
End Function

Function CheckFileKeyWordB(ByVal FilePath As String, ByVal Keyword As String, Optional ByVal cmCode As Byte = 0) As Boolean '���txt�ļ����Ƿ����ָ���Ĺؼ���
    Dim obj As Object
    Dim strx As String * 1024 '�����ַ���
    
    CheckFileKeyWordB = False
    Set obj = fso.OpenTextFile(FilePath, ForReading) 'ע���ı��ı���, ansi/Unicode/uft8������
    With obj
        Do While Not .AtEndOfStream                    'binaryģʽΪ0,textΪ1, dataΪ-2, ��Ҫ���ִ�Сʱ,ʹ��text
            strx = .Read(1024)                   '���ַ���ƫ��,����һ��18M���ļ������Ҫ1.5s/vbtext,vbinary �����Ƶı��ٶȸ���,���Դﵽ0.95s����
            If InStr(1, strx, Keyword, cmCode) > 0 Then CheckFileKeyWordB = True: Exit Do
        Loop
        .Close
    End With
    Set obj = Nothing
End Function

Sub dkslla()
'With New Stopwatch
'    .Restart
'    Debug.Print CheckFileKeyWord("C:\Users\adobe\Desktop\x31.txt", "��ϲ��", 0, 4)
'    .Pause
'    Debug.Print Format(.Elapsed, "0.000000"); " seconds elapsed"
'End With
Dim ad As New ADODB.Stream
Dim arr() As Byte
With ad
.Mode = 3
.type = 1
.Open
.LoadFromFile "C:\Users\adobe\Desktop\testcode\utf_nobom.txt"
arr = .Read(-1)
.Close
End With
UEFCheckUTF8NoBom arr
Set ad = Nothing
End Sub

Private Function UEFCheckUTF8NoBom(ByRef bufAll() As Byte)
    
    Dim i As Long
    Dim cOctets As Long         '��������UTF-8�����ַ����ֽڴ�С 4bytes
    Dim bAllAscii As Boolean    '���ȫ��ΪASCII��˵������UTF-8
    Dim fmt
    bAllAscii = True
    cOctets = 0
    
    For i = 0 To UBound(bufAll)
        If (bufAll(i) And &H80) <> 0 Then
            'ASCII��7λ���棬���λΪ0��������������0���Ͳ���ASCII
            '���ڵ��ֽڵķ��ţ��ֽڵĵ�һλ��Ϊ0������7λΪ������ŵ�unicode�롣
            '��˶���Ӣ����ĸ��UTF-8�����ASCII������ͬ��
            bAllAscii = False
        End If
        
        '����n�ֽڵķ��ţ�n>1������һ���ֽڵ�ǰnλ����Ϊ1����n+1λ��Ϊ0�������ֽڵ�ǰ��λһ����Ϊ10
        'cOctets = 0 ��ʾ���ֽ���leading byte
        If cOctets = 0 Then
            If bufAll(i) >= &H80 Then
                '��������cOctets�ֽڵķ���
                Do While (bufAll(i) And &H80) <> 0
                    'bufAll(i)����һλ
                    bufAll(i) = ShLB_By1Bit(bufAll(i))
                    cOctets = cOctets + 1
                Loop
                
                'leading byte����ӦΪ110x xxxx
                cOctets = cOctets - 1
                If cOctets = 0 Then
                    '����Ĭ�ϱ���
                    fmt = "UEF_ANSI"
                    Exit Function
                End If
            End If
        Else
            '��leading byte��ʽ������ 10xxxxxx
            If (bufAll(i) And &HC0) <> &H80 Then
                '����Ĭ�ϱ���
                fmt = "UEF_ANSI"
                Exit Function
            End If
            '׼����һ��byte
            cOctets = cOctets - 1
        End If
    
    Next i
    
    '�ı�����.  ��Ӧ���κζ����byte �м�Ϊ���� ����Ĭ�ϱ���
    If cOctets > 0 Then
        fmt = "UEF_ANSI"
        Exit Function
    End If
    
    '���ȫ��ascii.  ��Ҫע�����ʹ����Ӧ��code pages��ת��
    If bAllAscii = True Then
        fmt = "UEF_ANSI"
        Exit Function
    End If
    
    '�޳����� ���ڸ�ʽȫ����ȷ ����UTF8 No BOM�����ʽ
    fmt = "UEF_UTF8NB"
    
End Function

Private Function ShLB_By1Bit(ByVal Byt As Byte) As Byte

'����BYTE���׃������1λ�ĺ���������Byt�Ǵ���λ���ֹ�������������λ�Y��
'
'��(Byt And &H7F)���������������λ�� *2������һλ

ShLB_By1Bit = (Byt And &H7F) * 2

End Function

Function CheckFileKeyWord(ByVal FilePath As String, ByVal Keyword As String, Optional ByVal cmCode As Byte = 0, _
Optional ByVal iCodex As Byte = 1, Optional ByVal isIgnoreCase As Boolean = True) As Boolean
    '----------------------------cmcodeָ���Ƚϵ�����,cmcodexָ����ȡ���ļ�����
    Const defBuffer As Long = 131072 '1024 * 128, 128K
    Const chBuffer As Long = 4096
    Dim oAstream As Object
    Dim Codex As String
    Dim sBuffer As String   '* 1024 '�����ַ��������ڼӿ�������ٶ�
    Dim iPostion As Long, iBuffer As Long
    Dim iType As Byte
'    Dim cr As New cRegex
    '----------------http://wsh.style-mods.net/ref_stream/readtext.htm
    '����,���Ǳ���ı���
    'Ӣ��,���Ǳ���,���Ǵ�Сд(�����Ƶķ��������ִ�Сд) '����������ı��Ƚ�
    CheckFileKeyWord = False
    Select Case iCodex
        Case 1: Codex = "gb2312" '����/ansi '��õı���
        Case 2: Codex = "us-ascii" '��Ӣ��
        Case 3: Codex = "uft-8"
        Case 4: Codex = "unicode"
    End Select
    Set oAstream = New ADODB.Stream
    With oAstream
        .Mode = 3                   '��дȨ��
        If LenB(Codex) = 0 Then
            Dim tBytes() As Byte
            .type = adStreamType.adTypeBinary
            .Open
            .LoadFromFile FilePath
            iSize = .Size           '�ļ���С
            If iSize < 3 Then Exit Function
            tBytes = .Read(3)
            .Close
            If tBytes(0) = 239 Then
                If tBytes(1) = 187 And tBytes(2) = 191 Then Codex = "utf-8"
            ElseIf tBytes(0) = 255 Then
                If tBytes(1) = 254 Then Codex = "unicode"
            ElseIf tBytes(0) = 254 Then
                If tBytes(1) = 255 Then Codex = "utf-8"
            Else
            If iSize < chBuffer Then iBuffer = -1 Else iBuffer = chBuffer
            .Position = 0
            If iBuffer > 0 Then ReDim Bytes(iBuffer) Else ReDim tBytes(iSize)
            tBytes = .Read(iBuffer)                                           '��ȡ�������������жϾ�����Ƿ�Ϊutf-8 without BOM
            Check_Unicode (tBytes)
            End If
        End If
        .type = adStreamType.adTypeText '��ȡ�ı�
        .CharSet = Codex            '�ؼ�, ��Ȼ��ȡ�����ݻ��������
        .Open                       'https://docs.microsoft.com/en-us/previous-versions/exchange-server/exchange-10/ms526296(v=exchg.10)
        .LoadFromFile (FilePath)    '�����ļ�
        
        sBuffer = Space$(iSize)     '���Ĺؼ���һ��, �����������ٶ��𵽺�������, ���ַ����Ĺ���
        If iSize > defBuffer Then   '�ַ����Ļ�������СΪ1024 * 64, 64K, 65536
            iBuffer = defBuffer
        Else
            iBuffer = mReadText.adReadAll
        End If
        iPostion = 1                '1.64m���ļ���С��ȡ���������Ҫ0.0065-0.0070֮��
        .Position = 0               'ָ������ʼ������ǰλ�õ�ƫ���ֽ�����Ĭ��ֵΪ 0����ʾ���еĵ�һ���ֽڡ�
        Do Until .EOS = True        'eos��ʾ�ļ�����ĩβ '������ʾ��С������ֵ���ٶȶ������Ե�Ӱ��, 1024-131072֮������ݵ�Ч������Ϊ���Ե�
            Mid$(sBuffer, iPostion, iSize) = .ReadText(iBuffer)
            iPostion = iBuffer + i
                                    '������Ҫע��, ���������ȡ, ����ļ����󽫻ᵼ���ٶȴ�����½�,��1.64m�ļ���ȡΪ��, �ٶ��½�����20��
        Loop
        .Close
    End With
    Set oAstream = Nothing
    If InStr(1, sBuffer, Keyword, cmCode) > 0 Then
        CheckFileKeyWord = True
    Else
        If isIgnoreCase = False Then
            If InStr(1, sBuffer, Keyword, cmCode) > 0 Then CheckFileKeyWord = True
        End If
    End If
    sBuffer = vbNullString
End Function



Sub dkkd()
Dim cr As New cRegex
With cr
cr.oReg_Initial
cr.oReg_Pattern = "ab"
cr.oReg_Text = "babb"
Debug.Print cr.cTest
End With
Set cr = Nothing
End Sub

Function CheckFileKeyWordC(ByVal FilePath As String, ByVal Keyword As String, Optional ByVal blockSize As Long = 131072) As Boolean '�������ж��Ƿ�����ؼ���
    Dim arr() As Byte
    Dim arrx() As Byte
    Dim arrTemp() As Byte
    Dim ado As Object

    CheckFileKeyWordC = False
    Set ado = New ADODB.Stream
    ReDim arrx(blockSize - 1)
    With ado
        .Mode = 3
        .type = adTypeBinary
        .Open
        .Position = 0
        .LoadFromFile FilePath
        arrTemp = .Read(3)     '����ͨ���������, ����BOM, ��ȡ֮��,����3��λ��֮��ʼ��ȡ
        '-------------------https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/mid-function
        If AscB(MidB(arrTemp, 1, 1)) = &HEF And AscB(MidB(arrTemp, 2, 1)) = &HBB And AscB(MidB(arrTemp, 3, 1)) = &HBF Then '�ж��ļ��ı�������
            CheckFileKeyWordC = CheckFileKeyWord(FilePath, Keyword, , 3)
            .Close
            Set ado = Nothing
            Exit Function '"utf-8"  'uft8�Ĵ�����ʱ�Ҳ������ʵĴ�����
        ElseIf AscB(MidB(arrTemp, 1, 1)) = &HFF And AscB(MidB(arrTemp, 2, 1)) = &HFE Then
            arr = Keyword '"unicode"
        Else
            arr = StrConv(Keyword, vbFromUnicode) '"ANSI/gb2312"
        End If
         .Position = 0 'ע�������ȡ����֮�����»ָ���ȡ���ݿ�ʼ��λ��
        Do Until .EOS = True
            arrx = .Read(blockSize)
            If InStr(1, arrx, arr, vbBinaryCompare) > 0 Then CheckFileKeyWordC = True: Exit Do
        Loop
        .Close
    End With
    Erase arr: Erase arrx
    Set ado = Nothing
End Function

Function CheckTextCode(ByVal FilePath As String) As String '����ļ��ı�������
'---------------https://blog.csdn.net/hongsong673150343/article/details/88584753
'  ANSI�޸�ʽ����
'  EFBBBF    UTF-8
'  FFFE      UTF-16/UCS-2, little endian
'  FEFF      UTF-16/UCS-2, big endian
'  FFFE 0000 UTF-32/UCS-4, little endian
'  0000 FEFF UTF-32/UCS-4, big endian
' ��Ҫע�����utf, ��uft8��, ascii���������Ҳ����1λ�ֽڴ洢��, ���Զ�ȡ��Щ���ݵ�ʱ���޲��, Unicodeͳһ����2λ, �ֽ�, ����Ҫע��utf-8 ��bomͷ�����
'- ------------------------------------------- -
    Dim arr() As Byte
    Dim ado As Object
    Set ado = New ADODB.Stream
    With ado
        .type = 1
        .Mode = 3
        .Open
        .Position = 0
        .LoadFromFile FilePath
        arr = .Read(3)            'ascii��239,187,191����BOMͷ�� EF BB BF
        If AscB(MidB(arr, 1, 1)) = &HEF And AscB(MidB(arr, 2, 1)) = &HBB And AscB(MidB(arr, 3, 1)) = &HBF Then
            CheckTextCode = "utf-8"
        ElseIf AscB(MidB(arr, 1, 1)) = &HFF And AscB(MidB(arr, 2, 1)) = &HFE Then '255 '254
            CheckTextCode = "unicode"
        Else
            CheckTextCode = "gb2312" 'ansi
        End If
        .Close
    End With
    Set ado = Nothing
End Function

Function StrChinese(ByVal strText As String) As Boolean '�ж��ַ����Ƿ��������, �����������ĺ��ֺ����ĺ���(����ȫ����, Ƭ������), �������ķ���Ӱ��
    strText = StrConv(strText, vbNarrow)                'vbNarrow ���ַ�����˫�ֽ��ַ�ת�ɵ��ֽ��ַ�
    StrChinese = IIf(Len(strText) < LenB(StrConv(strText, vbFromUnicode)), True, False) '���ֻ��ת����Сд, ������ʹ��strconv����, ���Ե�����lcase, ucase����
End Function

'�����Ԥ��(negative assert)�����κβ�ƥ��pattern���ַ�����ʼ��ƥ������ַ���������һ���ǻ�ȡƥ�䣬
'Ҳ����˵����ƥ�䲻��Ҫ��ȡ���Ժ�ʹ�á�����"Windows(?!95|98|NT|2000)"��ƥ��"Windows3.1"�е�"Windows"��
'������ƥ��"Windows2000"�е�"Windows"��Ԥ�鲻�����ַ���Ҳ����˵����һ��ƥ�䷢���������һ��ƥ��֮��������ʼ��һ��ƥ��������������ǴӰ���Ԥ����ַ�֮��ʼ��
Sub Book_Analysis(ByVal FilePath As String)
    Dim arr() As Long
    Dim arrx
    Dim cre As New cRegex
    Dim strText As String
    Dim ado As New ADODB.Stream
    Dim i As Byte
'    Dim wb As Workbook
    
    DisEvents
'    arrx = Sheet8.Range("a1:a13").Value
'    ReDim arr(1 To 13)
    With ado
        .CharSet = "gb2312"
        .Mode = adModeReadWrite
        .type = adTypeText
        .Open
        .LoadFromFile FilePath
        strText = .ReadText
        .Close
    End With
    Set ado = Nothing
'    With cre
'        For i = 1 To 13
'            .oReg_Initial arrx(i, 1)
'            .sMatch strText
'            arr(i) = .sFirst_Index
'        Next
'    End With
'    strText = ""
     '"ͩԭ(?!���|������)��˾?" , ƥ��ͩԭ����ͩԭ��˾, ���ǲ�ƥ��ͩԭ����ͩԭ������
     '(����|����)?(?!�Ĵ�|����)(ѩ��) , ƥ������ѩ��,����ѩ�����ѩ��,���ǲ�ƥ����������, ��ƥ�������Ĵ�
     '�Gԫ(����)?, ƥ��Gԫ���߹Gԫ����
     With cre
        .oReg_Initial "�Gԫ(����)?"
        .xMatch strText
        arr = .aFirst_Index
     End With
    Set cre = Nothing
    Sheet8.Cells(1, 4).Resize(UBound(arr) + 1, 1) = Application.Transpose(arr)
'    Set wb = Workbooks.Add
'    With wb.Sheets(1)
'        .Cells(1, 1).Resize(44, 1) = arrx
'        .Cells(1, 2).Resize(44, 1) = wb.Application.Transpose(arr)
'    End With
'    Set wb = Nothing
    EnEvents
End Sub

Sub dkklldso()
Book_Analysis "C:\Users\adobe\Desktop\��ҹ��.txt"
End Sub


