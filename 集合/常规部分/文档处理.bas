Attribute VB_Name = "�ĵ�����"
Option Explicit
'------------------�漰PDF�Ĵ���,��ӡ, ���ˮӡ, ���ͼƬ, ����ҳ��, �ļ����ܼ��,���ܵȵȵĲ���,��Ҫ������Adobe Acrobat��PDFCreator
'-------------------https://acrobat.adobe.com/us/en/acrobat/pdf-reader.html
'-------------------https://www.pdfforge.org/pdfcreator
'-------------------��صĲ������߷����Ĳο���Դ��Ҫ������ṩ��SDK�ĵ�
'ע������:(�޷���ȡ)
'�ļ���
'�ļ�����(��ͬ��ʽ�ļ���),���ּ����ļ���Ȼ�����޸Ĳ��ֵ��ļ���Ϣ,�������ļ����ķ�ʽ���ļ�д������, ��һ���ֵļ��ܽ���ȫ�޷��޸�,ǿ���޸Ŀ�������ļ�����
#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute _
        Lib "shell32.dll" Alias "ShellExecuteA" _
            (ByVal hwnd As LongPtr, _
             ByVal lpOperation As String, _
             ByVal lpFile As String, _
             ByVal lpParameters As String, _
             ByVal lpDirectory As String, _
             ByVal nShowCmd As Long) _
        As LongPtr
#Else
    Private Declare Function ShellExecute _
        Lib "shell32.dll" Alias "ShellExecuteA" _
            (ByVal hwnd As Long, _
             ByVal lpOperation As String, _
             ByVal lpFile As String, _
             ByVal lpParameters As String, _
             ByVal lpDirectory As String, _
            ByVal nShowCmd As Long) _
        As Long
#End If
Private Const SW_HIDE = 0

Sub FilePrint(ByVal strFilePath As String) '��ӡ�ļ�
    Dim retVal As Long
        retVal = ShellExecute(0, "Print", strFilePath, 0, 0, SW_HIDE)
        If retVal < 32 Then
            MsgBox "An Error occured...could not print"
        End If
End Sub

'DefaultGuid Creates PDF files
'PdfaGuid    Creates PDF/pdfobj files
'PrintGuid   Sends output to pdfobj printer after conversion
'HighCompressionGuid PDFs with high compression images
'HighQualityGuid PDFs with high quality images
'JpegGuid    Creates JPEG files
'PngGuid Creates PNG files
'TiffGuid    Creates TIFF files
'https://www.pdfforge.org/, ������PDFCreator
'���´������3.3����г�汾����ע�������ִ��vba���붼�ǻ����Ͼɰ汾��pdfcreator���Ѿ����ʺ����°汾��ʹ��
'��Ҫע�⣬PdfCreator��װλ���µ�chm manual�кܶ���󣬰�����ҳ�ĺܶ����ݶ��Ǵ����,������һЩƴд�����ִ���<--->
'http://docs.pdfforge.org/pdfcreator/latest/en/pdfcreator/com-interface/
'http://docs.pdfforge.org/pdfcreator/latest/en/pdfcreator/com-interface/reference/settings/ '�����������ҳ������Ϊ��׼
Sub EncryptPDF(ByVal FilePath As String, ByVal outputfilepath As String, Optional ByVal owerpassword As String = "password", _
Optional ByVal userpassword As String = "password", Optional ByVal cmCode As Byte = 0, Optional ByVal securitylevel As Byte = 0) 'ע���޷�ֱ����ӷ�pdf�ļ�������
    Dim pdfobj As New PDFCreator_COM.PdfCreatorObj
    Dim pdfq As New PDFCreator_COM.Queue
    Dim job As Object
    Dim Encrypt As String
    'Rc40Bit, Rc128Bit, Aes128Bit, Aes256Bit '���ܷ�ʽ
    'owerpassword---------���Ʊ༭����
    'userpassword--------���ļ�����
    On Error GoTo ErrHandle
    Select Case securitylevel
        Case 1: Encrypt = "Rc40Bit"
        Case 2: Encrypt = "Rc128Bit"
        Case 3: Encrypt = "Aes256Bit" '�����������ҵ�汾
        Case Else: Encrypt = "Aes128Bit" 'Ĭ��ʹ�ô���Ŀ
    End Select
    
    If pdfobj.IsInstanceRunning = False Then pdfq.Initialize '�ж�PdfCreator�Ƿ������е�״̬
    pdfobj.AddFileToQueue (FilePath) '����ļ�������
    Set job = pdfq.NextJob
    With job
        .SetProfileByGuid ("DefaultGuid") '�ļ���ʽ
        .SetProfileSetting "PdfSettings.Security.Enabled", "true" '��ȫ����
        .SetProfileSetting "PdfSettings.Security.EncryptionLevel", Encrypt '���ܷ�ʽ
        If cmCode = 1 Then
            .SetProfileSetting "PdfSettings.Security.RequireUserPassword", "true"
            .SetProfileSetting "PdfSettings.Security.UserPassword", userpassword
        Else
            .SetProfileSetting "PdfSettings.Security.OwnerPassword", owerpassword
        End If
        .ConvertTo outputfilepath 'pdf��ʽ������µ��ļ�
    End With
    pdfq.ReleaseCom
    Set pdfobj = Nothing
    Set pdfq = Nothing
    Set job = Nothing
    Exit Sub
ErrHandle:
    If Not pdfq Is Nothing Then pdfq.ReleaseCom
    Set pdfobj = Nothing
    Set pdfq = Nothing
    Set job = Nothing
End Sub

Function WordConvertToPDF(ByVal FilePath As String, ByVal outputfilepath As String) As Boolean 'ֱ�ӽ�word�ļ�ת��pdf��ͬʱ����
    Dim printJob As Object
    Dim wd As Object
    Dim wdoc As Object
    Dim pdfobj As New PDFCreator_COM.PdfCreatorObj
    Dim pdfq As New PDFCreator_COM.Queue
    
    On Error GoTo ErrHandle
    WordConvertToPDF = False
    Set wd = CreateObject("word.application")
    wd.Application.ActivePrinter = "PDFCreator"
    Set wdoc = wd.documents.Open(FilePath, Visible:=False, ReadOnly:=True)
    wdoc.Activate
    'PrintOut(Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, FileName, ActivePrinterMacGX,
    'ManualDuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight)
    wdoc.PrintOut Background:=False '������Ա��룬ʹ�ü����ʹ��
    If pdfobj.IsInstanceRunning = False Then pdfq.Initialize '��f8����ģʽ�£�����ִ���
    If pdfq.WaitForJob(10) = True Then '�ж��Ƿ��ڹ涨��ʱ������ӵ�������
        Set printJob = pdfq.NextJob
        With printJob
            .SetProfileByGuid ("DefaultGuid")
            .SetProfileSetting "PdfSettings.Security.Enabled", "true" '��ȫ����
            .SetProfileSetting "PdfSettings.Security.EncryptionLevel", "Aes128Bit" '���ܷ�ʽ
            .SetProfileSetting "PdfSettings.Security.OwnerPassword", "123"
            .ConvertTo outputfilepath
            If (Not .IsFinished Or Not .IsSuccessful) Then '�ж��Ƿ���Чִ��
                WordConvertToPDF = False
            Else
                WordConvertToPDF = True
            End If
        End With
    End If
    wdoc.Close
    wd.Quit
    Set wd = Nothing
    Set wdoc = Nothing
    pdfq.ReleaseCom
    Set printJob = Nothing
    Set pdfobj = Nothing
    Set pdfq = Nothing
    Exit Function
ErrHandle:
    wdoc.Close
    wd.Quit
    pdfq.ReleaseCom
    Set wd = Nothing
    Set wdoc = Nothing
    Set printJob = Nothing
    Set pdfobj = Nothing
    Err.Clear
    WordConvertToPDF = False
End Function

'Const PDSaveFull = 1
'Const PDSaveBinaryOK = 16 (&H10)
'Const PDSaveCollectGarbage = 32 (&H20)
'Const PDSaveCopy = 2
'Const PDSaveIncremental = 0
'Const PDSaveLinearized = 4
'Const PDSaveWithPSHeader = 8
Function addWaterMarktoPDF(ByVal FilePath As String, ByVal watermark As String, Optional ByVal startpage As Integer = 1, Optional ByVal endpage As Integer = 0) As Byte '���ˮӡ��pdf�ļ���ȱ�ݣ����ˮӡ����(�仯���Ƚϴ�
    Dim jso As Object
    Dim acroApp As Object
    Dim myDocument As Object
    Dim i As Integer, k As Integer
    '��һ�����ˮӡ�ķ�����ͨ���ļ� -addWatermarkFromFile ,�����ѡ��Դ�ļ�����pdf�ļ�,adobe��ת��pdf�ļ�
    'The device-independent path of the source file to use for the watermark. If the file at this location is not a PDF file, Acrobat attempts to convert the file to a PDF file
    On Error GoTo ErrHandle
    addWaterMarktoPDF = 0
    Set acroApp = CreateObject("AcroExch.App")
    Set myDocument = CreateObject("AcroExch.PDDOc")
    myDocument.Open (FilePath)
    i = myDocument.GetNumPages      'ע�������ҳ���0��ʼ��
    If endpage = 0 Then endpage = i - 1
    If startpage > endpage Then k = startpage: startpage = endpage: endpage = k
    Set jso = myDocument.GetJSObject
    jso.addWatermarkFromText _
        cText:=watermark, _
        nRotation:=45, _
        nOpacity:=0.5, _
        nTextAlign:=1, _
        nHorizAlign:=2, _
        nVertAlign:=4 ', _
        nStart:=startpage, _
        nEnd:=endpage     'ѡ��ʼ�ͽ�β��ҳ�棬 ��������ã���Ĭ�ϰ������е�ҳ��
'    myDocument.Save 1, filepath '1 '�Ӳ���Ч������, ��һ����ѡ��1��ʱ��Ч�����(�仯�������С)
    myDocument.Save PDSaveFull + PDSaveLinearized + PDSaveCollectGarbage, FilePath '����ͬʱʹ�ö������
ErrHandle:
    addWaterMarktoPDF = 1
    If Err.Number = 448 Then addWaterMarktoPDF = 2: Err.Clear '����ļ�������,������ִ���
    Set jso = Nothing
    myDocument.Close
    acroApp.Exit
    Set myDocument = Nothing
    Set acroApp = Nothing
End Function

Sub SavePDFAs(ByVal PDFPath As String, ByVal fileExtension As String) '���pdf�ļ�Ϊ(���ļ���ת��,����ʹ��)
    Dim objAcroApp As Object
    Dim objAcroAVDoc As Object
    Dim objAcroPDDoc As Object
    Dim objJSO As Object
    Dim boResult As Boolean
    Dim ExportFormat As String
    Dim NewFilePath As String

    Set objAcroApp = CreateObject("AcroExch.App")
    Set objAcroAVDoc = CreateObject("AcroExch.AVDoc")
    boResult = objAcroAVDoc.Open(PDFPath, "")
    Set objAcroPDDoc = objAcroAVDoc.GetPDDoc
    Set objJSO = objAcroPDDoc.GetJSObject
    Select Case LCase(fileExtension)
        Case "eps": ExportFormat = "com.adobe.acrobat.eps"
        Case "html", "htm": ExportFormat = "com.adobe.acrobat.html"
        Case "jpeg", "jpg", "jpe": ExportFormat = "com.adobe.acrobat.jpeg"
        Case "jpf", "jpx", "jp2", "j2k", "j2c", "jpc": ExportFormat = "com.adobe.acrobat.jp2k"
        Case "docx": ExportFormat = "com.adobe.acrobat.docx"
        Case "doc": ExportFormat = "com.adobe.acrobat.doc"
        Case "png": ExportFormat = "com.adobe.acrobat.png"
        Case "ps": ExportFormat = "com.adobe.acrobat.ps"
        Case "rft": ExportFormat = "com.adobe.acrobat.rft"
        Case "xlsx": ExportFormat = "com.adobe.acrobat.xlsx"
        Case "xls": ExportFormat = "com.adobe.acrobat.spreadsheet"
        Case "txt": ExportFormat = "com.adobe.acrobat.accesstext"
        Case "tiff", "tif": ExportFormat = "com.adobe.acrobat.tiff"
        Case "xml": ExportFormat = "com.adobe.acrobat.xml-1-00"
        Case Else: ExportFormat = "Wrong Input"
    End Select

    If ExportFormat <> "Wrong Input" And Err.Number = 0 Then
        'Set the path of the new file. Note that Adobe instead of xls uses xml files.
        'That's why here the xls extension changes to xml.
        If LCase(fileExtension) <> "xls" Then
            NewFilePath = WorksheetFunction.Substitute(PDFPath, ".pdf", "." & LCase(fileExtension))
        Else
            NewFilePath = WorksheetFunction.Substitute(PDFPath, ".pdf", ".xml")
        End If
        boResult = objJSO.SaveAs(NewFilePath, ExportFormat)
        boResult = objAcroAVDoc.Close(True)
        boResult = objAcroApp.Exit
    Else
        boResult = objAcroAVDoc.Close(True)
        boResult = objAcroApp.Exit
    End If
    Set objAcroPDDoc = Nothing
    Set objAcroAVDoc = Nothing
    Set objAcroApp = Nothing
End Sub

'----------------------------------------------------------���������ֵ,��������word�������������
Function AddWordPassword(ByVal FilePath As String, Optional ByVal cmCode As Byte = 0, Optional ByVal cmCodex As Byte = 0, _
Optional ByVal opassword As String = "abc", Optional ByVal wpassword As String = "abc") As Byte '�������word��������߱༭����
    Dim wd As Object
    Dim wdoc As Object, i As Integer
    '-------------- --- word�����������д�ʱ�������, �༭Ȩ�޵�����
    On Error GoTo ErrHandle
    AddWordPassword = 0
    Set wd = CreateObject("word.application")
    '----------------------------------------https://docs.microsoft.com/zh-cn/office/vba/api/Word.Documents.Open
    Set wdoc = wd.documents.Open(FilePath, PasswordDocument:=opassword, WritePasswordDocument:=wpassword, Visible:=False)  '���ļ����������ʱ��,������������ò�Ӱ���������ļ�����
    If wdoc Is Nothing Then GoTo ErrHandle
    If cmCodex = 1 Then
        AddWordPassword = 4
        wdoc.Close
        wd.Quit
        Set wdoc = Nothing
        Set wd = Nothing
        Exit Function 'cmcodex���ڱ�ʾֻ����ļ��Ƿ�������,���ǲ�����ִ���������
    End If
    If cmCode = 1 Then
        wdoc.Password = opassword
    Else
        If wdoc.WriteReserved = False Then wdoc.WritePassword = wpassword
    End If
    wdoc.Save
    wd.Quit
    Set wd = Nothing
    Set wdoc = Nothing
    Exit Function
ErrHandle:
    i = Err.Number
    If i = 5408 Or i = 5174 Then '�򿪱���
        AddWordPassword = 1
    ElseIf i = 4198 Then
        AddWordPassword = 2 '�༭����
    Else
        AddWordPassword = 3 '��������
    End If
    wd.Quit
    Err.Clear
    Set wd = Nothing
    Set wdoc = Nothing
End Function

Sub WordToPDF(ByVal FilePath As String, ByVal cmCode As Byte, Optional ByVal waterprint As String) '��word�ĵ�ת��pdf(��ˮӡ)
    Dim wordapp As Object
    Dim newdc As Object
    Dim strx As String
    Dim Errcount As Byte '���ڼ�¼��������Ĵ������
    '�����excel���pdf��Ҫprint service����
    Set wordapp = CreateObject("Word.Application")
    If wordapp Is Nothing Then MsgBox "����word����ʧ��": Exit Sub
    On Error GoTo ErrHandle
    strx = Environ("UserProfile") & "\Desktop\" & Split(Right$(FilePath, Len(FilePath) - InStrRev(FilePath, "\")), ".")(0) & ".pdf" '�µ��ļ���
    If cmCode <> 1 Then
'        Set newdc = wordapp.Documents.Add(filepath)
Renterpassword1:
        Set newdc = wordapp.documents.Open(FilePath, ReadOnly:=True, Visible:=False) '����readonly�����Ա��������б༭����,���������д�����,visible��Ӱ�쵯���������봰��
        wordapp.Visible = False
        newdc.ExportAsFixedFormat OutputFileName:=strx, ExportFormat:=17, OpenAfterExport:=True '��ص�����ϸ�ڿ��Ѳ鿴word chm�ļ�
        newdc.Close
        wordapp.Quit
        Set newdc = Nothing
        Set wordapp = Nothing
    ElseIf cmCode = 1 Then
        If Len(waterprint) = 0 Then waterprint = ThisWorkbook.Application.Username & "��Ʒ" 'ˮӡ
Renterpassword2:
        Set newdc = wordapp.documents.Open(FilePath, ReadOnly:=True, Visible:=False)
        If newdc Is Nothing Then MsgBox "����word����ʧ��": GoTo 100
        wordapp.Visible = False
        With newdc.Application
            .ActiveWindow.ActivePane.View.SeekView = 9   'Ҫ��wdSeekCurrentPageHeader���ֳ���д��ֵ�ķ���,����ֵ�ķ���,��word�д���sub debug.print wdRelativeVerticalPositionMargin(��Ӧ�ĳ���)
            .Selection.HeaderFooter.Shapes.AddTextEffect(1, waterprint, "����", 36, False, False, 0, 0).Select
            With .Selection.ShapeRange
                .Name = "PowerPlusWaterMarkObject1"
                .TextEffect.NormalizedHeight = False
                .Line.Visible = False
                .Fill.Visible = True
                .Fill.Solid
                .Fill.ForeColor.RGB = RGB(192, 192, 192)
                .Fill.Transparency = 0.5
                .Rotation = 0
                .LockAspectRatio = True
                .Height = wordapp.Application.CentimetersToPoints(1.27)
                .Width = wordapp.Application.CentimetersToPoints(8.25)
                .WrapFormat.AllowOverlap = True
                .WrapFormat.Side = 3
                .WrapFormat.type = 3
                .RelativeHorizontalPosition = 0
                .RelativeVerticalPosition = 0
                .Left = -999995
                .Top = -999995
            End With
            .ActiveWindow.ActivePane.View.SeekView = 0
            '-------------------------------------------------------------'�����ļ�Ϊpdf
            newdc.ExportAsFixedFormat OutputFileName:= _
            strx, ExportFormat:=17, _
            OpenAfterExport:=True, OptimizeFor:=0, Range:= _
            0, from:=1, To:=1, item:=0, _
            IncludeDocProps:=True, KeepIRM:=True, CreateBookmarks:= _
            0, DocStructureTags:=True, BitmapMissingFonts:= _
            True, UseISO19005_1:=False
            newdc.Close savechanges:=False
        End With
        wordapp.Quit
        Set newdc = Nothing
        Set wordapp = Nothing
    End If
    Exit Sub
ErrHandle:
    If Err.Number = 5408 And Errcount < 2 Then
        Err.Clear
        Errcount = Errcount + 1
        If cmCode <> 1 Then
            Resume Renterpassword1 '���������Ҫ����������word,�������,��������3�ε�
        Else
            Resume Renterpassword2
        End If
    End If
100
    Set newdc = Nothing
    wordapp.Quit
    Set wordapp = Nothing
    Err.Clear
End Sub

Function PDFCheckPasswordIsOK(ByVal FilePath As String) As Boolean '�ж�pdf�ļ��Ƿ�������
    Dim objAcroApp As New Acrobat.acroApp
    Dim objAcroPDDoc As New Acrobat.AcroPDDoc
    Dim objJSO As Object
    
    On Error Resume Next
    PDFCheckPasswordIsOK = False
    objAcroPDDoc.Open FilePath
    Set objJSO = objAcroPDDoc.GetJSObject
    If Err.Number > 0 Then Err.Clear
    If IsNull(objJSO.securityHandler) = False Then PDFCheckPasswordIsOK = True '���û������ͻ���ִ���
    If Err.Number > 0 Then Err.Clear: PDFCheckPasswordIsOK = False
    objAcroPDDoc.Close
    objAcroApp.Exit
    Set objAcroPDDoc = Nothing
    Set objAcroApp = Nothing
    Set objJSO = Nothing
End Function

Sub OptimizePDF(ByVal FilePath As String) '�Բ��ֵ�pdf�����Ż�
    Dim objAcroApp As New Acrobat.acroApp
    Dim objAcroPDDoc As New Acrobat.AcroPDDoc
    objAcroPDDoc.Open FilePath
    objAcroPDDoc.Save PDSaveFull + PDSaveLinearized + PDSaveCollectGarbage, FilePath
    objAcroPDDoc.Close
    objAcroApp.Exit
    Set objAcroPDDoc = Nothing
    Set objAcroApp = Nothing
End Sub

'wdReplaceAll 2
'�滻����ƥ����
'wdReplaceNone 0
'���滻�κ�ƥ����
'wdReplaceOne 1
'�滻�����ĵ�һ��ƥ����
Sub Word_Replace_Punctuation(ByVal FilePath As String) '--------------------�滻��word�ĵ��е������ַ�
    Dim Cshape As Variant, Eshape As Variant, i As Byte
    Dim wd As Object
    Dim fRng As Object
    '���ţ���ţ��ٺţ��������ֺţ�ð�ţ������ţ����ţ�̾�ţ������ţ�˫���ţ��ʺţ����ۺţ����غţ����غ�(���ķ��ţ�
    '"[\u3002|\uff1f|\uff01|\uff0c|\u3001|\uff1b|\uff1a|\u201c|\u201d|\u2018|\u2019|\uff08|\uff09|\u300a|\u300b|\u3008|\u3009|\u3010|\u3011|\u300e|\u300f|\u300c|\u300d|\ufe43|\ufe44|\u3014|\u3015|\u2026|\u2014|\uff5e|\ufe4f|\uffe5]"
    Cshape = Array(65292, 12290, 12289, 12304, 12305, 65307, 65306, 8216, 8217, 65288, 65289, 65281, 12298, 12299, 8220, 8221, 65311, 8212, 8226, 12539)
    Eshape = Array(44, 46, 44, 91, 93, 59, 58, 39, 39, 40, 41, 33, 60, 62, 34, 34, 63, 45, 45, 45)
    Set wd = GetObject(FilePath)
    Set fRng = wd.Content
    For i = 0 To 19
        With fRng.Find
            .Text = ChrW(Cshape(i))
            .MatchByte = True 'True if Microsoft Word distinguishes between full-width and half-width letters or characters during a search. Read/write Boolean
            .Replacement.Text = ChrW(Eshape(i))
            .Format = False
            .Execute Replace:=2
        End With
    Next
    wd.Close savechanges:=True
    Set wd = Nothing
    Set fRng = Nothing
End Sub
'https://docs.microsoft.com/zh-cn/office/vba/Language/Reference/User-Interface-Help/chr-function
'https://www.cnblogs.com/jiading/p/11615329.html
'ȫ�ǰ�ǵ�����������ص�����ʵ���϶���ȫ�ǣ� ��ǡ�ȫ�ǵĲ�����Ҫ������Ӣ�ģ����֣���ص�������
Function FindvbWide(ByVal strText As String) As String '����ȫ��/�滻Ϊ���
    Dim myreg As Object
    Dim match As Object, Matches As Object
    Dim Patternx As String
    Dim strTemp As String
    Patternx = "[\uFF00-\uFFFF]+"
    Set myreg = CreateObject("VBScript.RegExp") '������ʽ
    With myreg
        .Pattern = Patternx
        .Global = True
        .IgnoreCase = True
        Set Matches = .Execute(strText)
        For Each match In Matches
            strTemp = StrConv(match.Value, vbNarrow) '����ȫ��תΪ���
            strText = .Replace(strText, strTemp)
        Next
    End With
    FindvbWide = strText
    Set Matches = Nothing
End Function

Option Explicit

Sub Word_Test()
    Dim FilePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False   '��ѡ��
        .Filters.Clear   '����ļ�������
        .Filters.Add "Excel Files", "*.doc;*.docx"
        If .Show = -1 Then
            If .SelectedItems.Count = 1 Then FilePath = .SelectedItems(1) Else Exit Sub
        Else
            Exit Sub
        End If
    End With
    Dim wd As Object
    Set wd = CreateObject(FilePath) 'ʹ��ǰ����رմ��ĵ��� 'ThisWorkbook.Path & "\test.docx"
    Dim p As Object
    ThisWorkbook.Application.ScreenUpdating = False
    For Each p In wd.Paragraphs
        If p.Range.Information(12) = False Then '�����䲻�ڱ��ķ�Χ���ݲ�ִ�ж���
            If InStr(p.Style, "����") > 0 Then
                p.LineSpacingRule = 1
                p.CharacterUnitFirstLineIndent = 2
                With p.Range
                    .Font.Size = 12
                    .Font.Name = "����"
                End With
            End If
        End If
    Next
    Dim pic As Object
    For Each pic In wd.Shapes
        pic.ConvertToInlineShape 'ͼƬתΪǶ��ʽ
    Next
    Dim ilshape As Object
    For Each ilshape In wd.InlineShapes
        ilshape.Range.ParagraphFormat.Alignment = 1 'ͼƬ���ж���
    Next
    Dim t As Object
    Dim r As Object
    Dim temprange As Object
    Dim i As Integer, k As Integer
    k = wd.Tables.Count
    For i = 1 To k
        DoEvents
        Set t = wd.Tables(i)
        Set temprange = wd.Range(t.Range.End, t.Range.End) '���������һ�в������
        temprange.InsertAfter Chr(13)
        Set r = t.Range
        With r.Font
            .NameFarEast = "����"
            .NameAscii = "����"
            .NameOther = "����"
            .Name = "����"
            .Size = 10.5
        End With
        With r.ParagraphFormat
            .SpaceBefore = 1
            .SpaceBeforeAuto = False
            .SpaceAfterAuto = False
            .SpaceAfter = 1
            .LineSpacingRule = 0
            .Alignment = 3
            .WidowControl = False
            .KeepWithNext = False
            .KeepTogether = False
            .PageBreakBefore = False
            .NoLineNumber = False
            .Hyphenation = True
            .OutlineLevel = 10
            .TextboxTightWrap = 0
            .CollapsedByDefault = False
            .AutoAdjustRightIndent = True
            .DisableLineHeightGrid = False
            .FarEastLineBreakControl = True
            .HangingPunctuation = True   'Ϊָ���������� true �����㡣 �������ĳЩָ���Ķ�������Ϊ True ��������Խ����� wdUndefined �� ��/д Long����ʾ��ʵ�ֵĹ����ǣ������ĵ���һ�εı���������߽硣
            .HalfWidthPunctuationOnTopOfLine = False
            .AddSpaceBetweenFarEastAndAlpha = True
            .AddSpaceBetweenFarEastAndDigit = True
            .BaselineAlignment = 4
            .CharacterUnitLeftIndent = 0
            .CharacterUnitRightIndent = 0
            .CharacterUnitFirstLineIndent = 0
            .LineUnitBefore = 0.2
            .LineUnitAfter = 0.2
            .FirstLineIndent = wd.Application.CentimetersToPoints(0)
            .LeftIndent = wd.Application.CentimetersToPoints(0)
            .RightIndent = wd.Application.CentimetersToPoints(0)
        End With
    Next
    Set t = Nothing
    Set r = Nothing
    ThisWorkbook.Application.ScreenUpdating = True
    MsgBox "�������"
    wd.Close savechanges:=True
    Set wd = Nothing
End Sub

Function Character_Case(ByVal strText As String, ByVal iMode As Byte) As String '������̬�Ĵ�Сд����
                                                                                ' ����ĸ��Сд, ȫ����Сд,.....
    Dim strTemp As String
    Select Case iMode
        Case 1: Character_Case = LCase$(strText)
        Case 2: Character_Case = UCase$(strText)
        Case 3: Character_Case = ThisWorkbook.Application.WorksheetFunction.Proper(strText)
        Case 4:
        strTemp = Left$(strTemp, 1)
        Left$(Character_Case, 1) = LCase$(strTemp)
        Case 5:
        strTemp = Left$(strTemp, 1)
        Left$(Character_Case, 1) = UCase$(strTemp)
        Case 6:
        Dim x As Variant
        x = Split(strText, ".", -1, vbBinaryCompare)
        '-----�ӵݹ�
    End Select
End Function

