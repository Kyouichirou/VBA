Attribute VB_Name = "文档处理"
Option Explicit
'------------------涉及PDF的处理,打印, 添加水印, 添加图片, 插入页面, 文件加密检查,加密等等的操作,主要依赖于Adobe Acrobat和PDFCreator
'-------------------https://acrobat.adobe.com/us/en/acrobat/pdf-reader.html
'-------------------https://www.pdfforge.org/pdfcreator
'-------------------相关的参数或者方法的参考来源主要是软件提供的SDK文档
'注意事项:(无法读取)
'文件损坏
'文件加密(不同方式的加密),部分加密文件依然可以修改部分的文件信息,或者用文件流的方式向文件写入内容, 而一部分的加密将完全无法修改,强行修改可以造成文件的损坏
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

Sub FilePrint(ByVal strFilePath As String) '打印文件
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
'https://www.pdfforge.org/, 依赖于PDFCreator
'以下代码基于3.3（和谐版本），注意网上现存的vba代码都是基于老旧版本的pdfcreator，已经不适合在新版本上使用
'需要注意，PdfCreator安装位置下的chm manual有很多错误，包括网页的很多内容都是错误的,甚至有一些拼写都出现错误<--->
'http://docs.pdfforge.org/pdfcreator/latest/en/pdfcreator/com-interface/
'http://docs.pdfforge.org/pdfcreator/latest/en/pdfcreator/com-interface/reference/settings/ '参数以这个网页的内容为基准
Sub EncryptPDF(ByVal FilePath As String, ByVal outputfilepath As String, Optional ByVal owerpassword As String = "password", _
Optional ByVal userpassword As String = "password", Optional ByVal cmCode As Byte = 0, Optional ByVal securitylevel As Byte = 0) '注意无法直接添加非pdf文件到队列
    Dim pdfobj As New PDFCreator_COM.PdfCreatorObj
    Dim pdfq As New PDFCreator_COM.Queue
    Dim job As Object
    Dim Encrypt As String
    'Rc40Bit, Rc128Bit, Aes128Bit, Aes256Bit '加密方式
    'owerpassword---------限制编辑密码
    'userpassword--------打开文件密码
    On Error GoTo ErrHandle
    Select Case securitylevel
        Case 1: Encrypt = "Rc40Bit"
        Case 2: Encrypt = "Rc128Bit"
        Case 3: Encrypt = "Aes256Bit" '此项仅限于商业版本
        Case Else: Encrypt = "Aes128Bit" '默认使用此项目
    End Select
    
    If pdfobj.IsInstanceRunning = False Then pdfq.Initialize '判断PdfCreator是否处于运行的状态
    pdfobj.AddFileToQueue (FilePath) '添加文件到队列
    Set job = pdfq.NextJob
    With job
        .SetProfileByGuid ("DefaultGuid") '文件格式
        .SetProfileSetting "PdfSettings.Security.Enabled", "true" '安全设置
        .SetProfileSetting "PdfSettings.Security.EncryptionLevel", Encrypt '加密方式
        If cmCode = 1 Then
            .SetProfileSetting "PdfSettings.Security.RequireUserPassword", "true"
            .SetProfileSetting "PdfSettings.Security.UserPassword", userpassword
        Else
            .SetProfileSetting "PdfSettings.Security.OwnerPassword", owerpassword
        End If
        .ConvertTo outputfilepath 'pdf格式’输出新的文件
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

Function WordConvertToPDF(ByVal FilePath As String, ByVal outputfilepath As String) As Boolean '直接将word文件转成pdf，同时加密
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
    wdoc.PrintOut Background:=False '这个属性必须，使用激活后使用
    If pdfobj.IsInstanceRunning = False Then pdfq.Initialize '在f8调试模式下，会出现窗口
    If pdfq.WaitForJob(10) = True Then '判断是否在规定的时间内添加到队列中
        Set printJob = pdfq.NextJob
        With printJob
            .SetProfileByGuid ("DefaultGuid")
            .SetProfileSetting "PdfSettings.Security.Enabled", "true" '安全设置
            .SetProfileSetting "PdfSettings.Security.EncryptionLevel", "Aes128Bit" '加密方式
            .SetProfileSetting "PdfSettings.Security.OwnerPassword", "123"
            .ConvertTo outputfilepath
            If (Not .IsFinished Or Not .IsSuccessful) Then '判断是否有效执行
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
Function addWaterMarktoPDF(ByVal FilePath As String, ByVal watermark As String, Optional ByVal startpage As Integer = 1, Optional ByVal endpage As Integer = 0) As Byte '添加水印到pdf文件，缺陷，添加水印后变大(变化幅度较大）
    Dim jso As Object
    Dim acroApp As Object
    Dim myDocument As Object
    Dim i As Integer, k As Integer
    '另一种添加水印的方法是通过文件 -addWatermarkFromFile ,如果所选的源文件不是pdf文件,adobe会转成pdf文件
    'The device-independent path of the source file to use for the watermark. If the file at this location is not a PDF file, Acrobat attempts to convert the file to a PDF file
    On Error GoTo ErrHandle
    addWaterMarktoPDF = 0
    Set acroApp = CreateObject("AcroExch.App")
    Set myDocument = CreateObject("AcroExch.PDDOc")
    myDocument.Open (FilePath)
    i = myDocument.GetNumPages      '注意下面的页码从0开始算
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
        nEnd:=endpage     '选择开始和结尾的页面， 如果不设置，将默认包括所有的页面
'    myDocument.Save 1, filepath '1 '从测试效果来看, 单一参数选择1的时候效果最好(变化的体积最小)
    myDocument.Save PDSaveFull + PDSaveLinearized + PDSaveCollectGarbage, FilePath '可以同时使用多个参数
ErrHandle:
    addWaterMarktoPDF = 1
    If Err.Number = 448 Then addWaterMarktoPDF = 2: Err.Clear '如果文件被加密,将会出现错误
    Set jso = Nothing
    myDocument.Close
    acroApp.Exit
    Set myDocument = Nothing
    Set acroApp = Nothing
End Function

Sub SavePDFAs(ByVal PDFPath As String, ByVal fileExtension As String) '另存pdf文件为(大文件的转换,谨慎使用)
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

'----------------------------------------------------------密码必须有值,否则会出现word的密码输入界面
Function AddWordPassword(ByVal FilePath As String, Optional ByVal cmCode As Byte = 0, Optional ByVal cmCodex As Byte = 0, _
Optional ByVal opassword As String = "abc", Optional ByVal wpassword As String = "abc") As Byte '检查和添加word打开密码或者编辑密码
    Dim wd As Object
    Dim wdoc As Object, i As Integer
    '-------------- --- word的密码设置有打开时候的密码, 编辑权限的密码
    On Error GoTo ErrHandle
    AddWordPassword = 0
    Set wd = CreateObject("word.application")
    '----------------------------------------https://docs.microsoft.com/zh-cn/office/vba/api/Word.Documents.Open
    Set wdoc = wd.documents.Open(FilePath, PasswordDocument:=opassword, WritePasswordDocument:=wpassword, Visible:=False)  '当文件密友密码的时候,这里的密码设置不影响正常的文件处理
    If wdoc Is Nothing Then GoTo ErrHandle
    If cmCodex = 1 Then
        AddWordPassword = 4
        wdoc.Close
        wd.Quit
        Set wdoc = Nothing
        Set wd = Nothing
        Exit Function 'cmcodex用于表示只检查文件是否有密码,但是不具体执行添加密码
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
    If i = 5408 Or i = 5174 Then '打开保护
        AddWordPassword = 1
    ElseIf i = 4198 Then
        AddWordPassword = 2 '编辑保护
    Else
        AddWordPassword = 3 '其他错误
    End If
    wd.Quit
    Err.Clear
    Set wd = Nothing
    Set wdoc = Nothing
End Function

Sub WordToPDF(ByVal FilePath As String, ByVal cmCode As Byte, Optional ByVal waterprint As String) '将word文档转成pdf(加水印)
    Dim wordapp As Object
    Dim newdc As Object
    Dim strx As String
    Dim Errcount As Byte '用于记录输入密码的错误次数
    '如果是excel另存pdf需要print service启动
    Set wordapp = CreateObject("Word.Application")
    If wordapp Is Nothing Then MsgBox "创建word对象失败": Exit Sub
    On Error GoTo ErrHandle
    strx = Environ("UserProfile") & "\Desktop\" & Split(Right$(FilePath, Len(FilePath) - InStrRev(FilePath, "\")), ".")(0) & ".pdf" '新的文件名
    If cmCode <> 1 Then
'        Set newdc = wordapp.Documents.Add(filepath)
Renterpassword1:
        Set newdc = wordapp.documents.Open(FilePath, ReadOnly:=True, Visible:=False) '设置readonly将可以避免设置有编辑密码,假如设置有打开密码,visible不影响弹出密码输入窗口
        wordapp.Visible = False
        newdc.ExportAsFixedFormat OutputFileName:=strx, ExportFormat:=17, OpenAfterExport:=True '相关的内容细节可已查看word chm文件
        newdc.Close
        wordapp.Quit
        Set newdc = Nothing
        Set wordapp = Nothing
    ElseIf cmCode = 1 Then
        If Len(waterprint) = 0 Then waterprint = ThisWorkbook.Application.Username & "出品" '水印
Renterpassword2:
        Set newdc = wordapp.documents.Open(FilePath, ReadOnly:=True, Visible:=False)
        If newdc Is Nothing Then MsgBox "创建word对象失败": GoTo 100
        wordapp.Visible = False
        With newdc.Application
            .ActiveWindow.ActivePane.View.SeekView = 9   '要将wdSeekCurrentPageHeader这种常量写成值的方法,查找值的方法,在word中创建sub debug.print wdRelativeVerticalPositionMargin(对应的常量)
            .Selection.HeaderFooter.Shapes.AddTextEffect(1, waterprint, "宋体", 36, False, False, 0, 0).Select
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
            '-------------------------------------------------------------'导出文件为pdf
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
            Resume Renterpassword1 '假如出现需要输入打开密码的word,输入错误,允许输入3次的
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

Function PDFCheckPasswordIsOK(ByVal FilePath As String) As Boolean '判断pdf文件是否有密码
    Dim objAcroApp As New Acrobat.acroApp
    Dim objAcroPDDoc As New Acrobat.AcroPDDoc
    Dim objJSO As Object
    
    On Error Resume Next
    PDFCheckPasswordIsOK = False
    objAcroPDDoc.Open FilePath
    Set objJSO = objAcroPDDoc.GetJSObject
    If Err.Number > 0 Then Err.Clear
    If IsNull(objJSO.securityHandler) = False Then PDFCheckPasswordIsOK = True '如果没有密码就会出现错误
    If Err.Number > 0 Then Err.Clear: PDFCheckPasswordIsOK = False
    objAcroPDDoc.Close
    objAcroApp.Exit
    Set objAcroPDDoc = Nothing
    Set objAcroApp = Nothing
    Set objJSO = Nothing
End Function

Sub OptimizePDF(ByVal FilePath As String) '对部分的pdf进行优化
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
'替换所有匹配项
'wdReplaceNone 0
'不替换任何匹配项
'wdReplaceOne 1
'替换遇到的第一个匹配项
Sub Word_Replace_Punctuation(ByVal FilePath As String) '--------------------替换掉word文档中的中文字符
    Dim Cshape As Variant, Eshape As Variant, i As Byte
    Dim wd As Object
    Dim fRng As Object
    '逗号，句号，顿号，【】，分号，冒号，单引号，括号，叹号，书名号，双引号，问号，破折号，着重号，着重号(日文符号）
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
'全角半角的区别，中文相关的内容实际上都是全角， 半角、全角的差异主要体现在英文（数字）相关的内容上
Function FindvbWide(ByVal strText As String) As String '查找全角/替换为半角
    Dim myreg As Object
    Dim match As Object, Matches As Object
    Dim Patternx As String
    Dim strTemp As String
    Patternx = "[\uFF00-\uFFFF]+"
    Set myreg = CreateObject("VBScript.RegExp") '正则表达式
    With myreg
        .Pattern = Patternx
        .Global = True
        .IgnoreCase = True
        Set Matches = .Execute(strText)
        For Each match In Matches
            strTemp = StrConv(match.Value, vbNarrow) '‘将全角转为半角
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
        .AllowMultiSelect = False   '单选择
        .Filters.Clear   '清除文件过滤器
        .Filters.Add "Excel Files", "*.doc;*.docx"
        If .Show = -1 Then
            If .SelectedItems.Count = 1 Then FilePath = .SelectedItems(1) Else Exit Sub
        Else
            Exit Sub
        End If
    End With
    Dim wd As Object
    Set wd = CreateObject(FilePath) '使用前建议关闭此文档先 'ThisWorkbook.Path & "\test.docx"
    Dim p As Object
    ThisWorkbook.Application.ScreenUpdating = False
    For Each p In wd.Paragraphs
        If p.Range.Information(12) = False Then '当段落不在表格的范围内容才执行动作
            If InStr(p.Style, "正文") > 0 Then
                p.LineSpacingRule = 1
                p.CharacterUnitFirstLineIndent = 2
                With p.Range
                    .Font.Size = 12
                    .Font.Name = "宋体"
                End With
            End If
        End If
    Next
    Dim pic As Object
    For Each pic In wd.Shapes
        pic.ConvertToInlineShape '图片转为嵌入式
    Next
    Dim ilshape As Object
    For Each ilshape In wd.InlineShapes
        ilshape.Range.ParagraphFormat.Alignment = 1 '图片居中对齐
    Next
    Dim t As Object
    Dim r As Object
    Dim temprange As Object
    Dim i As Integer, k As Integer
    k = wd.Tables.Count
    For i = 1 To k
        DoEvents
        Set t = wd.Tables(i)
        Set temprange = wd.Range(t.Range.End, t.Range.End) '表格段落最后一行插入空行
        temprange.InsertAfter Chr(13)
        Set r = t.Range
        With r.Font
            .NameFarEast = "宋体"
            .NameAscii = "宋体"
            .NameOther = "宋体"
            .Name = "宋体"
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
            .HangingPunctuation = True   '为指定段落启用 true 如果标点。 如果仅有某些指定的段落设置为 True ，则此属性将返回 wdUndefined 。 读/写 Long。本示例实现的功能是：允许活动文档第一段的标点可以溢出边界。
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
    MsgBox "处理完成"
    wd.Close savechanges:=True
    Set wd = Nothing
End Sub

Function Character_Case(ByVal strText As String, ByVal iMode As Byte) As String '各种形态的大小写设置
                                                                                ' 首字母大小写, 全部大小写,.....
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
        '-----加递归
    End Select
End Function

