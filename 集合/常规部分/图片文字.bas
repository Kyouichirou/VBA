Attribute VB_Name = "图片文字"
Option Explicit

Private Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function drawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, ByVal outputbuf As Long) As Long
Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As Long, ByVal FileName As Long, clsidEncoder As GUID, encoderParams As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, id As GUID) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hPal As Long, BITMAP As Long) As Long

Private Const LOGPIXELSX = 88
Private Const LF_FACESIZE = 32
Private Const TRANSPARENT = 1
Private Const LR_COPYRETURNORG = &H4
Private Const IMAGE_BITMAP = 0
Private Const DEFAULT_CHARSET = 1
Private Const DT_WORDBREAK = &H10

Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * LF_FACESIZE
End Type

Sub DrawTextOnPicture(sSourceFileName As String, STargetFileName As String, sText As String, Optional lTextColor As Long = 0, Optional TextFont As stdole.StdFont, _
                        Optional JPGquality As Integer = 100, Optional AutoWrap As Boolean = True)
    Dim oPic As StdPicture
    Dim hBmp As Long, hBmpPrev As Long
    Dim hDCScreen As Long, hDCmem As Long
    Dim hFont As Long, hFontPrev As Long
    Dim rc As RECT
    Dim lPixelsPerInch As Long, hBrushPrev As Long, hBrush As Long
    Dim lf As LOGFONT
    
    On Error Resume Next
    Set oPic = LoadPicture(sSourceFileName)
    '---------------------http://demon.tw/programming/vbs-loadpicture.html
    On Error GoTo 0
    If oPic Is Nothing Then
        MsgBox "不受支持格式，仅支持bmp,jpg(jpeg), ico等图片", vbCritical, "Tips"
        '-----------------可以由 LoadPicture 识别的图形格式有位图文件 (.bmp)、图标文件 (.ico)、行程编码文件 (.rle)、图元文件 (.wmf)、增强型图元文件 (.emf)、GIF (.gif) 文件和 JPEG (.jpg) 文件。
        Exit Sub
    End If
    hBmp = CopyImage(oPic.handle, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)
    hDCScreen = GetDC(0)
    lPixelsPerInch = GetDeviceCaps(hDCScreen, LOGPIXELSX)
    hDCmem = CreateCompatibleDC(hDCScreen)
    ReleaseDC 0, hDCScreen

    hBmpPrev = SelectObject(hDCmem, hBmp)
    hBrushPrev = SelectObject(hDCmem, hBrush)
    rc.Right = oPic.Width * lPixelsPerInch / 2540
    rc.Bottom = oPic.Height * lPixelsPerInch / 2540
    SetBkMode hDCmem, TRANSPARENT
    SetTextColor hDCmem, lTextColor
    If Not TextFont Is Nothing Then
       lf.lfHeight = -MulDiv(TextFont.Size, lPixelsPerInch, 72)
       lf.lfCharSet = DEFAULT_CHARSET
       lf.lfFaceName = TextFont.Name & Chr(0)
       lf.lfItalic = TextFont.Italic
       lf.lfUnderline = TextFont.Underline
       lf.lfStrikeOut = TextFont.Strikethrough
       lf.lfWeight = IIf(TextFont.Bold, 900, 0)
       hFont = CreateFontIndirect(lf)
       hFontPrev = SelectObject(hDCmem, hFont)
       drawText hDCmem, sText & Chr(0), -1, rc, IIf(AutoWrap, DT_WORDBREAK, 0)
       SelectObject hDCmem, hFontPrev
       DeleteObject hFont
    Else
        drawText hDCmem, sText & Chr(0), -1, rc, IIf(AutoWrap, DT_WORDBREAK, 0)
    End If
    SelectObject hDCmem, hBmpPrev
    DeleteDC hDCmem
    Call SavehBitmapToJPGFile(hBmp, STargetFileName, JPGquality)
    DeleteObject hBmp
End Sub

Sub AddCharToPicture(ByVal FilePath As String, ByVal outputfile As String, ByVal textchar As String) '添加文字到图片
    Dim oFont As stdole.StdFont
    '--------------------------------https://docs.microsoft.com/zh-cn/dotnet/framework/additional-apis/stdole.stdfont
    Set oFont = New StdFont
    With oFont
        .Name = "宋体"
        .Size = 30
    End With
    DrawTextOnPicture FilePath, outputfile, textchar
End Sub

Function SavehBitmapToJPGFile(ByVal hBmp As Long, ByVal sFilename As String, Optional ByVal quality As Byte = 80) As Integer
    Dim tSI As GdiplusStartupInput
    Dim lRes As Long
    Dim lGDIP As Long
    Dim lBitmap As Long
    '---------------------------http://blog.chinaunix.net/uid-31439230-id-5764435.html
    '---------------------------https://blog.csdn.net/happyboy200032/article/details/6071745
    tSI.GdiplusVersion = 1
    lRes = GdiplusStartup(lGDIP, tSI, 0)
    If lRes = 0 Then
        lRes = GdipCreateBitmapFromHBITMAP(hBmp, 0, lBitmap)
        If lRes = 0 Then
            Dim tJpgEncoder As GUID
            Dim tParams As EncoderParameters
            CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
            tParams.Count = 1
            With tParams.Parameter
                CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID
                .NumberOfValues = 1
                .type = 4
                .Value = VarPtr(quality)
            End With
            lRes = GdipSaveImageToFile(lBitmap, StrPtr(sFilename), tJpgEncoder, tParams)
            If lRes = 0 Then
                SavehBitmapToJPGFile = 0
            Else
                SavehBitmapToJPGFile = 1
            End If
            GdipDisposeImage lBitmap
        End If
        GdiplusShutdown lGDIP
    End If
End Function

