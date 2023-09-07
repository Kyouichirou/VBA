Attribute VB_Name = "图像处理"
Option Explicit
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long '获取图片的尺寸

Private Type BITMAP
    bmType   As Long
    bmWidth   As Long
    bmHeight   As Long
    bmWidthBytes   As Long
    bmPlanes   As Integer
    bmBitsPixel   As Integer
    bmBits   As Long
End Type

Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As Long
End Type
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, _
ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, ByRef lpBits As Any, _
ByRef lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As GUID, ByVal fPictureOwnsHandle As Long, ipic As IUnknown) As Long

Function BitmapToPicture(ByVal hBmp As Long, ByVal fPictureOwnsHandle As Long) As StdPicture

    If (hBmp = 0) Then Exit Function
    Dim oNewPic As IUnknown, tPicConv As PictDesc, IGuid As GUID
    With tPicConv
        .cbSizeofStruct = Len(tPicConv)
        .picType = 1 'vbPicTypeBitmap
        .hImage = hBmp
    End With
    With IGuid
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    OleCreatePictureIndirect tPicConv, IGuid, fPictureOwnsHandle, oNewPic
    Set BitmapToPicture = oNewPic
End Function

Function ByteArrayToPicture(ByVal lp As Long, ByVal nWidth As Long, ByVal nHeight As Long, Optional ByVal nLeftPadding As Long, _
Optional ByVal nTopPadding As Long, Optional ByVal nRightPadding As Long, Optional ByVal nBottomPadding As Long) As StdPicture
    Dim tBMI As BITMAPINFO
    Dim h As Long, hDC As Long, hBmp As Long
    Dim hbr As Long
    Dim r As RECT

    With tBMI.bmiHeader
        .biSize = 40&
        .biWidth = nWidth
        .biHeight = -nHeight
        .biPlanes = 1
        .biBitCount = 8
        .biSizeImage = nWidth * nHeight
        .biClrUsed = 256
    End With
    tBMI.bmiColors(0) = &HFFFFFF
    tBMI.bmiColors(2) = &H808080
    h = GetDC(0)
    hDC = CreateCompatibleDC(h)
    r.Right = nWidth + nLeftPadding + nRightPadding
    r.Bottom = nHeight + nTopPadding + nBottomPadding
    hBmp = CreateCompatibleBitmap(h, r.Right, r.Bottom)
    hBmp = SelectObject(hDC, hBmp)
    hbr = CreateSolidBrush(vbWhite)
    FillRect hDC, r, hbr
    DeleteObject hbr
    StretchDIBits hDC, nLeftPadding, nTopPadding, nWidth, nHeight, 0, 0, nWidth, nHeight, ByVal lp, tBMI, 0, 13369376 'vbSrcCopy
    hBmp = SelectObject(hDC, hBmp)
    DeleteDC hDC
    ReleaseDC 0, h
    Set ByteArrayToPicture = BitmapToPicture(hBmp, 1)
End Function

Sub WIAImage(ByVal FilePath As String, Optional ByVal heightx As Integer = 256, Optional ByVal widthx As Integer = 256) '缩小图片
    Dim outputfilepath As String, strx As String
    Dim strx1 As String, i As Byte, strx2 As String, strx3 As String, strx4 As String
    Dim img As Object, ip As Object
    'https://www.cnblogs.com/bjguanmu/articles/7559800.html
    'https://docs.microsoft.com/en-us/previous-versions/windows/desktop/wiaaut/-wiaaut-imagefile
    'https://docs.microsoft.com/en-us/previous-versions/windows/desktop/wiaaut/-wiaaut-imageprocess
    Set img = CreateObject("WIA.ImageFile")
    Set ip = CreateObject("WIA.ImageProcess")
    img.LoadFile FilePath
    With ip
        .Filters.Add ip.FilterInfos("Scale").FilterID
        .Filters(1).Properties("MaximumWidth") = widthx '宽度
        .Filters(1).Properties("MaximumHeight") = heightx '高度
    End With
    Set img = ip.Apply(img)
    i = InStrRev(FilePath, "\")
    strx = Left$(FilePath, i)   '文件位置
    strx1 = Right$(FilePath, Len(FilePath) - i) '文件名
    strx2 = Left$(strx1, InStrRev(strx1, "."))
    strx3 = Right$(strx1, Len(strx1) - Len(strx2)) '扩展名
    strx4 = Format(Now, "yyyymmddhhmmss")
    strx2 = Left$(strx2, Len(strx2) - 1) & strx4 & "."
    strx1 = strx2 & strx3 '------------------------------新的文件名
    outputfilepath = strx & strx1
    img.SaveFile outputfilepath
    Set img = Nothing
    Set ip = Nothing
End Sub

Sub ImageConvert(ByVal FilePath As String, Optional ByVal cmCode As Byte = 4) '图片格式转换 '默认转换为jpeg格式 '可用于转换下载书籍封面图片假如不是image控件支持的
    Dim ip As Object, img As Object, filex As String
    Dim target As String, strx As String
    Dim strx1 As String, i As Byte, strx2 As String, strx3 As String, strx4 As String
    'https://docs.microsoft.com/en-us/previous-versions/windows/desktop/wiaaut/-wiaaut-howto-use-filters
    'https://docs.microsoft.com/en-us/previous-versions/windows/desktop/wiaaut/-wiaaut-consts-formatid
    'https://docs.microsoft.com/en-us/previous-versions/windows/desktop/wiaaut/-wiaaut-imagefile
    Select Case cmCode  '----------------------------指定转换文件的类型
        Case 1: target = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}": filex = "bmp" 'bmp
        Case 2: target = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}": filex = "png" 'png
        Case 3: target = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}": filex = "gif" 'gif
        Case 4: target = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}": filex = "jpeg" 'jpeg(这个格式loadpicture也支持)
        Case 5: target = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}": filex = "tiff" 'tiff
        Case Else: target = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}": filex = "jpeg" 'jpeg
    End Select
    i = InStrRev(FilePath, "\")
    strx = Left$(FilePath, i)   '文件位置
    strx1 = Right$(FilePath, Len(FilePath) - i) '文件名
    strx2 = Left$(strx1, InStrRev(strx1, "."))
    strx3 = Right$(strx1, Len(strx1) - Len(strx2)) '扩展名
    strx3 = LCase(strx3)
    If filex = strx3 Then Exit Sub       '如果扩展文件名和需要转换的文件名相同
    strx4 = Format(Now, "yyyymmddhhmmss")
    strx2 = Left$(strx2, Len(strx2) - 1) & strx4 & "."
    strx1 = strx2 & filex '------------------------------新的文件名
    outputfilepath = strx & strx1
    Set img = CreateObject("WIA.ImageFile")
    Set ip = CreateObject("WIA.ImageProcess")
    img.LoadFile FilePath
    ip.Filters.Add ip.FilterInfos("Convert").FilterID
'    target = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
    ip.Filters(1).Properties("FormatID").Value = target
    Set img = ip.Apply(img)
    img.SaveFile outputfilepath
    Set img = Nothing
    Set ip = Nothing
End Sub

Function ObtainPicWH(ByVal FilePath As String) As Long() '获取图片的尺寸
    Dim bm As BITMAP
    Dim picPicture As IPictureDisp
    Dim arr(1) As Long
    
    Set picPicture = stdole.LoadPicture(FilePath)
    GetObjectAPI picPicture, Len(bm), bm
    arr(0) = bm.bmWidth
    arr(1) = bm.bmHeight
    ObtainPicWH = arr
    Set picPicture = Nothing
End Function
