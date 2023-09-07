Attribute VB_Name = "截图"
Option Explicit
Type RECT         '截图/图像处理两个模块都需要调用/不要使用private
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type POINTAPI
    x As Long
    y As Long
End Type

Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Type EncoderParameter '截图/图像处理两个模块都需要调用
    GUID As GUID
    NumberOfValues As Long
    type As Long
    Value As Long
End Type

Type EncoderParameters '截图/图像处理两个模块都需要调用
    Count As Long
    Parameter As EncoderParameter
End Type

Type GdiplusStartupInput '截图/图像处理两个模块都需要调用
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type
'-----------------------------------------------------------------https://docs.microsoft.com/en-us/windows/win32/gdiplus/-gdiplus-image-flat
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'-------------------------------------------------https://docs.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-createcompatiblebitmap
'-------------------------------------------------http://blog.sina.com.cn/s/blog_4ad042e50102e3a9.html
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
                                    ByVal x As Long, ByVal y _
                                    As Long, ByVal nWidth As Long, _
                                    ByVal nHeight As Long, _
                                    ByVal hSrcDC As Long, _
                                    ByVal XSrc As Long, _
                                    ByVal YSrc As Long, _
                                    ByVal dwRop As Long) As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
'----------------------------------------------------------------剪切板控制
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Integer) As Long
Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, ByVal outputbuf As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hPal As Long, BITMAP As Long) As Long
'---------------------------------------------------------------------------------------------https://www.cnblogs.com/liuzhaoyzz/p/4035045.html
Private Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, id As GUID) As Long
'------------------------------------------------------------------------------------------https://docs.microsoft.com/zh-cn/windows/win32/api/combaseapi/nf-combaseapi-stringfromclsid
Private Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As Long, ByVal FileName As Long, clsidEncoder As GUID, encoderParams As Any) As Long
Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long
Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
'-----------------------------------------------------------------------------------------------------https://www.cnblogs.com/BlackList-Sakura/p/6682156.html
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long '---------------https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getsystemmetrics
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long '-------------------https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getcursorpos
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
'Declare Sub GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                    lpRect As RECT)
Private Const SM_CXSCREEN = 0 'Width of screen
Private Const SM_CYSCREEN = 1 'Height of screen
Global Const SRCCOPY = &HCC0020
Global Const CF_BITMAP = 2

Function GetDesktopWindowRect(hwnd As Long, Rct As RECT, MousePos As POINTAPI) As Boolean
    Dim Execute As Integer
    Execute = GetWindowRect(hwnd, Rct)
    GetDesktopWindowRect = IIf(Execute = 0, False, True)
    GetCursorPos MousePos
End Function
                
Sub Get_System_Metrics(ByRef XVal As Long, ByRef YVal As Long)
    YVal = GetSystemMetrics(SM_CYSCREEN)
    XVal = GetSystemMetrics(SM_CXSCREEN)
End Sub

Sub ScreenShot(Picpath As String, ph As Long, pw As Long, st As Long, Sl As Long) '截图
   Dim AccessHwnd As Long, DeskHwnd As Long
   Dim hDC As Long
   Dim hDCmem As Long
   Dim RECT As RECT
   Dim junk As Long
   Dim fwidth As Long, fheight As Long
   Dim hBitMap As Long

   DeskHwnd = GetDesktopWindow()
   AccessHwnd = GetActiveWindow()

   Call GetWindowRect(AccessHwnd, RECT)
   fwidth = pw ' UserForm1.Label7.Width ' rect.right - rect.left
   fheight = ph ' UserForm1.Label7.Height ' rect.bottom - rect.top
   hDC = GetDC(DeskHwnd)
   hDCmem = CreateCompatibleDC(hDC)
   hBitMap = CreateCompatibleBitmap(hDC, fwidth, fheight)
   If hBitMap <> 0 Then
      junk = SelectObject(hDCmem, hBitMap)
      junk = BitBlt(hDCmem, 0, 0, fwidth, fheight, hDC, Sl, _
                     st, SRCCOPY)

      junk = OpenClipboard(DeskHwnd)
      junk = EmptyClipboard()
      junk = SetClipboardData(CF_BITMAP, hBitMap)
      junk = CloseClipboard()
   End If
   junk = DeleteDC(hDCmem)
   junk = ReleaseDC(DeskHwnd, hDC)
   Screen2JPG Picpath
End Sub

Function Screen2JPG(ByVal FileName As String, Optional ByVal quality As Byte = 80) As Boolean
    Dim tSI As GdiplusStartupInput
    Dim lRes As Long
    Dim lGDIP As Long
    Dim lBitmap As Long
    Dim hBitMap As Long
    Dim tJpgEncoder As GUID
    Dim tParams As EncoderParameters

    OpenClipboard 0&  '打开剪贴板
    hBitMap = GetClipboardData(CF_BITMAP)  '获取剪贴板中bitmap数据的句柄
    CloseClipboard   '关闭剪贴板
   
    tSI.GdiplusVersion = 1  '初始化 GDI+
    lRes = GdiplusStartup(lGDIP, tSI, 0)
    If lRes = 0 Then
        '---------------从句柄创建 GDI+ 图像
         lRes = GdipCreateBitmapFromHBITMAP(hBitMap, 0, lBitmap)
        If lRes = 0 Then
            
            '------------------------------------初始化解码器的GUID标识
            CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
            '------------------------------------------------------------------------------设置解码器参数
            tParams.Count = 1
                With tParams.Parameter ' Quality
                '-------------------------------得到Quality参数的GUID标识
                CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID
                .NumberOfValues = 1
                .type = 4
                .Value = VarPtr(quality)
            End With
            '---------保存图像
            lRes = GdipSaveImageToFile(lBitmap, StrPtr(FileName), tJpgEncoder, tParams)
            '------------------------销毁GDI+图像
            GdipDisposeImage lBitmap
        End If
        '--------------------销毁 GDI+
        GdiplusShutdown lGDIP
    End If
        Screen2JPG = Not lRes
End Function
