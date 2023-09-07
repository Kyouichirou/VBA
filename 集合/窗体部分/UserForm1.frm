VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "二维码"
   ClientHeight    =   9555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'https://docs.microsoft.com/zh-cn/windows/win32/api/stringapiset/nf-stringapiset-widechartomultibyte
'https://www.cnblogs.com/ranjiewen/p/5770639.html
Private Declare Function WideCharToMultiByte Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpWideCharStr As Any, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByRef lpDefaultChar As Any, ByRef lpUsedDefaultChar As Any) As Long
Private Const CP_UTF8 As Long = 65001
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Sub QRGenerator() '生成二维码
    Dim arr() As Byte, k As Integer
    Dim strx As String
    Dim i As Long, m As Long
    Dim obj As New clsQRCode
    
    With Me
        k = .ComboBox4.ListIndex
        Select Case k
            Case 1 '-------------------------utf-8(默认使用此项)
                strx = QRtextCN
                m = Len(strx)
                i = m * 3 + 64
                ReDim arr(i)
                m = WideCharToMultiByte(CP_UTF8, 0, ByVal StrPtr(strx), m, arr(0), i, ByVal 0, ByVal 0)
            Case Else '---------------------------------------------------------------------------------------ansi
                strx = StrConv(QRtextCN, vbFromUnicode) '-----------------------------------------------------------https://www.cnblogs.com/whchensir/p/4129345.html
                arr = strx
                m = LenB(strx)
        End Select
        Set .Image1.Picture = obj.Encode(arr, m, .ComboBox1.ListIndex + 1, .ComboBox5.ListIndex + 1, .ComboBox3.ListIndex) '(初始值 0-2-(-1))
    End With
End Sub

Private Sub CommandButton1_Click() '保存二维码为图片
    Dim ph As Long, pw As Long, Py As Long, Px As Long, Px1 As Long
    Dim i As Integer, j As Integer, k As Integer
    Dim Rects As RECT, ExecuteValue As Boolean
    Dim MousePoint As POINTAPI, Folderpath As String
    Dim objShell As Object, hDC As Long
    Dim objFolder As Object, hwnd As Long
    
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder(0, "选择保存位置", 0, 0)
    If objFolder Is Nothing Or objShell Is Nothing Then MsgBox "创建对象失败", vbCritical, "Warning": Exit Sub
    If Not objFolder Is Nothing Then        '** 用户选择了文件夹
        Folderpath = objFolder.self.Path & "\"
    Else
        Set objFolder = Nothing
        Set objShell = Nothing
        Exit Sub
    End If
    hwnd& = FindWindow(vbNullString, Me.Caption)
    hDC = GetDC(hwnd)
    ExecuteValue = GetDesktopWindowRect(hwnd, Rects, MousePoint)
    Get_System_Metrics Px, Py
    With Me
        If Len(.TextBox3.Value) = 0 Then
            i = 10
        Else
            i = .TextBox3.Value
        End If
        If Len(.TextBox4.Value) = 0 Then
            j = 15
        Else
            j = .TextBox4.Value
        End If
        If Len(.TextBox5.Value) = 0 Then
            k = 5
        Else
            k = .TextBox5.Value
        End If
        If i > 255 Then i = 10
        If j > 255 Then j = 15
        If k > 255 Then k = 5
        Px1 = ((Px - (Rects.Left + (Px - Rects.Right))) - .Width * 1.333) * 13
        'Py = Rects.Top + (Py - (Rects.Top + (Py - Rects.Bottom)) - Me.Height * 1.333) * 100 + Image1.Top * 1.333
        Py = Py - (Py - Rects.Bottom) - (.Height - .Image1.Top) * 1.333 + Px1 * 10
        Px = Rects.Left + Px1 + .Image1.Left * 1.333
        ph = Int(.Image1.Height * 1.333) - k * 2
        pw = Int(.Image1.Width * 1.333) - k * 2
        Py = Py + j + k
        Px = Px + i + k
    End With
    ScreenShot Folderpath & Format(Now, "yyyymmddhhmmss") & ".jpg", ph, pw, Py, Px '截图 保存路径,图像高度,图像宽度,屏幕上边距,屏幕左边距
    'SavePicture Image1.Picture, TextBox2.Text & "\" & Format(Now, "yyyymmddhhmmss") & ".jpg"  '通过savepicture直接保存控件上的图片(保存下来的图片非常小)
    Set objFolder = Nothing
    Set objShell = Nothing
End Sub

Private Sub CommandButton2_Click()
    If Len(QRtextCN) = 0 Then Exit Sub
    QRGenerator
End Sub

Private Sub TextBox3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode > 105 Or (KeyCode < 96 And KeyCode <> 8 And KeyCode <> 13 And KeyCode <> 46) Then KeyCode = 0
End Sub

Private Sub TextBox3_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    With Me.TextBox3
        If Len(.Text) = 0 Then
            .Text = 0
        ElseIf Mid(.Text, 1, 1) = "0" Then
            .Text = Mid(.Text, 2, 3)
        End If
    End With
End Sub

Private Sub TextBox4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer) '----------------控制键盘输入的内容
    If KeyCode > 105 Or (KeyCode < 96 And KeyCode <> 8 And KeyCode <> 13 And KeyCode <> 46) Then KeyCode = 0
End Sub

Private Sub TextBox4_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    With Me.TextBox4
        If Len(.Text) = 0 Then
            .Text = 0
        ElseIf Mid(.Text, 1, 1) = "0" Then
            .Text = Mid(.Text, 2, 3)
        End If
    End With
End Sub

Private Sub TextBox5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode > 105 Or (KeyCode < 96 And KeyCode <> 8 And KeyCode <> 13 And KeyCode <> 46) Then KeyCode = 0
End Sub

Private Sub TextBox5_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    With Me.TextBox5
        If Len(.Text) = 0 Then
            .Text = 0
        ElseIf Mid(.Text, 1, 1) = "0" Then
            .Text = Mid(.Text, 2, 3)
        End If
    End With
End Sub

Private Sub UserForm_Initialize()
    Dim i As Long
    If Statisticsx = 1 Then Exit Sub
    With Me
        For i = 1 To 40
            .ComboBox1.AddItem i
        Next i
        .ComboBox5.List = Array("L - 7%", "M - 15%", "Q - 25%", "H - 30%")
        For i = 0 To 7
            .ComboBox3.AddItem CStr(i)
        Next i
        .ComboBox4.List = Array("ANSI", "UTF-8")
        With .Label8
            .Left = Image1.Left + i
            .Top = Image1.Top + i
            .Height = Image1.Height - i * 2
            .Width = Image1.Width - i * 2
        End With
        .TextBox3.SetFocus
    End With
    If Len(QRtextCN) > 0 Then QRGenerator
End Sub
