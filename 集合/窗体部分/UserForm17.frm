VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm17 
   Caption         =   "词典"
   ClientHeight    =   10305
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6780
   OleObjectBlob   =   "UserForm17.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
Private Const SgUrl As String = "https://fanyi.sogou.com/"
Private Const GeUrl As String = "https://translate.google.cn/"
Private Const UAgent As String = "User-Agent:Mozilla/5.0 (iPhone; CPU iPhone OS 9_1 like Mac OS X) AppleWebKit/601.1.46 (KHTML, like Gecko) Version/9.0 Mobile/13B143 Safari/601.1"

Private Sub AddIcon()
    Dim hwnd As Long
    Dim lngRet As Long
    Dim hIcon As Long
    'hIcon = Image1.Picture.Handle
    'hWnd = FindWindow(vbNullString, Me.Caption)
    lngRet = SendMessage(hwnd, WM_SETICON, ICON_SMALL, ByVal hIcon)
    'lngRet = SendMessage(hWnd, WM_SETICON, ICON_BIG, ByVal hIcon)
    lngRet = DrawMenuBar(hwnd)
End Sub

Private Sub AddMinimiseButton()
    Dim hwnd As Long
    hwnd = GetActiveWindow
    Call SetWindowLong(hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) Or WS_MINIMIZEBOX)
    Call SetWindowPos(hwnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Private Sub AppTasklist(myForm)
    'Add this userform into the Task bar
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

Private Sub CommandButton1_Click()
    Dim url As String
    With Me
        If .CommandButton1.Caption = "Google" Then
            url = GeUrl
        Else
            url = SgUrl
        End If
        .WebBrowser2.Navigate url
    End With
End Sub

Private Sub CommandButton2_Click() '关闭主体窗口
    If UF3Show = 1 Then
        ThisWorkbook.Application.ScreenUpdating = False
        Me.Hide          '两个窗体同时存在时, 控制下一层的窗体,需要先隐藏或者卸载上层
        UF3Show = 3
        UserForm3.Hide
        ThisWorkbook.Application.ScreenUpdating = True
        Me.Show
    End If
End Sub

Private Sub CommandButton3_Click() '返回主窗口
    Unload UserForm17
    If UF3Show <> 1 Then UserForm3.Show
End Sub

Private Sub UserForm_Activate()
    If Statisticsx = 1 Then Exit Sub
    AddIcon    '添加图标
    AddMinimiseButton   '添加按钮
    AppTasklist Me    '添加任务栏
End Sub

Private Sub UserForm_Initialize()
    If Statisticsx = 1 Then Exit Sub
    With Me.WebBrowser2
        .Silent = True
        .Navigate SgUrl, , , , UAgent
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) '在隐藏后,在重新出现,会出现点击无效的问题
    If Statisticsx = 1 Then Exit Sub
    UserForm17.Hide '注意这里不能使用unload me, 否则会出现无法卸载本体的问题
    If UF3Show = 3 Then UserForm3.Show
End Sub

Private Sub WebBrowser2_NavigateComplete2(ByVal pDisp As Object, url As Variant)
    Dim strx As String
    With Me
        strx = .WebBrowser2.LocationURL
        If InStr(strx, "translate.google.cn") > 0 Then
            .CommandButton1.Caption = "Sogo"
        Else
            .CommandButton1.Caption = "Google"
        End If
    End With
End Sub

