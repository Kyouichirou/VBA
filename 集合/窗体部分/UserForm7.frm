VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm7 
   Caption         =   "About Me"
   ClientHeight    =   5460
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9000
   OleObjectBlob   =   "UserForm7.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const ChCode As String = "HLAstaticx"

Private Sub CommandButton1_Click() 'support me
    UserForm10.Show
End Sub

Private Sub Label19_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim textb As Object, strx As String
    
    With Me
        Set textb = .Controls.Add("Forms.TextBox.1", "Text1", False) '以创建临时textbox的方式实现复制内容
        strx = .Label19.Caption
        With textb
            .Text = strx
            .SelStart = 0
            .SelLength = Len(.Text)
            .Copy
        End With
        .Label18.Caption = "邮件地址已复制"
    End With
    Set textb = Nothing
End Sub

Private Sub UserForm_Initialize()
    Dim k As Integer, i As Integer, j As Integer, m As Integer, lc As Integer, bl As Integer, n As Integer, p As Integer
    Dim strx As String, a As Byte, b As Byte, c As Integer
    Dim obj As Object, strx1 As String, objdata As Object
    Dim UF As Object, wd As Object, wdc As Integer, strx2 As String
    
    On Error Resume Next
    Statisticsx = 1 '用于向所有的窗体表示,初始化的请求来自统计,减少窗体初始化的运行加载/卸载
    With ThisWorkbook
        .Application.ScreenUpdating = False
        lc = .VBProject.VBComponents.Count
        For Each obj In .VBProject.VBComponents '计算userform控件的数量
            If obj.Name <> "UserForm7" Then
                If obj.type = 3 Then
                    Set UF = UserForms.Add(obj.Name) '这个过程会初始化窗体,注意某些窗体在初始化或者激活时执行的事件, 是否存在问题,否则会造成错误
                    m = m + UF.Controls.Count
                    Unload UF
                End If
            End If
        Next
        For i = 1 To lc
            With .VBProject.VBComponents.item(i).CodeModule
                p = .CountOfLines '总行数
                k = k + p
                For n = p To 1 Step -1
                    strx = .Lines(n, 1)
                    strx1 = Trim(strx)
                    If strx1 = vbNullString Then
                        bl = bl + 1 '计算空行
                    Else
                        If InStr(1, strx1, Chr(39), vbBinaryCompare) > 0 Then j = j + 1 'chr(39)'单引号
                    End If
                Next
            End With
        Next
         '-----------------计算表格中的控件数量
        a = .Sheets.Count
        For b = 1 To a
            c = c + .Sheets(b).Shapes.Count
        Next
        '-----------------------------------------获取word部分的情况/也可以通过运行docm内的宏来统计
        strx2 = ThisWorkbook.Path & "\LB.docm"
        If fso.fileexists(strx2) = True Then '-------如果直接创建word对象,执行的速度较慢
            SetClipboard ChCode '传递一个值到剪切板,用于控制word文件的执行(open事件)
            If Err.Number > 0 Then Err.Clear
            Set wd = GetObject(strx2) '-有别于获取word对象, 这里可以直接获取到文档的对象(文档处于关闭的状态),不会出现错误
            If Err.Number > 0 Then Set wd = CreateObject(strx2): Err.Clear
            'wd.Application.Run "VBprojectStatic" '-----直接运行word内的sub进程
            '--------------------------------https://docs.microsoft.com/zh-cn/office/vba/api/word.application.run
            With wd.VBProject.VBComponents
                wdc = .Count
                lc = lc + wdc
                For i = 1 To wdc
                    With .item(i).CodeModule
                        p = .CountOfLines '总行数
                        k = k + p
                        For n = p To 1 Step -1
                            strx = .Lines(n, 1)
                            strx1 = Trim(strx)
                            If strx1 = vbNullString Then
                                bl = bl + 1 '计算空行
                            Else
                                If InStr(strx1, Chr(39)) > 0 Then j = j + 1 'chr(39)'单引号
                            End If
                        Next
                    End With
                Next
            End With
            If GetClipboard <> ChCode Then  '如果原来的word文件处于关闭的状态,就关闭word,在word中设置open事件
            '------------------------------------------------假设从剪切板获取到"HLAstaticx"值, 那么将会清空剪切板,假设docm文件处于打开的状态,将不会触发open事件
            '-------https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/methods-microsoft-forms
                wd.Close savechanges:=False
                Set objdata = New DataObject
                objdata.SetText "" '---------将剪切板设置为空,防止后续对word的启动造成干扰
                objdata.PutInClipboard
                Set objdata = Nothing
            End If
            Set wd = Nothing
        End If
        With Me
            .Label15 = Format(Now, "yyyy/mm/dd")
            .Label1.Caption = k
            .Label4.Caption = lc
            .Label6.Caption = m + c + 76
            .Label8.Caption = j
            .Label17.Caption = bl
        End With
        .Application.ScreenUpdating = True
    End With
    Statisticsx = 0
    Set obj = Nothing
    Set UF = Nothing
End Sub
