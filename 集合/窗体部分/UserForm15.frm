VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm15 
   Caption         =   "浏览器"
   ClientHeight    =   6495
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   22440
   OleObjectBlob   =   "UserForm15.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arrname() As Variant
Dim arrcode() As Variant
Dim arrgr As Byte
'-------------------------'搜索
Dim urladdress As String '地址栏地址
Dim listviewx As Byte '控制listview的生成
Dim alradd As String, alrurl As String '用于控制添加
Private Const DicUrl As String = "http://www.iciba.com/"
Private Const DbUrl As String = "https://book.douban.com/"
Private Const BdUrl As String = "https://www.baidu.com/"

Private Sub CommandButton130_Click()
    Me.WebBrowser1.Stop
End Sub

Private Sub CommandButton131_Click() '搜索
    Dim strx As String
    strx = Trim(Me.TextBox30.Text)
    If Len(strx) = 0 Then Exit Sub
    Search strx
End Sub

Private Sub CommandButton132_Click() '主页
    Me.WebBrowser1.GoHome
End Sub

Private Sub CommandButton133_Click() '前进
    Me.WebBrowser1.goForward
End Sub

Private Sub CommandButton134_Click() '后退
    Me.WebBrowser1.goBack
End Sub

Private Sub CommandButton135_Click() '搜索选中
    Dim selectobj As Object, rngobj As Object, i As Integer, strx As String
    
    Set selectobj = Me.WebBrowser1.Document.Selection
    If selectobj Is Nothing Then Exit Sub
    Set rngobj = selectobj.createrange
    If rngobj Is Nothing Then Exit Sub
    strx = rngobj.HTMLText
    If InStr(strx, ">") > 0 Then
        strx = Trim(Split(Split(strx, ">")(1), "<")(0)) '获取网页上选中的内容 '需要修正,不同的页面获取到的内容都不一样
    End If
    i = Len(strx)
    If i > 25 Then
        MsgShow "内容长度超出范围", "提示", 1200: Set selectobj = Nothing: Set rngobj = Nothing: Exit Sub
    ElseIf i = 0 Then
        Set selectobj = Nothing
        Set rngobj = Nothing
        Exit Sub
    End If
    Search strx
    Set selectobj = Nothing
    Set rngobj = Nothing
End Sub

Private Sub CommandButton136_Click() '展开
    Dim i As Integer, k As Integer
    With Me
        If listviewx = 0 Then
            With .ListView1
                .ColumnHeaders.Add , , "编码", 75, lvwColumnLeft
                .ColumnHeaders.Add , , "文件名", 185, lvwColumnLeft
                .View = lvwReport                            '以报表的格式显示
                .LabelEdit = lvwManual                       '使内容不可编辑
                .Gridlines = True
            End With
            listviewx = 1
        End If
'        i = Int(ActiveWindow.UsableHeight)
        k = Int(ActiveWindow.UsableWidth)
        If .CommandButton136.Caption = "信息编辑" Then
            .Width = 1134
            .StartUpPosition = 0 '调整窗体的位置
            .Left = k - 1127
            .CommandButton137.Enabled = True
            .CommandButton136.Caption = "关闭编辑"
        Else
            .Width = 858
            .StartUpPosition = 0
            .Left = 138 '(1134-858)/2
            .CommandButton137.Enabled = False
            .CommandButton136.Caption = "信息编辑"
        End If
    End With
End Sub

Private Sub CommandButton137_Click() '匹配
    Dim strx As String, i As Byte, k As Byte, m As Byte
    
    With Me
        strx = .WebBrowser1.LocationURL
        k = InStr(strx, "search.douban.com/book/") '分别匹配豆瓣搜索页面的值和书籍页面的值
        i = InStr(strx, "book.douban.com/subject")
        If i > 0 Or k > 0 Then
            strx = .Caption
            If i > 0 Then
                strx = Trim(Left$(strx, m - 4)) '(豆瓣)"
            Else
                If InStr(strx, "-") > 0 Then
                    If k > 0 Then strx = Trim(Split(strx, "-")(0)) '书名-读书-豆瓣搜索
                End If
            End If
            FiltterFile strx
        End If
    End With
End Sub

Private Function FiltterFile(ByVal strText As String) '从目录中筛选
    Dim i As Integer, strx1 As String * 1, p As Byte, c As Byte, j As Byte, m As Byte, k As Integer, strxtemp As String
    Dim arr() As String, strLen As Byte
    
    If arrgr = 0 Then ArrayLoad
    i = arrgr - 5
    If i < 1 Then Exit Function
    strLen = Len(strText)
    If Len(strLen) = 0 Then Exit Function
    ReDim arr(strLen)
    arr = ObtainKeyWord(strText) '分离中英文,数字
    With Me.ListView1.ListItems
        .Clear
        For k = 1 To i
            If InStr(1, arrname(k, 1), strText, vbTextCompare) > 0 Then '先进行整体处理-单个处理-分离后的整体处理
                m = 1
            Else
                strxtemp = Replace(strText, " ", "") '剔除掉空格
                j = Len(strxtemp)
                For p = 1 To j
                    strx1 = Mid$(strxtemp, p, 1)
                    If strx1 Like "[一-龥]" Then '处理单个的中文,中文可以使用二进制比较
                        If InStr(1, arrname(k, 1), strx1, vbBinaryCompare) > 0 Then c = c + 1
                    End If
                Next
                If Int((c / j) * 100) > 33 Then
                    m = 1 '匹配超过33%
                    c = 0
                Else
                    For p = 0 To strLen '分离后的整体处理
                        If Len(arr(p)) > 0 Then
                            If InStr(1, arrname(k, 1), arr(p), vbTextCompare) > 0 Then m = 1
                        End If
                    Next
                End If
            End If
            If m = 1 Then
                m = 0
                With .Add
                    .Text = arrcode(k, 1)
                    .SubItems(1) = arrname(k, 1)
                End With
            End If
        Next
    End With
    Erase arr
End Function

Private Sub TextBox31_Change()
    Dim i As Integer, j As Integer, k As Integer
    Dim strx As String, mi As Byte
    
    strx = Me.TextBox31.Value
    strx = Replace(strx, "/", " ") '替换掉"/"符号
    ArrayLoad
    blow = arrgr - 5
    If blow < 1 Then Exit Sub
    With Me.ListView1.ListItems
        If Len(strx) >= 2 Then
            .Clear
            mi = 0
            For k = 1 To blow
                If InStr(1, arrname(k, 1) & "/" & arrcode(k, 1), strx, vbTextCompare) > 0 Then
                    With .Add
                        .Text = arrcode(k, 1)
                        .SubItems(1) = arrname(k, 1)
                    End With
                    mi = mi + 1
                    If mi > 10 Then Exit For
                Else
                   If mi = 0 Then .Clear
                End If
            Next
        Else
            .Clear
        End If
    End With
End Sub

Private Sub ArrayLoad()
    Dim i As Integer
    If arrgr > 0 Then Exit Sub
    With ThisWorkbook.Sheets("书库")
            i = .[e65536].End(xlUp).Row
            arrgr = i
            If i <= 5 Then Exit Sub
            If i = 6 Then
                ReDim arrname(1 To 1, 1 To 1)
                ReDim arrcode(1 To 1, 1 To 1)
                arrname(1, 1) = .Cells(6, 3).Value
                arrcode(1, 1) = .Cells(6, 2).Value
            Else
                arrname = .Range("c6:c" & i).Value
                arrcode = .Range("b6:b" & i).Value
            End If
        End If
    End With
End Sub

Private Sub CommandButton138_Click() '抓取豆瓣信息
    Dim strx As String, i As Integer, k As Integer, arr() As String
    Dim strx1 As String, strx2 As String, strx3 As String, strx4 As String, strx5 As String, yesno As Variant
    Dim strx6 As String, strx7 As String, xi As Variant
    
    With Me
        strx = .WebBrowser1.LocationURL
        If alrurl = strx Then Exit Sub '已经添加过
        If InStr(strx, "book.douban.com/subject") = 0 Then Exit Sub
        With .ListView1
            i = .ListItems.Count
            If i = 0 Then Exit Sub
            For k = 1 To i
                If .ListItems(k).Selected = True Then
                    strx1 = .ListItems(k).Text
                    If alradd = strx1 Then
                        yesno = MsgBox("内容已存在,是否继续添加", vbYesNo, "Warning") '已经添加内容,但不一定相同
                        If yesno = vbNo Then Exit Sub
                    End If
                    SearchFile strx1
                    If Rng Is Nothing Then MsgShow "文件已经丢失", "Warning", 1200: Set Rng = Nothing: Exit Sub
                    If Me.WebBrowser1.Busy = True Then Me.WebBrowser1.Stop       '停止页面加载
                    With Me.WebBrowser1.Document
                    '-----直接读取innerhtml的内容会出现没有部分的内容没有的情况,只能读取到网页的部分信息和在浏览器中直接查看源码看到的信息不一样
                        strx2 = .getElementById("interest_sectl").InnerHtml '获得评分部分的源码 '注意某些html的特性,IE不一定支持
                        strx3 = .getElementById("info").InnerHtml '作者/国籍,
                        strx4 = .getElementById("mainpic").InnerHtml '书名+封面
                    End With
    
                    ReDim arr(1 To 5)
                    arr = DoubanTreat(strx2, strx3, strx4)
                    With Rng '将内容写入表格
                        .Offset(0, 23) = arr(3) '名称
                        .Offset(0, 24) = arr(1) '评分
                        .Offset(0, 25) = strx   '链接
                        .Offset(0, 14) = arr(2) '作者
                        strx5 = arr(4) '封面名称'CheckRname
                        xi = Split(strx5, "/")
                        strx5 = xi(UBound(xi)) '文件名
                        strx5 = Right$(strx5, Len(strx5) - InStrRev(strx5, ".") + 1)
                        strx5 = strx1 & strx5
                        strx6 = ThisWorkbook.Path & "\" & "bookcover"
                        If fso.folderexists(strx6) = False Then fso.CreateFolder strx6
                        strx5 = strx6 & "\" & strx5 '封面的存储路径
                        If fso.fileexists(strx5) = True Then fso.DeleteFile strx5, True '如果存在同名文件就删掉
                        strx7 = LCase(Right$(arr(4), 3))
                        If strx7 = "jpg" Or strx7 = "png" Then '判断链接的内容是否满足要求
                            If DownloadFilex(arr(4), strx5) = True Then
                                .Offset(0, 34) = arr(4) '封面链接(用于封面丢失,重新下载)
                                .Offset(0, 36) = strx5 '封面路径
                            Else
                                strx5 = ""
                            End If
                        End If
                        If Len(arr(5)) > 0 Then .Offset(0, 37) = arr(5) '作者国籍
                    End With
                    
                    If UF3Show > 0 Then
                        With UserForm3 '将内容写入窗体
                            If Len(.Label29.Caption) > 0 And .Label55.Visible = False Then
                                If strx1 = .Label29.Caption Then
                                    .Label106.Caption = strx1
                                    .TextBox3.Text = arr(3)
                                    .TextBox4.Text = arr(2)
                                    .Label69.Caption = arr(1)
                                End If
                                If Len(strx5) > 0 Then BookCoverShow strx5 '封面下载成功(显示封面)
                            End If
                        End With
                    End If
                    alradd = strx1
                    alrurl = strx
                    Exit For
                End If
            Next
        End With
    End With
    Set Rng = Nothing
    Erase arr
    MsgShow "操作完成", "Tips", 1200
End Sub

Private Sub CommandButton139_Click() '豆瓣
    Webbrowserx DbUrl
End Sub

Private Sub CommandButton140_Click() '金山词霸
    Webbrowserx DicUrl
End Sub

Private Sub CommandButton141_Click() '链接二维码
    Dim strx As String
    strx = Me.WebBrowser1.LocationURL
    If Len(strx) = 0 Then Exit Sub
    If strx Like "*[一-龥]*" Then
        QRtextCN = strx
        UserForm1.Show
    Else
        QRtextEN = strx
        UserForm18.Show
    End If
End Sub

Private Sub CommandButton142_Click() '复制链接地址
    Dim textb As Object, strx As String
    
    With Me
        strx = .WebBrowser1.LocationURL
        If Len(strx) = 0 Then Exit Sub
        Set textb = .Controls.Add("Forms.TextBox.1", "Text1", False) '以创建临时textbox的方式实现复制内容;
        With textb
            .Text = strx
            .SelStart = 0
            .SelLength = Len(.Text)
            .Copy
        End With
    End With
    Set textb = Nothing
End Sub

Private Sub CommandButton143_Click() '外部链接打开
    Webbrowser Me.WebBrowser1.LocationURL
End Sub

Private Sub ListView1_DblClick() '双击打开文件夹
    Dim strx1 As String, strx2 As String, i As Integer
    
    With Me.ListView1
        If .ListItems.Count = 0 Then Exit Sub
        strx1 = .SelectedItem.Text
        strx2 = LCase(Right(strx1, Len(strx1) - InStrRev(strx1, "."))) '文件扩展名
        If strx2 Like "xl*" Then MsgShow "不允许打开Excel关联文件", "Warning", 1500: Exit Sub
        SearchFile strx1
        i = .SelectedItem.Index
        If Rng Is Nothing Then
            .ListItems.Remove (i)
            MsgShow "文件不存在", "Warning", 1800
            Set Rng = Nothing
            Exit Sub
        End If
        If fso.fileexists(Rng.Offset(0, 3)) = False Then .ListItems.Remove (i)
    End With
    With Rng
        OpenFile .Offset(0, 0), .Offset(0, 1), .Offset(0, 2), .Offset(0, 3), 1, .Offset(0, 26), 1
    End With
    Set Rng = Nothing
End Sub

Private Sub TextBox30_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '双击复制信息
    With Me.TextBox30 '可以利用textbox的这个属性作为实现复制信息到剪切板的间接途径
        If Len(.Text) And Len(urladdress) = 0 Then Exit Sub
        .Text = urladdress
        .SelStart = 0
        .SelLength = Len(.Text)
        .Copy
    End With
End Sub

Private Sub TextBox30_Enter() '当光标进入地址栏
    urladdress = Me.TextBox30.Text
    Me.TextBox30.Text = ""
End Sub

Private Sub TextBox30_Exit(ByVal Cancel As MSForms.ReturnBoolean)   '当光标离开地址栏
    If Len(Trim(Me.TextBox30.Text)) = 0 Then Me.TextBox30.Text = urladdress
End Sub

Private Sub TextBox30_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim strx As String
    If KeyCode = 13 Then
        strx = Trim(Me.TextBox30.Text)
        If Len(strx) = 0 Then Exit Sub
        If CheckUrl(strx) = False Then Search strx: Exit Sub
        Webbrowserx strx
    End If
End Sub

Private Sub UserForm_Initialize()
    If Statisticsx = 1 Then Exit Sub
    If Len(Turlx) = 0 Then Turlx = BdUrl
    With Me
        .Width = 858
    End With
    Webbrowserx Turlx
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Turlx = "" '外部调用重置
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, url As Variant) '地址栏显示
    With Me
        .TextBox30.Text = .WebBrowser1.LocationURL
    End With
End Sub

Function Search(ByVal searchkey As String) '搜索
    Dim sengine As String, Urlx As String
    
    With Me
        If .OptionButton13.Value = True Then
            sengine = "https://www.baidu.com/s?wd="
            Urlx = sengine & Replace(searchkey, " ", "%20") 'baidu
        
        ElseIf .OptionButton14.Value = True Then
            sengine = "https://book.douban.com/subject_search?cat=1003&search_text="
            Urlx = sengine & Replace(searchkey, " ", "+")  'douban
        
        ElseIf .OptionButton15.Value = True Then
            sengine = "https://www.bing.com/search?q="  'bing
            Urlx = sengine & Replace(searchkey, " ", "%20")
        
        ElseIf .OptionButton13.Value = False And .OptionButton14.Value = False And .OptionButton15.Value = False Then
            sengine = "https://www.baidu.com/s?wd="
            Urlx = sengine & Replace(searchkey, " ", "%20") '默认百度
        End If
        Webbrowserx (Urlx)
    End With
End Function

Function CheckUrl(ByVal strText As String) As Boolean '判断地址栏上的信息是不是url
    CheckUrl = True
    If Len(strText) < 3 Then CheckUrl = False: Exit Function
    strText = LCase(strText)
    If InStr(strText, "http") = 0 And InStr(strText, ".com") = 0 And InStr(strText, ".org") = 0 And InStr(strText, ".cn") = 0 And InStr(strText, ".net") = 0 And _
        InStr(strText, ".edu") = 0 And InStr(strText, ".html") = 0 And InStr(strText, ".cc") = 0 Then CheckUrl = False
End Function

Private Function Webbrowserx(ByVal url As String) '执行浏览器动作 'webbrowser - https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa752043%28v%3dvs.85%29
    With Me.WebBrowser1
        .MenuBar = False
        .Silent = True
        .Navigate (url)
    End With
End Function

Private Sub WebBrowser1_TitleChange(ByVal Text As String) '窗体显示网站title
    Me.Caption = Text
End Sub
