Attribute VB_Name = "表窗"
Option Explicit
Public Text3Ch As String, Text4Ch As String '检测主文件/作者是否经过修改
Dim imgx As Byte
Public Fileptc As Byte '文件保护状态
Public Imgurl As String '封面图片的路径
Public Filenamei As String, Filepathi As String, Folderpathi As String '用于临时存储包含非ansi字符的字符串

Sub DisEvents() '禁用干扰项
    '------------------------ http://www.360doc.com/content/15/0401/06/7835172_459703611.shtml
    '------------------------ https://docs.microsoft.com/zh-cn/office/vba/api/excel.application.interactive
    With ThisWorkbook.Application
        .ScreenUpdating = False '禁止屏幕刷新
        .EnableEvents = False '禁用事件
        .Calculation = xlCalculationManual '禁用自动计算
        .Interactive = False '禁止交互(在执行宏时,如果在表格输入内容会造成宏终止)
    End With
End Sub

Sub EnEvents() '启用
    With ThisWorkbook
        With .Application
            .ScreenUpdating = True
            .EnableEvents = True
            .Calculation = xlCalculationAutomatic
            .Interactive = True
        End With
'        .Save
    End With
End Sub

Function ClearAll(Optional ByVal cmCode As Byte) '清除所有的内容
    With ThisWorkbook
        If cmCode = 0 Then
            .Sheets("书库").Range("b6:ah10000").ClearContents
            .Sheets("目录").Range("a4:q1000").ClearContents
            With .Sheets("主界面")
                .Range("e37:j100").ClearContents
                .Range("d27:l33").ClearContents
                .Range("p23:x33").ClearContents
            End With
        Else
            .Sheets("目录").Range("a4:q1000").ClearContents
        End If
    End With
End Function

'----------------------------------------------------------显示搜索结果
Sub ShowDetail(ByVal filecodex As String) '显示搜索结果
    Dim strx As String
    With UserForm3
        If Len(.Label56.Caption) > 0 Then
            If filecodex <> .Label56.Caption Then .Label55.Visible = False '删除的文件提示
        End If
        .Label29.Caption = Rng.Offset(0, 0) '统一编码
        Filenamei = Rng.Offset(0, 1) '文件名
        .Label23.Caption = Filenamei
        .Label24.Caption = Rng.Offset(0, 2) '文件类型
        Filepathi = Rng.Offset(0, 3) '文件路径
        .Label25.Caption = Filepathi
        Folderpathi = Rng.Offset(0, 4) '文件位置
        .Label26.Caption = Folderpathi
        Call FileChange
        .Label28.Caption = Rng.Offset(0, 8) '文件创建时间
        .Label30.Caption = Rng.Offset(0, 10) '文件类别
        .Label33.Caption = Rng.Offset(0, 21) '标识编码
        .TextBox5.Text = Rng.Offset(0, 19) '标签1
        .TextBox6.Text = Rng.Offset(0, 20) '标签2
        .TextBox3.Text = Rng.Offset(0, 12) '主文件名
        .TextBox4.Text = Rng.Offset(0, 37) & Rng.Offset(0, 14) '国籍 & 作者
        Text4Ch = .TextBox4.Text
        .ComboBox3.Text = Rng.Offset(0, 15) 'pdf清晰度
        .ComboBox4.Text = Rng.Offset(0, 16) '文本质量
        .ComboBox5.Text = Rng.Offset(0, 17) '内容评分
        .ComboBox2.Text = Rng.Offset(0, 18) '推荐评分
        .ComboBox12.Text = Rng.Offset(0, 31) '文字类型
        .Label69.Caption = Rng.Offset(0, 24) '豆瓣评分
        .ComboBox14.Text = Rng.Offset(0, 42) '阅读状态
        With ThisWorkbook.Sheets("temp") '是否启用编辑
            If Len(.Cells(31, "ab").Value) = 0 Then
                Call DisablEdit      '禁止编辑
            Else                    '如果启用编辑,那么就需要进一步检查文件的状态
                Call EnablEdit
            End If
        End With
        
        If Rng.Offset(0, 25) = "ERC" Then
            .Label74.Caption = "Y"
        Else
            .Label74.Caption = "N"
        End If
        
        If Len(Rng.Offset(0, 13).Value) = 0 Then '获取部分文件名的信息
            Call AtoName
        Else
            .TextBox3.Text = Rng.Offset(0, 13) '主文件名
        End If
        
        .CheckBox13.Value = False
        If Len(Rng.Offset(0, 32)) > 0 Then Fileptc = 1: .CheckBox13.Value = True '文件保护状态 'checkbox赋值将会触发click事件
        
        .Label236.Caption = ""
        If Len(Rng.Offset(0, 35)) > 0 Then
            .Label236.Caption = "Y"  '文件是否存在密码
        Else
            .Label236.Caption = "N"
        End If
        
        Call Text2a '获取文件的摘要信息
        .Label106.Caption = Rng.Offset(0, 25) '豆瓣链接 '如果豆瓣信息存在就显示编辑豆瓣信息
        If Len(Rng.Offset(0, 25).Value) > 0 Then
            .CommandButton53.Enabled = True
            .CommandButton53.Caption = "编辑豆瓣信息"
        End If
        strx = Rng.Offset(0, 36)
        BookCoverShow strx
        If .MultiPage1.Value <> 1 Then .MultiPage1.Value = 1 '转到页面
    End With
    Set Rng = Nothing
End Sub

Sub BookCoverShow(ByVal FilePath As String)
    If Len(FilePath) > 0 Then
        If fso.fileexists(FilePath) = True Then '封面
            Imgurl = strx
            .Frame2.Width = 158
            .TextBox2.Width = 143
            .CommandButton134.Left = 610
            .CommandButton125.Left = 654
            With .Label239
                .Visible = True
                .Left = 728
                .Top = 94
                .Caption = "豆瓣封面"
            End With
            With .Image1
                .Left = 708
                .Top = 108
                .Width = 84
                .Height = 122
                .Visible = True
                .Picture = LoadPicture(FilePath)
                .PictureSizeMode = fmPictureSizeModeStretch '调整图片
            End With
            imgx = 1
        End If
    Else
        If imgx = 1 Then
            .Image1.Visible = False
            .Label239.Visible = False
            .Frame2.Width = 246
            .TextBox2.Width = 231
            .CommandButton134.Left = 698
            .CommandButton125.Left = 742
            imgx = 0
            Imgurl = ""
        End If
    End If
End Sub

Sub FileChange() '容易发生变化的信息
    With UserForm3
        .Label76.Caption = Rng.Offset(0, 5) '文件初始大小
        .Label112.Caption = Rng.Offset(0, 6) '文件修改时间
        .Label27.Caption = Rng.Offset(0, 7) '文件大小
        .Label32.Caption = Rng.Offset(0, 12) '打开次数
        .Label31.Caption = Rng.Offset(0, 11) '最近打开时间
        .Label71.Caption = Rng.Offset(0, 9) 'MD5
    End With
End Sub

Function Text2a() '获取文件的摘要信息
    Dim TableName As String, str As String
    
    TableName = "摘要记录"
    With UserForm3
        str = .Label29.Caption
        .TextBox2.Text = ""
        If Len(str) = 0 Then GoTo 100
        If RecData = True Then
            SQL = "select * from [" & TableName & "$] where 统一编码='" & str & "'" '查询数据,如果无,则可以写入输入,如果有就可以显示数据
            Set rs = New ADODB.Recordset
            rs.Open SQL, Conn, adOpenKeyset, adLockOptimistic
            If rs.BOF And rs.EOF Then             '用于判断有无找到数据
                rs.Close
                Set rs = Nothing
            Else
                If IsNull(rs(5)) = False Then .TextBox2.Text = rs(5) '注意这里的数据为空的问题(null), 注意与其他表示空的区别 如empty的等
                rs.Close
                Set rs = Nothing
            End If
        End If
    End With
100
End Function

Function AtoName() '辅助创建主文件名
    Dim i As Byte
    Dim strx As String, strx1 As String
    Dim strLen As Byte
    
    With UserForm3
        strx = .Label23.Caption
        strLen = Len(strx)
        If strLen = 0 Or strLen < 5 Then Exit Function
        If strLen > 16 Then
            i = 10
        ElseIf strLen > 12 And strLen < 17 Then
            i = 7
        ElseIf strLen > 8 And strLen < 13 Then
            i = 5
        Else
            i = 3
        End If
        strx1 = Left$(strx, i)
        .TextBox3.Text = strx1
        Text3Ch = .TextBox3.Text '注意这里不要直接使用strx1的值,防止非ansi编码字符带来的干扰
    End With
End Function

Sub EnablEdit() '启用编辑
  Dim strx As String
  With UserForm3
        If Len(.Label29.Caption) = 0 Then Exit Sub '当页面有信息的时候才能编辑,删除文件不可操作
        .TextBox2.Enabled = True
        .TextBox3.Enabled = True
        .TextBox3.SetFocus
        .TextBox4.Enabled = True
        .TextBox5.Enabled = True
        .TextBox6.Enabled = True
        .ComboBox2.Enabled = True
        .ComboBox5.Enabled = True
        .ComboBox12.Enabled = True
        .CommandButton8.Enabled = True
        .ComboBox14.Enabled = True
        .ComboBox4.Enabled = True '文本质量
        strx = UCase(.Label24.Caption) '扩展名
        If strx = "EPUB" Or strx = "MOBI" Or strx = "TXT" Then .CommandButton53.Enabled = True '豆瓣
        If strx = "PDF" Then
            .ComboBox3.Enabled = True
            .CommandButton53.Enabled = True
        End If
    End With
End Sub

Sub DisablEdit() '窗体-搜索结果禁止编辑
    With UserForm3
        .TextBox2.Enabled = False '摘要信息
        .TextBox3.Enabled = False '主文件名
        .TextBox4.Enabled = False '作者
        .TextBox5.Enabled = False '标签
        .TextBox6.Enabled = False '标签
        .ComboBox2.Enabled = False '推荐指数
        .ComboBox3.Enabled = False 'PDF清晰度
        .ComboBox4.Enabled = False '文本质量
        .ComboBox5.Enabled = False '内容评分
        .ComboBox12.Enabled = False '文字类型
        .ComboBox14.Enabled = False '阅读状态
        .CommandButton53.Enabled = False '豆瓣评分
        .CommandButton54.Visible = False '添加豆瓣信息
        .TextBox16.Visible = False '豆瓣信息编辑
        .TextBox17.Visible = False '豆瓣信息编辑
        .TextBox15.Visible = False '豆瓣信息编辑
        .CommandButton8.Enabled = False '添加信息
        .CommandButton56.Enabled = False 'md5计算
    End With
End Sub
'----------------------------------------------------------显示搜索结果
Sub Rewds() '重置窗口
    If ThisWorkbook.Application.Visible = False Then
        ThisWorkbook.Application.Visible = True
        UserForm4.Hide
        UserForm4.Show
        UserForm4.Caption = "Mini"
    End If
    ThisWorkbook.Windows(1).WindowState = xlMinimized
End Sub

Sub HideOption() '美化首页的界面
    With ThisWorkbook.Windows(1)
        .DisplayFormulas = False
        .DisplayHeadings = False
        .DisplayHorizontalScrollBar = False
        .DisplayVerticalScrollBar = False
        .DisplayWorkbookTabs = False
    End With
    ThisWorkbook.Application.DisplayFormulaBar = False
End Sub

Sub Showoption() '恢复原来的界面
ThisWorkbook.Application.DisplayFormulaBar = True
    With ThisWorkbook.Windows(1)
        .DisplayFormulas = True
        .DisplayHeadings = True
        .DisplayHorizontalScrollBar = True
        .DisplayVerticalScrollBar = True
        .DisplayWorkbookTabs = True
    End With
End Sub

Function DataSwitch()      '单元格转为特定的时间格式
    With ThisWorkbook.Sheets("主界面")
        .Range("w27:w33").NumberFormatLocal = "yyyy/m/d h:mm;@"
    End With
End Function

Function TextSwitch() '文本格式转换
    ThisWorkbook.Sheets("书库").Columns("c:c").NumberFormatLocal = "@"  '强制转换为文本, 对于标题为长数字的文件,如果以数字显示可能出现格式异常
End Function
