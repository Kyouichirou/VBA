VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm13 
   Caption         =   "备份文件"
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8520
   OleObjectBlob   =   "UserForm13.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FilePath As String
Dim filecode As String
Dim filez As Long
Dim Folderpath As String
Dim filex As String
Dim filen As String
Dim arrfilepath(1 To 30) As String
'Dim clickindex As Byte
Dim fc As Byte

Private Sub CommandButton1_Click() '备份文件
    Dim strfolder As String, strx As String, tagx As String, strx1 As String, strx2 As String, i As Byte, dics As String
    Dim strx3 As String
    
    If LenB(filecode) = 0 Then Exit Sub
    If LenB(Folderpath) > 0 Then
        strx3 = Folderpath & "\" & filecode '子文件夹
        If fso.folderexists(Folderpath) = True Then
            If fso.folderexists(strx3) = False Then '创建子文件夹
                fso.CreateFolder strx3
            Else
                If fso.GetFolder(strx3).Files.Count >= 30 Then
                    MsgBox "文件数量超过30", vbOKOnly, "Warning"
                    OpenFileLocation (strx3)
                    If filex Like "xl*" Then Unload Me
                    Exit Sub
                End If
            End If
        Else
            If InStr(Folderpath, "\") > 0 Then
                fso.CreateFolder Folderpath
                fso.CreateFolder strx3
            Else
                GoTo 100
            End If
        End If
    Else
100
        With Application.FileDialog(msoFileDialogFolderPicker) '文件夹选择窗口(文件夹以此只能选一个)
            .Show
            If .SelectedItems.Count = 0 Then Exit Sub '未选择文件夹则退出sub
            strfolder = .SelectedItems(1)
        End With
        If CheckFileFrom(strfolder, 2) = True Then MsgBox "文件夹位置受限", vbCritical, "Warning"
        ThisWorkbook.Sheets("temp").Cells(40, "ab") = strfolder
        Folderpath = strfolder
        strx3 = Folderpath & "\" & filecode '子文件夹
    End If
    '----------------------------------创建对应的备份文件夹和对应的子文件夹
    dics = Left$(Folderpath, 1)
    If filez > fso.GetDrive(dics).AvailableSpace Then MsgBox "磁盘空间不足!", vbCritical, "Warning": Exit Sub '判断磁盘是否有足够的空间
    
    With Me
        tagx = Trim(.TextBox1.Text) '自定义-限制长度
        i = Len(tagx)
        If i > 20 Then
            .Label2.Caption = "自定义长度超出范围"
            .TextBox1.SetFocus
            Exit Sub
        End If
        .Label2.Caption = "备份中...请勿关闭窗口"
        strx3 = strx3 & "\" '具体的文件夹路径
        DoEvents
        fso.CopyFile (FilePath), strx3
        FilePath = strx3 & filen '复制后的文件路径
    
        If i = 0 Then
            If .ComboBox1.Text = "标记文件" Then
                tagx = "NA"
            Else
                strx2 = Trim(.ComboBox1.Text)
                If Len(strx2) > 0 Then
                    tagx = strx2
                Else
                    tagx = "NA"
                End If
            End If
        End If
        strx1 = CStr(Format(Now, "yyyymmddhhmmss"))
        strx = strx1 & "-" & tagx & "." & filex '新的文件名-编号
        fso.GetFile(FilePath).Name = strx          '注意这里不能使用cmd等外部命令复制文件,因为需要等待文件复制完成
        FilePath = strx3 & strx           '新的文件路径
        .CommandButton1.Enabled = False
        .CommandButton2.Enabled = True
        If fc > 0 Then
            With .ListView1.ListItems.Add
                .Text = strx1
                .SubItems(1) = tagx
                .SubItems(2) = Now
                .SubItems(3) = fso.GetFile(FilePath).DateLastModified
            End With
        End If
        fc = fc + 1
        arrfilepath(fc) = FilePath
        .Label2.Caption = "备份成功"
    End With
End Sub

Private Sub CommandButton2_Click() '撤销备份
Dim i As Byte
    On Error GoTo 100
    With Me
        If fso.fileexists(FilePath) = True Then fso.DeleteFile (FilePath)
        .CommandButton2.Enabled = False
        .CommandButton1.Enabled = True
        fc = fc - 1
        If fc > 0 Then
            With .ListView1
                .ListItems.Remove (.ListItems.Count)
            End With
            If fc = 0 Then Exit Sub
            For i = 1 To fc
                arrfilepath(i) = arrfilepath(i)
            Next
        Else
            arrfilepath(1) = ""
        End If
    End With
    Exit Sub
100
    If Err.Number = 70 Then
        Me.Label2.Caption = "文件处于打开的状态"
    Else
        Me.Label2.Caption = "异常"
    End If
    Err.Clear
End Sub

Private Sub CommandButton3_Click() '查看备份
    Dim fd As Folder, i As Byte, litm As Variant, filetemp As Variant, strx As String
    Dim fl As File, strx1 As String, strx2 As String, strx3 As String
    
    With Me.ListView1.ListItems
        If fso.folderexists(Folderpath) = True Then
            strx1 = Folderpath & "\" & filecode '备份文件夹路径
            If fso.folderexists(strx1) = True Then
                Set fd = fso.GetFolder(strx1)
                For Each fl In fd.Files
                    strx = fl.Name
                    If InStr(strx, "-") = 0 Then GoTo 100 '符合格式要求
                    i = i + 1
                    If i = 31 Then GoTo 1000 '限制30
                    arrfilepath(i) = fl.Path
                    filetemp = Split(strx, "-")
                    Set litm = .Add()
                    litm.Text = filetemp(0)
                    strx2 = filetemp(1)
                    strx3 = Left$(strx2, InStrRev(strx2, ".") - 1) '状态
                    litm.SubItems(1) = strx3
                    litm.SubItems(2) = fl.DateCreated
                    litm.SubItems(3) = fl.DateLastModified
100
                Next
            Else
                Me.Label2.Caption = "尚未创建备份"
            End If
        Else
            Me.Label2.Caption = "尚未创建备份"
        End If
    End With
    Me.CommandButton4.Enabled = True
1000
    fc = i
    Set litm = Nothing
    Set fd = Nothing
End Sub

Private Sub CommandButton4_Click() '删除备份
    Dim i As Byte, k As Integer, j As Byte
    
    With Me.ListView1
        i = .ListItems.Count
        If i = 0 Then Exit Sub
        For k = i To 1 Step -1  '在倒着删除的时候,数据类型不能是byte类型(0-255)
        If .ListItems(k).Selected = True Then
            On Error GoTo 100 '如果文件处于打开的状态
            If fso.fileexists(arrfilepath(k)) = True Then fso.DeleteFile (arrfilepath(k))
            .ListItems.Remove (k)
        Else
            j = j + 1
            arrfilepath(j) = arrfilepath(k)
        End If
100
        Next
        fc = j
    End With
End Sub

Private Sub CommandButton5_Click() '清空文件夹
    Dim strx As String
    
    If LenB(Folderpath) = 0 Then Exit Sub
    On Error GoTo 100
    strx = Folderpath & "\" & filecode
    If fso.folderexists(strx) = True Then
        fso.DeleteFolder (strx)
        Me.ListView1.ListItems.Clear
        Me.Label2.Caption = "成功操作"
    End If
    Exit Sub
100
    If Err.Number = 70 Then
        Me.Label2.Caption = "文件夹中的文件处于打开状态"
    Else
        Me.Label2.Caption = "异常"
    End If
End Sub

Private Sub ListView1_DblClick() '双击打开文件
    Dim i As Byte, strx As String
    
    With Me.ListView1
        If .ListItems.Count = 0 Then Exit Sub
        If UCase(filex) Like "XL*" Then Exit Sub
        i = .SelectedItem.Index
        If ErrCode(arrfilepath(i), 1) > 1 Then strx = "ERC"
        Call OpenFile("a", filen, filex, arrfilepath(i), 1, strx, 1)
    End With
End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    With Me
        Select Case KeyAscii
            Case asc("/"), asc("\"), asc(":"), asc("*"), asc("?"), asc("<"), asc(">"), asc("|")
            .Label2.ForeColor = &HFF&
            .Label2.Caption = "请勿输入非法字符：""/\ : * ? <> |"
            .TextBox1.Text = ""
            KeyAscii = 0
        Case Else
            .Label2.Caption = ""
        End Select
    End With
End Sub

Private Sub UserForm_Initialize()
    If Statisticsx = 1 Then Exit Sub
    With Me.ListView1 '窗口初始化
        .ColumnHeaders.Add , , "编号", 94, lvwColumnLeft
        .ColumnHeaders.Add , , "状态", 94, lvwColumnLeft
        .ColumnHeaders.Add , , "创建时间", 94, lvwColumnLeft
        .ColumnHeaders.Add , , "修改时间", 94, lvwColumnLeft
        .View = lvwReport                            '以报表的格式显示
        .LabelEdit = lvwManual                       '使内容不可编辑
        .Gridlines = True
        .MultiSelect = True '支持多选
    End With
    Me.TextBox1.SetFocus
    Me.ComboBox1.List = Array("原件", "初始件", "进展A", "进展B", "进展C", "完成")
    With UserForm3
        FilePath = Filepathi
        If LenB(FilePath) = 0 Then Exit Sub
        filecode = .Label29.Caption
        filez = .Label76.Caption
        filex = .Label24.Caption
        filen = Filenamei
    End With
    Folderpath = ThisWorkbook.Sheets("temp").Cells(40, "ab").Value
End Sub
