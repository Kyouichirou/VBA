VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm21 
   Caption         =   "FTP管理"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3975
   OleObjectBlob   =   "UserForm21.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents cf As cFTP    '需要响应事件
Attribute cf.VB_VarHelpID = -1
Dim strHost As String           '主机名称，IP地址形式
Dim strUser As String           '用户名
Dim strPassword As String       '用户密码
Dim strFlag As String           '设置状态栏的内容
Dim FilePath As String          '将userform3的文件路径提取出来

'cFTP类模块下载或上传文件的事件，可以使用这个事件做成进程状态栏
Private Sub cf_FileTransferProgress(lCurrentBytes As Long, lTotalBytes As Long)
    If strFlag = "Downloading" Then     '如果设置成Downloading，表示下载单个文件
        txtStatus.Text = Format(lCurrentBytes / lTotalBytes * 100, "0.00") & "% Done..."
    Else
        '多个文件上传，在进程比例之前添加更多信息
        txtStatus.Text = strFlag & " | " & Format(lCurrentBytes / lTotalBytes * 100, "0.00") & "% Done..."
    End If
End Sub

Private Sub cmdConnect_Click() '连接
    Dim strContent As String
    Dim lReturn As Long
    Dim lPort As Long
    Dim i As Integer
    Dim arrCont() As String
    
    If cmdConnect.Caption = "连接" Then
        '设置主机名称
        If Trim(txtHost.Text) = "" Then
            MsgBox "必须设置主机名称！", vbInformation, "提示"
            Exit Sub
        Else
            strHost = Trim(txtHost.Text)
        End If
        constatus = 0
        If IsPing(strHost) = False Then MsgBox "设备未连接": Exit Sub '检查设备是否处于链接状态 '由于模块的自检ftp是否链接非常慢(需要等待10+s)
        lPort = Trim(txtPort.Text)
        If CheckPort(strHost, lPort) = False Then MsgBox "设备FTP未打开": Exit Sub '检查设备的FTP的打开状态
        '因为类模块中的自带的检查链接状态在设备断开的时会出现非常长时间的卡顿,所以需要额外的功能来实现检查
        '如果用户名为空，则表示匿名
        If Trim(txtUser.Text) = "" And Trim(txtPassword.Text) = "" Then
            strUser = "anonymous"
            strPassword = ""
        Else
            strUser = Trim(txtUser.Text)
            strPassword = Trim(txtPassword.Text)
        End If
        Set cf = New cFTP
        If chkPassive.Value = True Then
            cf.SetModePassive
        Else
            cf.SetModeActive
        End If
        
        '建立连接，返回False表示连接失败
        If lPort = "" Then
            lReturn = cf.OpenConnection(strHost, strUser, strPassword)
        Else
            lReturn = cf.OpenConnection(strHost, strUser, strPassword, lPort)
        End If
        
        If lReturn = False Then
            GoTo ErrHandle
        Else
            '获取当前目录的文件列表，并加入到列表框中，并将当前目录置于文本框中
            '文件夹在尾部添加"/"字符，文件则在尾部添加文件大小
            strContent = cf.GetFTPDirectory
            txtPath.Text = strContent
            strContent = cf.GetFTPDirectoryContent
            '返回内容为“False”时则表示没有文件
            If strContent = "False" Then
                lstFile.Clear
            Else
                arrCont = Split(strContent, vbCrLf)
                lstFile.Clear
                For i = 0 To UBound(arrCont)
                    If Trim(arrCont(i)) <> "" Then
                        lstFile.AddItem arrCont(i)
                    End If
                Next i
            End If
            cmdConnect.Caption = "断开连接"
            txtStatus.Text = "用户" & strUser & "已连接到主机" & strHost
        End If
    Else
        cf.CloseConnection
        txtPath.Text = ""
        lstFile.Clear
        cmdConnect.Caption = "连接"
        txtStatus.Text = "用户退出，未连接"
    End If
    Exit Sub
ErrHandle:
    MsgBox cf.GetLastErrorMessage
End Sub

'在FTP的当前目录下创建新的目录
Private Sub cmdCreate_Click()
    Dim strPath As String
    If cmdConnect.Caption = "连接" Then Exit Sub
    If Right(strPath, 1) = "/" Then strPath = Mid(strPath, 1, Len(strPath) - 1)
    strPath = Application.InputBox("请输入你要创建的文件夹名称：", "创建文件夹", "Default")
    If Trim(strPath) = "" Then Exit Sub
    If cf.CreateFTPDirectory(strPath) = True Then
        lstFile.AddItem strPath & "/"
        txtStatus.Text = "创建目录" & strPath & "成功"
    Else
        MsgBox cf.GetLastErrorMessage
    End If
End Sub
'删除指定文件或文件夹，如果文件夹下存在其它文件，将不能删除
Private Sub cmdDelete_Click()
    Dim strPath As String
    If cmdConnect.Caption = "连接" Then Exit Sub
    If lstFile.Text = "" Then Exit Sub
    strPath = lstFile.Text
    If MsgBox("你确定需要删除文件或目录" & strPath & "吗？", vbYesNo) = vbNo Then Exit Sub
    If Right(strPath, 1) = "/" Then
        If cf.RemoveFTPDirectory(strPath) = True Then
            txtStatus.Text = "已删除目录" & strPath
            lstFile.RemoveItem lstFile.ListIndex
        Else
            MsgBox cf.GetLastErrorMessage
        End If
    Else
        strPath = Mid(strPath, 1, Len(strPath) - 1)
        strPath = Mid(strPath, 1, InStrRev(strPath, "(") - 1)
        If cf.DeleteFTPFile(strPath) = True Then
            txtStatus.Text = "已删除文件" & strPath
            lstFile.RemoveItem lstFile.ListIndex
        Else
            MsgBox cf.GetLastErrorMessage
        End If
    End If
End Sub
'下载文件
Private Sub cmdDownload_Click()
    Dim strFile As String
    Dim strSource As String
    Dim fd As FileDialog, i As Byte
    '使用FileDialog对象获取文件夹名称用来保存下载文件
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    strSource = ""
    With fd
        fd.AllowMultiSelect = False
        fd.Filters.Clear
        If .Show = -1 Then strSource = .SelectedItems(1)
    End With
    Set fd = Nothing
    If strSource = "" Then Exit Sub
    If cmdConnect.Caption = "连接" Then Exit Sub
    strFlag = "Downloading"
    strFile = lstFile.Text
    If Trim(strFile) = "" Or Right(strFile, 1) = "/" Then Exit Sub
    strFile = Mid(strFile, 1, Len(strFile) - 1)
    strFile = Mid(strFile, 1, InStrRev(strFile, "(") - 1)
    '可以触发事件
    i = 2 '已二进制的方式传输文件
    If LCase(Right$(strFile, Len(strFile) - InStrRev(strFile, "."))) Like "txt" Then i = 1 '只有文本型文件例外
    If cf.FTPDownloadFile(strSource & "\" & strFile, strFile, i) = True Then
        txtStatus.Text = "下载文件" & strFile & "成功"
    Else
        MsgBox cf.GetLastErrorMessage
    End If
End Sub
'上传文件，可以多选
Private Sub cmdUpload_Click()
    Dim arrFile As Variant
    Dim i As Integer, k As Byte, m As Byte, j As Byte
    Dim strBaseFile As String
    Dim strError As String
    Dim strLocalFile As String
    
    If cmdConnect.Caption = "连接" Then Exit Sub
    strError = ""
    m = 1
    If Len(FilePath) = 0 Then
        arrFile = Application.GetOpenFilename("所有文件(*.*),*.*", , "选择文件上传", , True)
        If IsArray(arrFile) = False Then Exit Sub
        m = UBound(arrFile)
    End If
    For i = 1 To m
        strLocalFile = arrFile(i)
        If ErrCode(strLocalFile, 1) > 1 Then GoTo 100 '检查文件的路径是否包含非ansi, 此上传文件模块采用open的方式,需要修改为ado.stream
        j = j + 1
        strBaseFile = Mid(strLocalFile, InStrRev(strLocalFile, "\") + 1)
        strFlag = i & "/" & UBound(arrFile) & "上传"
        k = 2 '已二进制的方式传输文件
        If LCase(Right$(strLocalFile, Len(strLocalFile) - InStrRev(strLocalFile, "."))) Like "txt" Then k = 1 '只有文本型文件例外
        If cf.FTPUploadFile(strLocalFile, strBaseFile, k) = True Then
            lstFile.AddItem strBaseFile & "(" & Format(FileLen(strLocalFile) / 1024, "0.00") & "kB)" '修改的时候,注意也需要将filelen这种函数给同时处理掉
        Else
            strError = strError & vbCrLf & cf.GetLastErrorMessage
        End If
100
    Next i
    If strError <> "" Then MsgBox strError: Exit Sub
    If j < m Then
        If j = 0 Then
            MsgBox "上传文件失败", vbInformation, "Warning"
        Else
            MsgBox "选择" & m & "个文件" & ";成功上传" & "j" & "个"
        End If
        Exit Sub
    End If
    txtStatus.Text = "上传文件完成"
End Sub
'将当前目录改成上一级文件夹，同时更新列表框中的内容
Private Sub cmdUpper_Click()
    Dim strTemp As String
    Dim strContent As String
    Dim arrCont() As String
    Dim i As Integer
    
    If cmdConnect.Caption = "连接" Then Exit Sub
    If txtPath.Text = "/" Then Exit Sub
    strTemp = txtPath.Text
    strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
    strTemp = Mid(strTemp, 1, InStrRev(strTemp, "/"))
    txtPath.Text = strTemp
    If cf.SetFTPDirectory(txtPath.Text) = True Then
        strContent = cf.GetFTPDirectoryContent
        If strContent = "False" Then
            lstFile.Clear
        Else
            arrCont = Split(strContent, vbCrLf)
            lstFile.Clear
            For i = 0 To UBound(arrCont)
                lstFile.AddItem arrCont(i)
            Next i
        End If
    End If
End Sub

Private Sub CommandButton1_Click() '添加文件
    With UserForm3
    If .Label74.Caption = "Y" Then MsgBox "不支持此文件的上传": Exit Sub
    FilePath = .Label25.Caption
    End With
End Sub

'双击列表框内容可以打开该文件夹，并更新列表框内容
Private Sub lstFile_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim strContent As String
    Dim arrCont() As String
    Dim i As Integer
    Cancel = True
    If cmdConnect.Caption = "连接" Then Exit Sub
    If Right(lstFile.Text, 1) = "/" Then
        strTemp = txtPath.Text
        txtPath.Text = strTemp & lstFile.Text
        If cf.SetFTPDirectory(txtPath.Text) = True Then
            strContent = cf.GetFTPDirectoryContent
            If strContent = "False" Then
                lstFile.Clear
            Else
                arrCont = Split(strContent, vbCrLf)
                lstFile.Clear
                For i = 0 To UBound(arrCont)
                    lstFile.AddItem arrCont(i)
                Next i
            End If
        Else
            txtPath.Text = strTemp
            GoTo errhandle1
        End If
    End If
    Exit Sub
errhandle1:
    MsgBox cf.GetLastErrorMessage
End Sub

Private Sub UserForm_Initialize()
    Dim strx As String, strx1 As String, strx2 As String, strx3 As String

    If Statisticsx = 1 Then Exit Sub
    txtStatus.Text = "未连接"
    With UserForm20
        strx = .Label5.Caption 'ip
        strx1 = .TextBox6.Text '用户名
        strx2 = .TextBox7.Text '密码
        strx3 = .TextBox5.Text '端口
    End With
    With Me
        .txtHost.Text = strx
        .txtPassword.Text = strx2
        .txtPort.Text = strx3
        .txtUser.Text = strx1
    End With
End Sub

Private Sub UserForm_Terminate()
    If Statisticsx = 1 Then Exit Sub
    If Not cf Is Nothing Then Set cf = Nothing
End Sub
