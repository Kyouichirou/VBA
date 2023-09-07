VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm25 
   Caption         =   "文本转语音"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17880
   OleObjectBlob   =   "UserForm25.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------https://wutils.com/com-dll/constants/constants-SpeechLib.htm
Private Const SAFT48kHz16BitStereo = 39
Private Const SSFMCreateForWrite As Byte = 3 ' Creates file even if file exists and so destroys or overwrites the existing file
Private Const SSFMOpenForRead As Byte = 0
Private Const SSFMOpenReadWrite As Byte = 1
Private Const SSFMCreate As Byte = 2
Private Const FlagsAsync As Byte = 1
Dim oVoice As Object, oFileOpen As Object
Dim Flapx As Boolean
Dim voi As Byte, Vox As Byte, vor As Integer
Dim FilePath As String
Private Declare Function timeGetTime Lib "winmm.dll" () As Long '时间api可以精确到毫秒
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim oStream As Object '文件流
Dim IsClose As Boolean
Dim Lbox1 As Integer

Private Function ObtainText(ByVal Urlx As String, ByVal FilePath As String) As Boolean '抓取页面小说内容
    Dim fl As Object, FilePath As String
    Dim bookt As Object
    Dim strHtml As String, textl As Object, HtmF As Object
    Dim XmlH As Object
    Dim rt As Long
    Dim i As Integer, strTemp As String
    Dim k As Integer, j As Integer, p As Integer
    
    ObtainText = True
    Set HtmF = CreateObject("htmlfile")
    Set XmlH = CreateObject("Msxml2.ServerXMLHTTP") 'https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms766431(v=vs.85)
    HtmF.DesignMode = "on" ' 开启编辑模式
    With XmlH
        .Open "GET", Urlx, True 'https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms757849(v=vs.85)
        .Send
        rt = timeGetTime
        Do While .readyState <> 4 And timeGetTime - rt < 6000 '等待信息返回
            Sleep 25
            DoEvents
        Loop
        If .readyState <> 4 Then '没有成功获取到信息
            XmlH.Close
            ObtainText = False
            Set XmlH = Nothing
            Exit Function
        End If
        strHtml = .responseText
        XmlH.Close
    End With
    HtmF.Write strHtml ' 写入数据
    Set bookt = HtmF.getElementById("BookText") '根据id提取数据
    Set fl = fso.OpenTextFile(FilePath, ForWriting, True, TristateUseDefault)
    k = bookt.Children.Length - 1
    j = bookt.ChildNodes.Length - 1 '以这个两个的数据为准,而不是直接使用elementcount
    For Each textl In bookt.Children '每个节点
        strTemp = Trim(textl.innertext)
        If Len(strTemp) > 0 Then
            If p < k Then
                fl.WriteLine strTemp '将数据写入txt
                p = p + 1
            End If
        Else
            i = i + 1 '用于检测判断信息存储的位置
            If i = 3 Then ObtainText = False: Exit For
        End If
    Next
    If i = 3 Then
        strTemp = ""
        i = 0
        p = 0
        ObtainText = True
        For Each textl In bookt.ChildNodes
            If textl.NodeType = 3 Then
                strTemp = Trim(textl.Data)
                If Len(strTemp) > 0 Then
                    If p < j Then   '-------最后的一个节点是无关信息
                        fl.WriteLine strTemp
                        p = p + 1
                    End If
                Else
                    i = i + 1
                    If i = 3 Then ObtainText = False: Exit For
                End If
            End If
        Next
    End If
    fl.Close
    Set fl = Nothing
    Set XmlH = Nothing
    Set HtmF = Nothing
End Function

Private Function TextSpeak(ByVal strText As String) '将文本转换为语音
    If Vox = 0 Then Set oVoice = CreateObject("SAPI.SpVoice"): Vox = Vox + 1
    oVoice.Speak strText, FlagsAsync '异步播放
End Function

Private Sub CommandButton1_Click()
    If IsClose = True Then Exit Sub
    With Me.CommandButton1
        If .Caption = "暂停" Then
            .Caption = "播放"
            oVoice.Pause
        Else
            .Caption = "暂停"
            oVoice.Resume
        End If
    End With
End Sub

Private Sub CommandButton10_Click() '跳转播放位置
If IsClose = True Then Exit Sub

End Sub

Private Sub CommandButton11_Click() '制成音频文件
    Dim strx As String
    Dim FilePath As String
    
    strx = Trim(Me.TextBox2.Text)
    If Len(strx) = 0 Then Exit Sub
    FilePath = ThisWorkbook.Path & "\" & Format(Now, "yyyymmddhhmmss") & ".wav"
    If TextToVoice(FilePath, 0, strText:=strx) = True Then
        strx = "转换成功"
    Else
        strx = "转换失败"
    End If
    MsgShow strx, "Tips", 1200
End Sub

Private Sub CommandButton12_Click() '获取文本
    If IsClose = True Then Exit Sub
    If Not oStream Is Nothing Then Me.TextBox2.Text = GetString(sPosition:=0)
End Sub

Private Sub CommandButton13_Click() '重新写入文本

End Sub

Private Sub CommandButton14_Click() '播放文本


End Sub

Private Sub CommandButton2_Click() '关闭
    If IsClose = True Then Exit Sub
    oVoice.Pause
    Set oVoice = Nothing
    fl.Close
    Set fl = Nothing
    IsClose = True
End Sub
'-----------------------------------------------Values for the Volume property range from 0 to 100
Private Sub CommandButton3_Click() '音量+
    If IsClose = True Then Exit Sub
    If voi = 100 Then
        Me.CommandButton3.Enabled = False
    Else
        voi = voi + 10
        oVoice.Volume = voi
    End If
End Sub

Private Sub CommandButton4_Click() '音量-
    If IsClose = True Then Exit Sub
    If voi = 0 Then
        Me.CommandButton4.Enabled = False
    Else
        voi = voi - 10
        oVoice.Volume = voi
    End If
    Me.CommandButton3.Enabled = True
End Sub

Private Sub ListBox1_Click() '搜索结果点击其他的结果
    If Lbox1 = -1 Then Exit Sub
    With Me
        If .ListBox1.ListCount = 0 Then Exit Sub
        If Lbox1 <> .ListBox1.ListIndex Then
            If .CommandButton5.Caption <> "打开" Then .CommandButton5.Caption = "打开"
        Else
            .CommandButton5.Caption = "前进"
        End If
    End With
End Sub

Private Sub CommandButton5_Click() '打开链接
    Dim strx As String, yesno As Variant
    Dim arr() As String
    Dim i As Integer, j As Byte
    Dim strTemp As String
    
    With Me
        strTemp = .CommandButton5.Caption
        If strTemp = "打开" Then
            If .ListBox1.ListCount = 0 Then Exit Sub
            i = .ListBox1.ListIndex
            If i = -1 Then Exit Sub '-1表示未选中
            Lbox1 = i
            arr = ObtainPage_Info(.ListBox1.List(i, 1), 0)
            If Pagexs > 0 Then
                .ComboBox1.Text = 1
                For j = 1 To Pagexs
                    .ComboBox1.AddItem j
                Next
            End If
            j = UBound(arr)
            With .ListBox2
                For i = 0 To j
                    .AddItem
                    .List(i, 0) = arr(i, 0)
                    .List(i, 1) = arr(i, 1)
                Next
            End With
            .ListBox1.Visible = False
            .ListBox2.Visible = True
            .CommandButton5.Caption = "返回"
            .CommandButton9.Visible = True
        ElseIf strTemp = "返回" Then
            .CommandButton9.Visible = False
            .ListBox1.Visible = True
            .ListBox2.Visible = False
            .CommandButton5.Caption = "前进"
        ElseIf strTemp = "前进" Then
            .ListBox1.Visible = False
            .ListBox2.Visible = True
            .CommandButton9.Visible = True
            .CommandButton5.Caption = "返回"
        End If
    End With
End Sub

Private Sub CommandButton9_Click() '播放
    Dim i As Integer
    Dim FilePath As String
    Dim folderx As Folder, fl As File
    Dim strx As String
    Dim strTemp As String
    '----------先检查本地文件是否存在, 如果存在则优先播放本地文件
    With Me
        If .ListBox2.ListCount = 0 Then Exit Sub
        i = .ListBox2.ListIndex
        If i = -1 Then Exit Sub
        strx = .ListBox2.List(i, 0)
        strTemp = ThisWorkbook.Path & "\" & .ListBox1.List(Lbox1, 0)
        FilePath = strTemp
        If fso.folderexists(FilePath) = True Then '查看是否存在本地文件夹
            Set folderx = fso.GetFolder(FilePath)
            FilePath = ""
            For Each fl In folderx.Files
            If strx = fl.Name Then
                If fl.Size > 0 Then FilePath = fl.Path: Exit For
            End If
            Next
            Set folderx = Nothing
            If Len(FilePath) = 0 Then
                If IsNetConnectOnline = False Then
                    MsgBox "网络不可用", vbInformation, "Tips": Exit Sub
                Else
                    FilePath = strTemp & "\" & strx & ".txt"
                    ObtainText strx, FilePath
                End If
            Else
            FilePath = strTemp & "\" & strx & ".txt"
            TextSpeak (GetString(FilePath))
            End If
        Else
        fso.CreateFolder FilePath '创建本地文件夹
        End If
    End With
IsClose = False
End Sub

Private Sub CommandButton6_Click() '搜索
    Dim strx As String
    Dim arr() As String
    Dim i As Integer
    Dim k As Integer
    
    strx = Trim(Me.TextBox1.Text)
    If Len(strx) < 2 Then Exit Sub
    If IsNetConnectOnline = False Then MsgShow "网络未连接", "Tips", 1200: Exit Sub
    arr = ObtainPage_Info(strx, 1)
    i = UBound(arr)
    With Me.ListBox1 '增加搜索不到结果
        For k = 0 To i
            .AddItem
            .List(k, 0) = arr(k, 0)
            .List(k, 1) = arr(k, 1)
        Next
    End With
    Me.CommandButton5.Caption = "打开"
End Sub
'-----------------------------------------------Values for the Rate property range from -10 to 10
Private Sub CommandButton7_Click() '速度+
    If IsClose = True Then Exit Sub
    If vor = 10 Then
        Me.CommandButton7.Enabled = False
    Else
        vor = vor + 1
        oVoice.Volume = vor
    End If
    Me.CommandButton8.Enabled = True
End Sub

Private Sub CommandButton8_Click() '速度-
    If IsClose = True Then Exit Sub
    If vor = -10 Then
        Me.CommandButton8.Enabled = False
    Else
        vor = vor - 1
        oVoice.Volume = vor
    End If
    Me.CommandButton7.Enabled = True
End Sub

Private Sub UserForm_Initialize()
    If Statisticsx = 1 Then Exit Sub
    voi = 50
    vor = 0
    Flapx = False
    IsClose = True
    Lbox1 = -1
    With Me
        .ComboBox2.List = Array("10%", "20%", "30%", "40%", "50%", "60%", "70%", "80%", "90%")
    End With
End Sub

Private Function GetString(Optional ByVal FilePath As String, Optional ByVal sCharset As String = "gb2312", Optional ByVal sPosition As Long = 0) As String
    Const adTypeBinary As Byte = 1
    Const adTypeText As Byte = 2
    Const adModeRead As Byte = 1
    Const adModeWrite As Byte = 2
    Const adModeReadWrite As Byte = 3
    
    '----------------------从指定的位置读取信息
    If oStream Is Nothing Then
        If fso.fileexists(FilePath) = False Then Exit Function
        Set oStream = CreateObject("ADODB.Stream")
        With oStream
            .Mode = 3
            .type = adTypeText
            .CharSet = sCharset
            .Open
            .LoadFromFile (FilePath)
        End With
    Else
        With oStream
            .Position = sPosition
            GetString = .ReadText()
        End With
    End If
End Function

Private Sub UserForm_Terminate()
    If Statisticsx = 1 Then Exit Sub
    If IsClose = True Then Exit Sub
    If Not oStream Is Nothing Then
        oStream.Close
        Set oStream = Nothing
    End If
End Sub
