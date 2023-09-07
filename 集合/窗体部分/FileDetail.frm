VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FileDetail 
   Caption         =   "文件详情"
   ClientHeight    =   12735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6960
   OleObjectBlob   =   "FileDetail.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "FileDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wd As Object
Dim ArrValue(1 To 34) As String
Dim ArrModule(1 To 34) As String
Dim filex As String
Dim EditSave As Boolean

Private Sub CommandButton1_Click() '修改
    Dim i As Byte
    Dim textbox As Object
    
    For i = 1 To 34
        If i = 1 Or i = 2 Or i = 3 Or i = 4 Or i = 5 Or i = 7 Or i = 18 Or i = 20 Or i = 21 Or i = 32 Then
            Set textbox = Me.Controls("TB" & i)
            textbox.Enabled = True
        End If
    Next
    EditSave = False
    Me.CommandButton3.Enabled = True
    Me.CommandButton2.Enabled = True
    Set textbox = Nothing
    Me.CommandButton1.Enabled = False
End Sub

Private Sub CommandButton2_Click() '保存
    Dim i As Byte
    Dim textbox As Object
    For i = 1 To 34
        If i = 1 Or i = 2 Or i = 3 Or i = 4 Or i = 5 Or i = 7 Or i = 18 Or i = 20 Or i = 21 Or i = 32 Then
            Set textbox = Me.Controls("TB" & i)
            wd.BuiltinDocumentProperties(i) = textbox.Text
            textbox.Enabled = False
        End If
    Next
    EditSave = True
    Me.CommandButton2.Enabled = False
    Me.CommandButton3.Enabled = False
    Set textbox = Nothing
    Me.CommandButton1.Enabled = True
End Sub

Private Sub CommandButton3_Click() '模板
    Dim i As Byte
    Dim textbox As Object
    
    For i = 1 To 34
        If i = 1 Or i = 2 Or i = 3 Or i = 4 Or i = 5 Or i = 7 Or i = 18 Or i = 20 Or i = 21 Or i = 32 Then
            Set textbox = Me.Controls("TB" & i)
            textbox.Text = ArrModule(i)
        End If
    Next
    Set textbox = Nothing
End Sub

Private Sub UserForm_Initialize()
    Dim lb As Object
    Dim i As Byte
    Dim h As Integer
    Dim textbox As Object
    Dim strx As String
    Dim FilePath As String
    Dim p As Variant
    
    If Statisticsx = 1 Then Exit Sub
    On Error Resume Next '部分的属性（value）无法读取
    If UF3Show > 0 Then
        FilePath = Filepathi '路径
        filex = UserForm3.Label24.Caption '扩展名
    End If
    If Len(FilePath) = 0 Then Exit Sub
    If fso.fileexists(FilePath) = False Then MsgBox "文件不存在", vbCritical, "Warning": Exit Sub
    Set wd = CreateObject(FilePath)
    If Err.Number > 0 Then Set wd = Nothing: Exit Sub
    h = 6
    For Each p In wd.BuiltinDocumentProperties
        i = i + 1
        Set lb = Me.Controls.Add("Forms.Label.1", "LB" & i, True) '动态创建标签
        Set textbox = Me.Controls.Add("Forms.TextBox.1", "TB" & i, True) '动态创建文本框
        With lb
            .Caption = p.Name & ":"
            .TextAlign = fmTextAlignRight
            .Height = 14
            .Left = 8
            .Width = 160
            .Top = h
        End With
         strx = p.Value
        With textbox
            .Text = strx
            .Height = 14
            .Left = 172
            .Width = 160
            .Top = h
            .Enabled = False
        End With
        ArrModule(i) = ThisWorkbook.Sheets("temp").Cells(i, "ah").Value
        ArrValue(i) = strx
        strx = ""
        h = h + 18
    Next
    EditSave = True
    Set lb = Nothing
    Set textbox = Nothing
End Sub

Private Sub UserForm_Terminate()
    Dim yesno As Variant
    If Statisticsx = 1 Then Exit Sub
    If wd Is Nothing Then Exit Sub
    If EditSave = False Then
        yesno = MsgBox("是否保存修改", vbYesNo, "Tips")
        If yesno = vbNo Then
            If filex Like "ppt*" Then '不保存
                wd.Saved = True
                wd.Close
            Else
                wd.Close savechanges:=False
            End If
        Else
            If filex Like "ppt*" Then
                wd.Save
                wd.Close
            Else
                wd.Close savechanges:=True 'ppt文件不支持这个属性
            End If
        End If
    Else
        If filex Like "ppt*" Then
            wd.Saved = True
            wd.Close
        Else
            wd.Close savechanges:=False
        End If
    End If
    Set wd = Nothing
End Sub
