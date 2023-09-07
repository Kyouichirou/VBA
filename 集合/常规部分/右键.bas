Attribute VB_Name = "右键"
Option Explicit

Sub MyNewMenu()              '创建菜单
Dim mymenu As Object
Dim CM1 As Variant, CM2 As Variant, CM3 As Variant, CM4 As Variant, CM5 As Variant, CM6 As Variant, CM7 As Variant, CM8 As Variant

    With ThisWorkbook.Application.CommandBars("cell")
        .Reset
        Set mymenu = .Controls.Add(type:=msoControlPopup, Before:=1)
        .Controls(2).BeginGroup = True
    End With
    
    With mymenu
        .Caption = "文件操作"          '根据需要可改变显示名称
        Set CM1 = .Controls.Add(type:=msoControlButton)
        Set CM2 = .Controls.Add(type:=msoControlButton)
        Set CM3 = .Controls.Add(type:=msoControlButton)
        Set CM4 = .Controls.Add(type:=msoControlButton)
        Set CM5 = .Controls.Add(type:=msoControlButton)
        Set CM6 = .Controls.Add(type:=msoControlButton)
        Set CM7 = .Controls.Add(type:=msoControlButton)
        Set CM8 = .Controls.Add(type:=msoControlButton)
    End With
    
    With CM1
        .Caption = "复制"        '根据需要可改变显示名称
        .OnAction = "Rfilecopy"          '根据需要可改变执行宏
        .FaceId = 266             '根据需要可改变显示图标
    End With
    
    With CM2
        .Caption = "移动"
        .OnAction = "RfileMove"
        .FaceId = 49
    End With
    
     With CM3
        .Caption = "重命名"
        .OnAction = "Filerename"
        .FaceId = 59 '图标
    End With
    
    With CM4
        .Caption = "删除文件"
        .OnAction = "Rdelefile"
        .FaceId = 79
    End With
    
    With CM5
        .Caption = "打开所在文件夹"
        .OnAction = "Ropenfilelocation"
        .FaceId = 89
    End With
    
     With CM6
        .Caption = "文件属性"
        .OnAction = "Filedata"
        .FaceId = 89
    End With
    
    With CM7
        .Caption = "添加到优先阅读"
        .OnAction = "Raddplist"
        .FaceId = 89
    End With
    
    With CM8
        .Caption = "打开文件"
        .OnAction = "Openfiletemp"
        .FaceId = 89
    End With
    
    Set mymenu = Nothing
End Sub

Sub ResetMenu()                      '重置菜单
    ThisWorkbook.Application.CommandBars("cell").Reset
End Sub

Sub Rfilecopy()  '文件复制
    Dim addrowx As Integer
    addrowx = Selection.Row
    With ThisWorkbook.Sheets("书库")
        If FileCopy(.Range("e" & addrowx).Value, .Range("c" & addrowx).Value, addrowx) = True Then
            .Label1.Caption = "文件复制成功"
        Else
            .Label1.Caption = "文件复制失败"
        End If
    End With
End Sub

Sub RfileMove() '文件移动
    FileMove
End Sub

Sub RdeleFile()                  '删除文件
    Dim addrowx As Integer, excodex As Byte, p As Byte
    Dim yesno As Variant
    
    With ThisWorkbook.Sheets("书库")
        If Len(Selection) = 0 Then Exit Sub
        addrowx = Selection.Row
        yesno = MsgBox("是否删除本地文件?_", vbYesNo) '如果无法正常连接存储文件的处理
        If yesno = vbNo Then DeleMenu (addrowx): Exit Sub
        If Len(.Cells(addrowx, "ab").Value) > 0 Then p = 1 '标记是否存在非ansi
        Call FileDeleExc(.Cells(addrowx, "e").Value, addrowx, p, 0, Cells(addrowx, "c").Value)
    End With
End Sub

Sub FileData() '文件属性
    Dim addrowx As Integer, strx As String, strx1 As String
    ThisWorkbook.Application.ScreenUpdating = False
    If Len(Selection) = 0 Then Exit Sub
    addrowx = Selection.Row
    With ThisWorkbook.Sheets("书库")
'    If fso.FileExists(.Cells(addrowx, "e").Value) = False Then
'        DeleMenu addrowx
'        MsgBox "文件不存在", vbCritical, "Warning"
'        Exit Sub
'    End If
        strx = .Cells(addrowx, 2)
        Set Rng = .Range("b" & addrowx)
    End With
    If UF3Show = 0 Then
        Load UserForm3
        Call ShowDetail(strx)
    ElseIf UF3Show = 1 Or UF3Show = 3 Then
        strx1 = UserForm3.Label29.Caption
        If UF4Show = 1 Then Unload UserForm4
        If Len(strx1) > 0 Then
            If strx1 <> strx Then Call ShowDetail(strx)
        End If
    End If
    ThisWorkbook.Application.ScreenUpdating = True
    UserForm3.Show
End Sub

Sub FileReName()                     '文件重命名
    If Len(Selection) = 0 Then Exit Sub
    If UF3Show = 1 Then Unload UserForm3
    If UF3Show = 3 Then Unload UserForm4
    UserForm9.Show
End Sub

Sub RopenFileLocation()  '打开文件所在位置
    Call OpenFileLocation(Range("f" & Selection.Row))
End Sub

Sub RAddPlist() '添加到优先阅读
    Dim addrowx As Integer, strx As String
    
    addrowx = Selection.Row
    With ThisWorkbook.Sheets("书库")
        strx = .Range("b" & addrowx).Value
        If Len(strx) = 0 Then Exit Sub
        Call AddPList(strx, .Range("c" & addrowx).Value, 0)
    End With
End Sub

Sub OpenFileTemp() '打开文件-在书库中打开文件
    Dim slr As Integer, i As Byte, strx As String
    Dim addressx As String
    Dim filext As String
    
    With ThisWorkbook.Sheets("书库")
        If Len(Selection) = 0 Then Exit Sub '修改
        slr = Selection.Row
        If slr < 6 Then Exit Sub
        addressx = .Range("e" & slr).Value '路径
        filext = LCase(.Range("d" & slr).Value) '扩展名
        If filext Like "xl*" Then
            If filext <> "xls" And filext <> "xlsx" Then MsgBox "禁止打开此类文件", vbCritical, "Warning": Exit Sub
        End If
        If Len(.Cells(slr, "ak").Value) = 0 Then
            i = FileTest(addressx, filext, .Cells(slr, "c").Value)
            Select Case i
                Case 1: strx = "未获取到有效值"
                Case 2: strx = "文件不存在": Call DeleMenu(slr) '删除表格目录
                Case 3: strx = "文件是txt文件"
                Case 4: strx = "文件处于打开的状态"
                Case 5: strx = "文件处于打开的状态"
                Case 7: strx = "进程异常"
            End Select
        Else
            If Not filext Like "xls*" Then
                If WmiCheckFileOpen(addressx) = True Then .Label1.Caption = "文件已打开": Exit Sub
            End If
        End If
        If i = 0 Or i = 3 Or i = 6 Then
            If i = 6 Then .Cells(slr, "ak") = 1 '标记文件具有密码保护
            Set Rng = .Range("b" & slr)
            OpenFile .Range("b" & slr).Value, filext, .Range("d" & slr).Value, addressx, 2, .Range("ab" & slr).Value: Exit Sub
        Else
            .Label1.Caption = strx
        End If
    End With
End Sub

Function CheckRightMenu() As Boolean
    If ActiveWorkbook.Name <> ThisWorkbook.Name Then
        CheckRightMenu = False
        Exit Function
    Else
        If ThisWorkbook.ActiveSheet.Name <> "书库" Then
            CheckRightMenu = False
            Exit Function
        Else
            If Selection.Row < 6 Or Selection.Column < 1 Then
                CheckRightMenu = False
            Else
                CheckRightMenu = True
            End If
        End If
    End If
End Function
