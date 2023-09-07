VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm9 
   Caption         =   "修改文件名"
   ClientHeight    =   3480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8730
   OleObjectBlob   =   "UserForm9.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text  '不区分大小写
Option Explicit
Dim addrowx As Integer '缓存
Dim standfn As String
Dim standfp As String, standfp1 As String
Dim standx1 As String
Dim standx2 As String
Dim standx3 As String
Dim standx4 As String
Dim standx5 As String
Dim standx6 As String

Private Sub CommandButton1_Click() '修改文件名
    Dim newname As String
    Dim str1 As String
    
    On Error GoTo 100 '错误处理
    If Len(Trim(Me.TextBox1.Text)) = 0 Or Len(standfn) = 0 Then Exit Sub '未输入内容
    
    With ThisWorkbook.Sheets("书库")
        If fso.fileexists(standfp) = False Then '文件为空
            Me.Label1.Caption = "文件不存在"
            Exit Sub
        End If
        
        newname = Me.TextBox1.Text
        str1 = newname & "." & .Range("d" & addrowx)  '带有扩展名的新文件名
        
        If fso.fileexists(.Range("f" & addrowx) & "\" & str1) = True Then '判断文件是否重名
            Me.Label1.Caption = "文件重名"
            Me.TextBox1.Text = ""
            Exit Sub
        End If
        
        If ErrCode(newname, 1) > 1 Then '当文件不重名就去检查是否存在异常字符
            Me.Label1.Caption = "输入的文件名存在异常字符" '修改文件名的是否存在异常字符''如果输入的名称出现重名/存在空格,将重新执行输入-检查循环(inputbox)
            Me.TextBox1.Text = ""
            Exit Sub
        Else
            fso.GetFile(standfp).Name = str1 '操作文件的能使用fso的全部使用fso,不会出现非ansi字符的问题
        End If
        .Range("c" & addrowx) = str1 '新的文件名
        str1 = .Range("f" & addrowx) & "\" & str1
        .Range("e" & addrowx) = str1 '新的路径
        standfp1 = str1
        .Range("ae" & addrowx) = "" '如果原来的名字存在异常字符那么就去掉这个标记
        
        If .Range("ac" & addrowx) = "EDC" Then '文件夹/文件都存在非ansi编码
            .Range("ac" & addrowx) = "EPC" '只有文件夹部分存在特殊字符
        Else
            .Range("ac" & addrowx) = ""
            .Range("ab" & addrowx) = ""
        End If
        If Len(standx5) > 0 Then ThisWorkbook.Sheets("目录").Cells(standx5, standx6) = Now
        Me.Label1.Caption = "修改成功!"
        Me.CommandButton2.Enabled = True '
        If .Range("d" & addrowx) = "txt" Then Me.Label2.Caption = "txt文档可以在打开的状态下移动\重命名\删除"      '提醒用户,txt文档可以在打开的状态下移动和重命名等操作
        End With
    Exit Sub
100
    If Err.Number = 70 Then '70文件打开,76文件目录不存在
        Me.Label1.Caption = "文件处于打开状态,改名失败"
    Else
        Me.Label1.Caption = "操作异常,请重试"
    End If
    Err.Clear
End Sub

Private Sub CommandButton2_Click() '恢复原有的文件名
    With ThisWorkbook.Sheets("书库")
    If fso.fileexists(standfp1) = True Then
        fso.GetFile(standfp1).Name = standfn '恢复本地文件名称
    Else
        DeleMenu (addrowx)
        Me.Label1.Caption = "文件已经丢失"
        Exit Sub
    End If
        .Range("c" & addrowx) = standfn
        .Range("e" & addrowx) = standfp
        .Range("ae" & addrowx) = standx1
        .Range("ab" & addrowx) = standx2
        .Range("ac" & addrowx) = standx3
        .Range("af" & addrowx) = standx4
        If Len(standx5) > 0 Then ThisWorkbook.Sheets("目录").Cells(standx5, standx6) = Now '文件重新被修改的时间
    End With
End Sub

Private Sub CommandButton3_Click() '主文件名
    With Me
        .TextBox1 = .Label3.Caption
    End With
End Sub

Private Sub CommandButton4_Click() '豆瓣文件名
    With Me
        .TextBox1 = .Label4.Caption
    End With
End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) '限制文件名出现这些系统禁止的字符
    With Me
        Select Case KeyAscii
            Case asc("/"), asc("\"), asc(":"), asc("*"), asc("?"), asc("<"), asc(">"), asc("|") 'MsgBox "请勿输入非法字符：""/\ : * ? <> |""", vbInformation, "提醒"
            .Label1.Caption = "请勿输入非法字符：""/\ : * ? <> |"
            .TextBox1.Text = ""
            KeyAscii = 0
        Case Else
           .Label1.Caption = ""
        End Select
    End With
End Sub

Private Sub UserForm_Initialize() '获取初始值,用于恢复
    Dim rngtime As Range
    Dim str As String, strx1 As String, strx2 As String
    
    If Statisticsx = 1 Then Exit Sub
    addrowx = Selection.Row
    If addrowx < 6 Then Exit Sub
    With ThisWorkbook.Sheets("书库")
        standfn = .Range("c" & addrowx).Value '文件名
        Me.Label2.Caption = standfn
        standfp = .Range("e" & addrowx).Value '文件路径
        standx1 = .Range("ae" & addrowx).Value '异常字符标记
        standx2 = .Range("ab" & addrowx).Value '异常字符标记
        standx3 = .Range("ac" & addrowx).Value
        standx4 = .Range("af" & addrowx).Value
        str = .Range("f" & addrowx).Value & "\"
        strx1 = .Range("o" & addrowx).Value '主文件名
        strx2 = .Range("y" & addrowx).Value '豆瓣名称
    End With
    
    With Me
        If Len(strx1) > 0 Then .Label3.Caption = strx1: .CommandButton3.Visible = True
        If Len(strx2) > 0 Then .Label4.Caption = strx2: .CommandButton4.Visible = True
    End With
    
    With ThisWorkbook.Sheets("目录")
        Set rngtime = .Cells(4, 3).Resize(.[b65536].End(xlUp).Row, .Cells.SpecialCells(xlCellTypeLastCell).Column).Find(str, lookat:=xlWhole)
        If Not rngtime Is Nothing Then
            standx5 = rngtime.Row
            standx6 = rngtime.Column
        End If
    End With
End Sub
