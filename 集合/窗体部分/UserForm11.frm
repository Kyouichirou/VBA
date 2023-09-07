VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm11 
   Caption         =   "初始化完成"
   ClientHeight    =   3915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9075
   OleObjectBlob   =   "UserForm11.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Dim strx As String
    strx = ThisWorkbook.Sheets("temp").Range("ab2").Value
    Call OpenFileLocation(strx) '打开程序所在的新位置
    ThisWorkbook.Close savechanges:=True
End Sub

Private Sub UserForm_Initialize()
    Dim arr() As Variant, i As Byte
    
    If Statisticsx = 1 Then Exit Sub
    With ThisWorkbook.Sheets("temp")
        i = .[aa65536].End(xlUp).Row
        arr = .Range("ab1:ac" & i).Value
    End With
    With Me
        .Label10.Caption = arr(23, 1)
        .Label11.Caption = arr(24, 1)
        
        If Len(arr(4, 1)) > 0 Then      'Powershell
            If Len(arr(4, 2)) > 0 Then
                .Label12.Caption = "支持"
            Else
                .Label12.Caption = "版本太低"
            End If
        Else
            .Label12.Caption = "不支持"
        End If
        
        If Len(arr(5, 1)) > 0 Then
            .Label13.Caption = "支持"
        Else
            .Label13.Caption = "不支持"
        End If
        
        If Len(arr(6, 1)) > 0 Then         'IE
            If Len(arr(6, 2)) > 0 Then
                .Label14.Caption = "支持"
            Else
                .Label14.Caption = "版本太低"
            End If
        Else
            .Label14.Caption = "不支持"
        End If
        
        If Len(arr(8, 1)) > 0 Then
            .Label15.Caption = "完整"
        Else
            .Label15.Caption = "不完整"
        End If
        
        If Len(arr(7, 1)) > 0 Then
            .Label16.Caption = "完整"
        Else
            .Label16.Caption = "不完整"
        End If
        
        If Len(arr(17, 1)) > 0 Then
            .Label17.Caption = "完整"
        Else
            .Label17.Caption = "不完整"
        End If
        
        If Len(arr(25, 1)) > 0 Then
            .Label18.Caption = "完整"
        Else
            .Label18.Caption = "不完整"
        End If
        
        If Len(arr(32, 1)) > 0 Then 'zip
            .Label21.Caption = "支持"
        Else
            .Label21.Caption = "不支持"
        End If
        
        If Len(arr(10, 1)) > 0 Then 'chrome
            .Label22.Caption = "支持"
        Else
            .Label22.Caption = "不支持"
        End If
        
        .TextBox1.Text = "1. 本程序仅用于交流学习使用,请勿用于商业用途." & vbCr & "2. 本程序不包含任何恶意代码." & vbCr & _
        "3. 本程序不确保用户在使用过程中造成的意外损失,尽管程序经过严格测试,但无法保证软件的Bug对用户不造成危害(如Md5算法获得的Hash值,将会用于比较文件,相同的文件将被删除), 使用前请仔细评估风险." & vbCr _
        & "4. 由于引用的代码的来源太过于广泛,无法一一标注原出处, 再此对所有的开源代码的作者表示集中感谢." & vbCr _
        & "5. 转载或者是二次修改,希望能保留出处."
    End With
    
    Erase arr
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Exit Sub
    Cancel = False
End Sub
