Attribute VB_Name = "状态控制"
Option Explicit
'需要保持状态的sub,打开在多个工作簿,或者在本工作簿输入内容时会造成进程终止
Private Appspy As Eventspy

Sub EveSpy() '创建针对整个excel的事件监听-用于判断工作簿是否出现最小化
    Set Appspy = New Eventspy
End Sub

Function RecData() As Boolean '执行存储文件链接检查
    Dim FilePath As String
                                        '每次激活时,检查conn连接是否正常
    If Conn.State = adStateClosed Then
       On Error GoTo 100
       Set Conn = Nothing
       FilePath = ThisWorkbook.Sheets("temp").Range("ab3").Value
       If fso.fileexists(FilePath) = False Then                '检查文件是否存在
          RecData = False
          MsgBox "!数据连接异常,程序无法正常运行", vbCritical, "Warning"
          Exit Function
        End If
       Conn.Open "Provider = Microsoft.ACE.OLEDB.12.0;Data Source=" & FilePath & ";extended properties=""excel 12.0;HDR=YES""" '打开数据存储文件
       RecData = True
    End If
    RecData = True
    Exit Function
100
    Debug.Print Err.Number
    RecData = False '如果在链接过程中出现错误
End Function

Sub LockWorkSheet() '锁定表格禁止手动修改,只允许程序去修改
    Dim i As Byte, k As Byte
    
    With ThisWorkbook
        k = .Worksheets.Count
        For i = 1 To k
            .Worksheets(i).Protect "123", UserInterfaceOnly:=True
        Next
    End With
End Sub

Sub UnLockWorkSheet() '解锁表格
    Dim i As Byte, k As Byte
    
    With ThisWorkbook
        k = .Worksheets.Count
        For i = 1 To k
            .Worksheets(i).Unprotect ("123")
        Next
    End With
End Sub
