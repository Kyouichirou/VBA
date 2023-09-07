Attribute VB_Name = "MD5A"
Option Explicit

Function GetFileHashMD5(ByVal FilePath As String, Optional ByVal errx As Integer) As String '计算md5速度最好的方法,支持非ansi字符路径
    Dim Filehashx As Object
    Dim WDI As Object                                          '可以计算大于2G以上的文件,单个文件测试最大12+G
    Dim HashValue As String
    Dim i As Integer       'https://docs.microsoft.com/zh-cn/windows/win32/msi/msifilehash-table
    Dim k As Byte, j As Byte, m As Byte
    
    On Error GoTo ErrHandle '出错直接调用其他的模块
    Set WDI = CreateObject("WindowsInstaller.Installer") ''https://docs.microsoft.com/zh-cn/windows/win32/msi/installer-object
    Set Filehashx = WDI.FileHash(FilePath, 0)           '核心
    If WDI Is Nothing Or Filehashx Is Nothing Then GoTo ErrHandle '假设对象没有创建成功'调用其他的md5模块进行计算
    k = Filehashx.FieldCount '4
    For i = 1 To k
        HashValue = HashValue & BigEndianHex(Filehashx.IntegerData(i))
    Next
    GetFileHashMD5 = HashValue
CheckLen:
    j = Len(GetFileHashMD5)   ' m用于控制执行的次数
    If j <> 32 And j <> 2 Then
        If m = 0 Then
            GoTo ErrHandle
        Else
            GetFileHashMD5 = "UC"
        End If
    End If
    Set Filehashx = Nothing
    Set WDI = Nothing
    Exit Function
ErrHandle:
    m = m + 1
    If errx = 0 Then errx = ErrCode(FilePath, 1)
    GetFileHashMD5 = GetMD5Hash_File(FilePath, errx) '调用其他md模块进行计算
    Resume CheckLen
End Function

Private Function BigEndianHex(ByVal xl As Long) As String 'https://blog.csdn.net/weixin_42066185/article/details/83755433
    Dim Result As String
    Dim strx1 As String * 2, strx2 As String * 2, strx3 As String * 2, strx4 As String * 2
    '-------------------------------------https://stackoverrun.com/ja/q/8312292
    '-----------https://docs.microsoft.com/zh-CN/office/vba/api/excel.application.worksheetfunction
    '-----------https://docs.microsoft.com/zh-CN/office/vba/api/excel.worksheetfunction.dec2hex
    '-----------Result = ThisWorkbook.Application.WorksheetFunction.Dec2Hex(xl, 8) '会出现超出8位的问题
    Result = Hex(xl) '-----------------------------------------------最多八个十六进制字符
    If Len(Result) < 8 Then Result = Right$("00000000" & Result, 8) '补位
    strx1 = Mid$(Result, 7, 2)
    strx2 = Mid$(Result, 5, 2)
    strx3 = Mid$(Result, 3, 2)
    strx4 = Mid$(Result, 1, 2)
    BigEndianHex = strx1 & strx2 & strx3 & strx4
End Function
