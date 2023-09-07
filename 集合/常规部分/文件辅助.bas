Attribute VB_Name = "文件辅助"
Option Explicit
Option Compare Text
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) '时间 -单词训练
Private Declare Function timeGetTime Lib "winmm.dll" () As Long '时间 -单词训练
Private Declare Function CharUpper Lib "user32.dll" Alias "CharUpperA" (ByVal lpsz As String) As String                              '将字符编码统一转换为ansi(大写)

Public errcodenx As String '非ansi编码字符出现在字符串的的位置
Public errcodepx As String
Public Tagfnansi As Boolean, Tagfpansi As Boolean '标记哪个位置出现非ansi编码的字符
Private Declare Function IsTextUnicode Lib "Advapi32.dll" (ByVal intP As Long, ByVal sBuffer As Long) As Long



Sub utetst()
Dim s As String
s = Cells(20, 2).Value
Debug.Print IsTextUnicode(StrPtr(s), LenB(s))
'Dim arr() As Byte
'arr = StrConv(s, vbFromUnicode)
'Cells(21, 2) = CharUpper(s)
End Sub


Function ErrCode(ByVal strFilen As String, ByVal Exccode As Byte, Optional ByVal strFilep As String) As Integer '检查文件名是否存在异常字符
    Dim strFile As String
    Dim i As Byte, n As Byte, p As Byte, m As Byte   'Exccode用于区分不同来源的请求'strfilen,strfilep路径,需要确定出现ansi编码的位置是文件名还是路径
    Dim arrn() As String, arrp() As String, k As Byte
    Dim strx As String, strx1 As String
    
    ErrCode = 1                                                   'If InStr(Mid(strFile, i, 1), CharUpper(Mid(strFile, i, 1))) = 0 Then
    Tagfnansi = False '使用前进行参数重置
    Tagfpansi = False
    errcodenx = ""
    errcodepx = ""
    
    n = Len(strFilen)
    p = Len(strFilep)
    
    If p = 0 And n = 0 Then '极其特殊的情况下会出现无法读取路径文件的问题
        ErrCode = -1
        Exit Function
    End If
    
    If n > 0 Then ReDim arrn(1 To n)
    If p > 0 Then ReDim arrp(1 To p)
    
    If Exccode = 1 Then ''比较哪个部分的字符长度长,先计算较短的,节省时间,因为要分别获取不同状态下的非ansi,1的时候只需要发现存在异常字符,不需要知道字符出现在什么位置,0的时候需要获取文件路径不同部分的字符存在异常字符的状况
        If p < n Then
            m = 1
        Else
            m = 2
        End If
    End If
    
    If m = 1 Or Exccode = 0 Then
100
        strFile = strFilen
        For i = 1 To n
            strx = Mid$(strFile, i, 1)
            strx1 = strx                  '注意这里不能同时使用strx
            If InStr(strx, CharUpper(strx1)) = 0 Then 'mid/left/rgiht 后面的$表示里面的内容按照string数据类型进行处理
                ErrCode = ErrCode + 1
                If Exccode = 1 Then Exit Function '这样可以加快代码的运转速度,但是无法将异常字符的位置全部标注出来
                If Tagfnansi = False Then Tagfnansi = True
                arrn(i) = i '暂时存储标记出异常字符在路径中出现的位置
            End If
        Next
        If m = 1 And Tagfpansi = False Then GoTo 101
    End If
    
    m = 0 '重置参数, 防止m=2,没有异常字符的时候形成死循环
    
    If m = 2 Or Exccode = 0 Then
101
        strFile = strFilep
        For k = 1 To p
            strx = Mid$(strFile, k, 1)
            strx1 = strx
            If InStr(strx, CharUpper(strx1)) = 0 Then
                ErrCode = ErrCode + 1
                If Exccode = 1 Then Exit Function '这样可以加快代码的运转速度,但是无法将异常字符的位置全部标注出来
                If Tagfpansi = False Then Tagfpansi = True
                arrp(k) = k '暂时存储标记出异常字符在路径中出现的位置
            End If
        Next
        If m = 2 And Tagfnansi = False Then GoTo 100 '找不到继续查找
    End If
    
    If ErrCode > 0 And Exccode = 0 Then
        If Tagfpansi = True Then errcodepx = Trim(Join(arrp, " ")) '数据合并起来(注意数组需要是字符串类型)
        If Tagfnansi = True Then errcodenx = Trim(Join(arrn, " "))
    End If
End Function

Function OpenBy(ByVal FilePath As String) As String '获取文件类型默认关联程序
    Dim str$, Result$
    
    str = LCase(Right(FilePath, Len(FilePath) - InStrRev(FilePath, ".") + 1)) '带.号的文件后缀名
    With CreateObject("wscript.shell")
        On Error Resume Next
        Result = .RegRead("HKEY_CLASSES_ROOT\" & str & "\") '获取后缀名对应的注册类型
        If Len(Result) > 0 Then
            Result = .RegRead("HKEY_CLASSES_ROOT\" & Result & "\shell\open\command\") '由注册类型找到打开的程序路径
            If Result Like """*" Then Result = Split(Result, """")(1) Else Result = Split(Result, " ")(0)
            OpenBy = Result
        End If
    End With
End Function

Function CheckFileFrom(ByVal xpath As String, ByVal cmCode As Byte) As Boolean '检查文件是否来源于系统文件夹
    Dim strx As String
    Dim strdsk As String, strdlw As String, strdcm As String, struserfile As String
    
    CheckFileFrom = False
    If cmCode = 1 Then     '表示文件
        strx = fso.GetFile(xpath).ParentFolder & "\"
    ElseIf cmCode = 2 Then
        strx = xpath & "\"     '表示文件夹
    End If
    
    If fso.GetDriveName(xpath) = Environ("SYSTEMDRIVE") Then '如果文件所在磁盘和系统位于同一个盘
        struserfile = Environ("UserProfile") '用户文件夹
        strdsk = struserfile & "\Desktop\"
        strdlw = struserfile & "\Downloads\"
        strdcm = struserfile & "\Documents\" '限制文件只能来自系统盘的三个位置-桌面-下载-文档
        If InStr(strx, strdsk) = 0 And InStr(strx, strdlw) = 0 And InStr(strx, strdcm) = 0 Then CheckFileFrom = True
    End If
End Function
'设置了word的写入密码/打开密码,openstream并不会出现错误, 无法通过此方法来检测word文件是否设置有密码
'部分的pdf加密会出现openstream无法打开
Function FileStatus(ByVal filecodex As String, Optional ByVal cmCode As Byte) As Byte '判断文件是否存在/重名excel/是否处于打开的状态     'https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/openastextstream-method
    Dim i As Byte, k As Byte
    Dim address As String, strx As String              '用于判断是否需要打开文件
    Dim fl As File, flop As Object, filex As String
                                                        'Boolean-可改成byte这里可以调节可以根据同的阶段出现的内容来给fileexist赋值,如1,2,3分别代表不同的含义
    On Error GoTo 100
    Call SearchFile(filecodex) '查找文件是否存在于目录
    If Rng Is Nothing Then
        FileStatus = 1 '文件不存在于目录
    Else
        If cmCode = 1 Then FileStatus = 2: Exit Function
        address = Rng.Offset(0, 3).Value '存在于目录
        If fso.fileexists(address) = False Then
            FileStatus = 3
            Call DeleMenu(Rng.Row) '如果文件不存在则执行清除目录信息
        Else
            If cmCode = 2 Then FileStatus = 4: Exit Function '不执行后面的判断-只判断文件是否存在于本地以及目录
            filex = LCase(Rng.Offset(0, 2).Value)
            If filex Like "xl*" Then  '如果文件的类型是excel,那么判断打开的文件是否重名/或者是本文件,因为Excel无法打开重名的文件
                k = Workbooks.Count
                strx = Rng.Offset(0, 1).Value
                For i = 1 To k
                    If strx = Workbooks(i).Name Then FileStatus = 5: Exit Function
                Next
            Else
                If Len(Rng.Offset(0, 35).Value) > 0 Then '文件带有密码
                    If WmiCheckFileOpen(address) = False Then
                        FileStatus = 0
                    Else
                        FileStatus = 6
                    End If
                    Set fl = Nothing
                    Set flop = Nothing
                    Exit Function
                End If
                If filex <> "txt" Then '判断文件是否处于打开的状态,支持非ansi字符路径
                    Set fl = fso.GetFile(address) '获取文件对象
                    Set flop = fl.OpenAsTextStream(ForAppending, TristateUseDefault) '注意这里不要选forwriting参数,否则会彻底损坏文件 ,ForAppending表示在最后一行的位置准备写入信息
                    flop.Close
                End If
            End If
        End If
    End If
    Set fl = Nothing
    Set flop = Nothing
    Exit Function
100
    If Err.Number = 70 Then
       FileStatus = 7     '处于打开的状态
       If WmiCheckFileOpen(address) = False Then FileStatus = 0: Rng.Offset(0, 35) = 1 '标记文件是否有密码保护
    Else
        FileStatus = 8 '出现其他的错误
    End If
    Err.Clear
    Set fl = Nothing
    Set flop = Nothing
End Function

Function FileTest(ByVal FilePath As String, ByVal filex As String, ByVal FileName As String) As Byte '仅判断文件是否处于打开的状态/和判断文件是否存在是并行的
    Dim fl As File, i As Byte, k As Byte, c As Byte
    Dim flop As Object
    Dim errx As Integer
    '-----------------------OpenAsTextStream这种方法有个较大的局限,那就是无法判断密码保护的文件的打开状态,例如有密码保护的pdf文件, txt文件也无法判断
    On Error GoTo 100
    FileTest = 0 '初始值
    If Len(FilePath) = 0 Then FileTest = 1: Exit Function '传递过来的值为空或者无法有效读取数据
  
    If fso.fileexists(FilePath) = False Then FileTest = 2: Exit Function
    
    filex = LCase(filex)
    If filex = "txt" Then FileTest = 3: Exit Function
    If filex Like "xl*" Then
        i = Workbooks.Count
        For k = 1 To i
            If FileName = Workbooks(k).Name Then FileTest = 4
        Next
        c = 1
    End If
    Set fl = fso.GetFile(FilePath)
    Set flop = fl.OpenAsTextStream(ForAppending, TristateUseDefault) '通过判断文件的可访问状态(是否锁定)
    flop.Close
    Set fl = Nothing
    Set flop = Nothing
    Exit Function
100
    errx = Err.Number
    If errx = 70 Then
        FileTest = 5 '文件处于打开的状态
        If c <> 1 Then
            If WmiCheckFileOpen(FilePath) = False Then FileTest = 6 '文件具有密码保护
        End If
    Else
        FileTest = 7
    End If
    Err.Clear
    Set fl = Nothing
    Set flop = Nothing
End Function
