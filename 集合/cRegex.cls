VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cRegex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@ name: 正则细化
'@ function: 测试是否有匹配值, 替换匹配值, 返回单个匹配值, 返回多个匹配值, 返回多个匹配值子项, 返回单个匹配位置, 返回多个匹配位置, 返回匹配值累计数量
'@ update: 2020-05-29
'@ author: HLA
'@ version: 1.0
'@ status: stable
'@ isbug: n/a
'@ standby: 增加常用规则
'---------------------------------
'@ update: 2020-06-04
'@ 细化执行动作, 更好控制输出, 和输入内容, 容错能力增强
'---------------------------------------------detail
'----------structure
' -初始化-对象设置
' -检查输入的内容
' -执行
' -返回结果
' -获取属性
' -销毁
'---------structure
'注意事项:
'vbs不支持反向环向, 否定/肯定预查, 否则将会出错, 但是支持正向 '(?<=pattern) '(?<!pattern)
'javascript也不支持反向
'-------------------------------------------------------------------------------------------------
Private oRegE As Object
Private F_index As Long       '匹配值出现的第一个位置
Private F_arrindex() As Long  '匹配值所有出现的位置
Private mCount As Long        '匹配到的值
Private isReady As Boolean    '用于判断执行状态
Private isInitial As Boolean  '用于判断是否初始化
Private isPattern As Boolean  '用于判断设置匹配样式是否满足要求
Private Matches As Object     '匹配值集合
Private tmpString As String   '存储文本
Private isChange As Byte      '参数是否发生变化, 减少不必要的重复执行
Private isOK As Boolean       '检查输入的内容
'------------------------------
Private Const Url_Pattern As String = "([hH][tT]{2}[pP]://|[hH][tT]{2}[pP][sS]://|[wW]{3}.|[wW][aA][pP].|[fF][tT][pP].|[fF][iI][lL][eE].)[-A-Za-z0-9+&@#/%?=~_|!:,.;]+[-A-Za-z0-9+&@#/%=~_|]"
Private Const IP_Pattern As String = "((1\d{2}|25[0-5]|2[0-4]\d|[1-9]?\d)\.){3}(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)"
Private Const Zh_Pattern As String = "[\u4e00-\u9fa5]{1,}"  '匹配中文 '[^\x00-\xff] '双字节, 包含中文
Public Enum commPattern
    mUrl = 1
    mIP = 2
    mChinese = 3
End Enum
'------------------------------常用规则

Private Sub Class_Initialize()
    F_index = -1             '计算是从0开始, 故而初始值不能设置为0
    Set oRegE = CreateObject("VBScript.RegExp")
End Sub

Private Sub Class_Terminate()
    Set oRegE = Nothing
    F_index = 0
    mCount = 0
    Erase F_arrindex
    isReady = False
    isInitial = False: isPattern = False: isChange = 0: isOK = False
    tmpString = vbNullString
    Set Matches = Nothing
End Sub

Friend Property Get sFirst_Index() As Long '获取单项匹配项出现的位置
    sFirst_Index = F_index
End Property

Friend Property Get aFirst_Index() As Long() '获取多项目匹配项出现的位置
    aFirst_Index = F_arrindex
End Property

Friend Property Get oReg_Status() As Boolean '判断执行状态
    oReg_Status = isReady
End Property

Friend Property Get Match_Count() As Long '获取匹配项目
    If isChange < 2 Then oReg_Execute
    Match_Count = mCount
End Property
'------property 不支持参数为optional, 这里使用sub 来执行设置
Friend Sub oReg_Initial(Optional ByVal isGlobal As Boolean = True, Optional ByVal IgnoreC As Boolean = True, Optional ByVal isMultiLine As Boolean = True) '初始化_如果不开启全局,将只匹配单个, 这两个参数变化较少
    isInitial = True
    With oRegE
        .Global = isGlobal
        .IgnoreCase = IgnoreC    '\n匹配一个换行符 vbnewline。等价于 \x0a 和 \cJ。\r 匹配一个回车符。等价于 \x0d 和 \cM。
        .MultiLine = isMultiLine 'VBS 正则表达式 MultiLine 属性的这个特性可能会导致某些问题，所以在用 VBS 正则表达式处理文本时我一般会先把 "\r\n" 替换成 "\n"：这里需要注意大字符串,
                                 '正则表达式一个最重要的特性就是将匹配成功的模式的某部分进行存储供以后使用这一能力。请回想一下，
                                 '对一个正则表达式模式或部分模式两边添加圆括号将导致这部分表达式存储到一个临时缓冲区中。可以使用非捕获元字符 '?:', '?=', or '?!' 来忽略对这部分正则表达式的保存。
                                 '所捕获的每个子匹配都按照在正则表达式模式中从左至右所遇到的内容存储。存储子匹配的缓冲区编号从 1 开始，连续编号直至最大 99 个子表达式。
                                 '每个缓冲区都可以使用 '\n' 访问，其中 n 为一个标识特定缓冲区的一位或两位十进制数。
    End With
    isChange = 1
End Sub

Friend Property Let oReg_Text(ByRef strText As String)
    tmpString = strText
    If LenB(Trim$(tmpString)) = 0 Then isOK = False: isReady = False: isChange = 0: Exit Property
    isOK = True
    isChange = 1
End Property

Friend Property Let oReg_Pattern(ByVal sPattern As String) '设置匹配值
    Const BlockRule As String = "(?<"
    On Error GoTo ErrHandle
    Dim i As Integer
    i = Len(sPattern)
    isPattern = False
    If i = 0 Then
        Exit Property '还需要增加检查保留符号
    ElseIf i = 1 Then
        Select Case sPattern
            Case "\", ".", "+", "^", "$", "?", "*"  '注意保留字符 , 这些字符可以用ascw, asc等替换掉, w整体理论上更快一点, 因为w是处理Unicode的函数
            Exit Property
        End Select
    Else
        If InStr(1, Trim$(sPattern), BlockRule, vbBinaryCompare) = 1 Then Exit Property '不支持反向预查
        oRegE.Pattern = sPattern
        isPattern = True
        isChange = 1
    End If
    Exit Property
ErrHandle:
    isPattern = False
End Property

Friend Property Let comm_Pattern(ByVal cPattern As commPattern) '预设匹配规则
    Dim sPattern As String
    Select Case cPattern
        Case 1: sPattern = Url_Pattern
        Case 2: sPattern = IP_Pattern
        Case 3: sPattern = Zh_Pattern
    End Select
    isChange = 1
    oRegE.Pattern = sPattern
    isPattern = True
End Property

'--------------------------------------------类
'-----------执行
Friend Function ReplaceText(ByVal rpText As String) As String '替换 '可以使用$1, $2, 等符号提取特定的数据
    On Error GoTo ErrHandle
    Dim x
    If Check_isReady = False Then Exit Function
    If isChange = 2 And mCount < 1 Then Exit Function
    ReplaceText = oRegE.Replace(tmpString, rpText)
    Exit Function
ErrHandle:
    isReady = False
End Function

Friend Function cTest() As Boolean '判断是否存在匹配值
    '----检查是否有匹配值
    On Error GoTo ErrHandle
    If Check_isReady = False Then Exit Function
    If isChange = 2 And mCount < 1 Then Exit Function
    cTest = oRegE.Test(tmpString)
    Exit Function
ErrHandle:
    isReady = False
End Function

Friend Function sMatch() As String
    '--------------------------获取单个匹配值
    Dim match  As Object
    If isChange < 2 Then oReg_Execute
    If mCount < 1 Then Exit Function
    For Each match In Matches
        F_index = match.FirstIndex     '值从0开始
        sMatch = match.Value: Exit For
    Next
End Function

Friend Function xMatch() As String() ' 获取多个匹配值 & 出现的位置
    Dim i As Long
    Dim match As Object
    Dim arr() As String
    '-------------获取多个匹配值
    If isChange < 2 Then oReg_Execute
    i = mCount - 1
    If i < 0 Then Exit Function
    ReDim xMatch(i)
    ReDim arr(i)
    ReDim F_arrindex(i)
    i = 0
    For Each match In Matches
        F_arrindex(i) = match.FirstIndex
        arr(i) = match.Value: i = i + 1
    Next
    xMatch = arr
End Function

Friend Function xMatch_Index() As Long() ' 单独_获取多个匹配值出现的位置
    Dim i As Long
    Dim match As Object
    Dim arr() As Long
    '-------------获取多个匹配值
    If isChange < 2 Then oReg_Execute
    If mCount < 1 Then Exit Function
    i = mCount - 1
    ReDim xMatch_Index(i)
    ReDim arr(i)
    i = 0
    For Each match In Matches
        arr(i) = match.FirstIndex
        i = i + 1
    Next
    xMatch_Index = arr
End Function

Friend Function xSubmatch() As String() ' 获取主项+子项匹配
    Dim i As Long, p As Long, n As Long, m As Integer, j As Integer
    Dim match As Object
    Dim arr() As String
    '-------------获取多个子项匹配值, sPattern中需要匹配的子项, 可以用括号括起
    'sPartten="songID":+[\d]{6,} 换成("songID"):+([\d]{6,})
    
    If isChange < 2 Then oReg_Execute
    If mCount < 1 Then Exit Function
    n = mCount - 1
    ReDim arr(n, 0)
    ReDim xSubmatch(n, 0) '一维数组不能再次redim为二维
    ReDim F_arrindex(n)
    j = 0
    For Each match In Matches
        With match
            p = .SubMatches.Count
            If p > m Then m = p: ReDim Preserve arr(n, m): ReDim Preserve xSubmatch(n, m)
            arr(j, 0) = .Value
            F_arrindex(j) = .FirstIndex
            For i = 1 To p
                arr(j, i) = .SubMatches(i - 1)
            Next
            j = j + 1
        End With
    Next
    xSubmatch = arr
End Function

Private Sub oReg_Execute() '执行匹配
    On Error GoTo ErrHandle
    If isChange = 2 Then Exit Sub
    If Check_isReady = False Then Exit Sub
    Set Matches = oRegE.Execute(tmpString)
    mCount = Matches.Count
    isChange = 2
    isReady = True
    If mCount = 0 Then isReady = False     '无匹配值
    Exit Sub
ErrHandle:
    isReady = False
End Sub

Private Function Check_isReady() As Boolean '确保所有的基础参数都设置完成
    isReady = isOK * isInitial * isPattern
    Check_isReady = isReady
End Function
