Attribute VB_Name = "时间戳"
Option Explicit

Function Timezone() As Integer '获取所在的时区
    '--------------------需要借助outlook来实现获取所在的时区, 北京时间所在的时区为东8区, 偏差为-480, 转换 /60
    On Error Resume Next
    Dim outpp As Object
    Set outpp = CreateObject("Outlook.Application")
    If outpp Is Nothing Then MsgBox "无法创建outlook对象", vbCritical, "Warning": Exit Sub '假如没有outlook
    Timezone = outpp.TimeZones.CurrentTimeZone.Bias
    Set outpp = Nothing
End Function

Function TimeStamp(Optional ByVal stamplen As Byte = 13) As String '生成时间戳
    Dim strx As String
    Dim obj As Object
    Set obj = CreateObject("MSScriptControl.ScriptControl")
    obj.AllowUI = True
    obj.Language = "JavaScript"
    strx = strx & "function abc()" & vbCrLf
    strx = strx & "{" & vbCrLf
    strx = strx & "var timestamp = new Date().getTime();" & vbCrLf
    strx = strx & "return timestamp;" & vbCrLf
    strx = strx & "}" & vbCrLf
    strx = strx & "abc()" & vbCrLf
    TimeStamp = obj.eval(strx)
    TimeStamp = Left$(TimeStamp, stamplen) '默认获取13位的长度
    Set obj = Nothing
End Function

'把UNIX时间戳转换为标准时间
'参数：intTime:要转换的UNIX时间戳；intTimeZone：该时间戳对应的时区
'返回值：intTime所代表的标准时间
'示例：FromUnixTime("1211511060", +8)，返回值2008-5-23 10:51:0
Function FromUnixTime(intTime, intTimeZone) '转为标准时间
    If IsEmpty(intTime) Or Not IsNumeric(intTime) Then
        FromUnixTime = Now()
        Exit Function
    End If
    If IsEmpty(intTime) Or Not IsNumeric(intTimeZone) Then intTimeZone = 0
    FromUnixTime = DateAdd("s", intTime, "1970-1-1 0:0:0")
    FromUnixTime = DateAdd("h", intTimeZone, FromUnixTime)
End Function

'把标准时间转换为UNIX时间戳
'参数：strTime:要转换的时间；intTimeZone：该时间对应的时区
'返回值：strTime相对于1970年1月1日午夜0点经过的秒数
'示例：ToUnixTime("2008-5-23 10:51:0", +8)，返回值为1211511060
Function ToUnixTime(strTime, intTimeZone) '转为UNIX时间戳
    If IsEmpty(strTime) Or Not IsDate(strTime) Then strTime = Now
    If IsEmpty(intTimeZone) Or Not IsNumeric(intTimeZone) Then intTimeZone = 0
    ToUnixTime = DateAdd("h", -intTimeZone, strTime)
    ToUnixTime = DateDiff("s", "1970-1-1 0:0:0", ToUnixTime)
End Function
