Attribute VB_Name = "ʱ���"
Option Explicit

Function Timezone() As Integer '��ȡ���ڵ�ʱ��
    '--------------------��Ҫ����outlook��ʵ�ֻ�ȡ���ڵ�ʱ��, ����ʱ�����ڵ�ʱ��Ϊ��8��, ƫ��Ϊ-480, ת�� /60
    On Error Resume Next
    Dim outpp As Object
    Set outpp = CreateObject("Outlook.Application")
    If outpp Is Nothing Then MsgBox "�޷�����outlook����", vbCritical, "Warning": Exit Sub '����û��outlook
    Timezone = outpp.TimeZones.CurrentTimeZone.Bias
    Set outpp = Nothing
End Function

Function TimeStamp(Optional ByVal stamplen As Byte = 13) As String '����ʱ���
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
    TimeStamp = Left$(TimeStamp, stamplen) 'Ĭ�ϻ�ȡ13λ�ĳ���
    Set obj = Nothing
End Function

'��UNIXʱ���ת��Ϊ��׼ʱ��
'������intTime:Ҫת����UNIXʱ�����intTimeZone����ʱ�����Ӧ��ʱ��
'����ֵ��intTime������ı�׼ʱ��
'ʾ����FromUnixTime("1211511060", +8)������ֵ2008-5-23 10:51:0
Function FromUnixTime(intTime, intTimeZone) 'תΪ��׼ʱ��
    If IsEmpty(intTime) Or Not IsNumeric(intTime) Then
        FromUnixTime = Now()
        Exit Function
    End If
    If IsEmpty(intTime) Or Not IsNumeric(intTimeZone) Then intTimeZone = 0
    FromUnixTime = DateAdd("s", intTime, "1970-1-1 0:0:0")
    FromUnixTime = DateAdd("h", intTimeZone, FromUnixTime)
End Function

'�ѱ�׼ʱ��ת��ΪUNIXʱ���
'������strTime:Ҫת����ʱ�䣻intTimeZone����ʱ���Ӧ��ʱ��
'����ֵ��strTime�����1970��1��1����ҹ0�㾭��������
'ʾ����ToUnixTime("2008-5-23 10:51:0", +8)������ֵΪ1211511060
Function ToUnixTime(strTime, intTimeZone) 'תΪUNIXʱ���
    If IsEmpty(strTime) Or Not IsDate(strTime) Then strTime = Now
    If IsEmpty(intTimeZone) Or Not IsNumeric(intTimeZone) Then intTimeZone = 0
    ToUnixTime = DateAdd("h", -intTimeZone, strTime)
    ToUnixTime = DateDiff("s", "1970-1-1 0:0:0", ToUnixTime)
End Function
