Attribute VB_Name = "模块7"
Option Explicit
Enum SpeechStreamFileMode
    SSFMOpenForRead = 0
    SSFMOpenReadWrite = 1
    SSFMCreate = 2
    SSFMCreateForWrite = 3
End Enum
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Sub textv()
Dim a As New cJSON
a.parse
End Sub
Function textvx(ByVal Inputfile As String) '将文本转换为语音
    Dim oFileStream As Object, oVoice As Object, oFileOpen As Object
    Dim cs As New cSAPI
    Dim a As New SpeechLib.SpVoice
    Dim b As New SpeechLib.SpFileStream
    Dim d As New SpeechLib.SpVoice
    Dim c As ISpeechBaseStream
    Dim e
'    Set oFileOpen = CreateObject("SAPI.SpFileStream")
'    oFileOpen.Open Inputfile, SSFMOpenForRead, False
''    Set oVoice = CreateObject("SAPI.SpVoice")
'    oVoice.SpeakStream oFileOpen, 1

Dim ad As New ADODB.Parameter

     a.Speak Inputfile, SVSFlagsAsync
     a.Pause
     Set a = Nothing
     d.Speak Byte2String("C:\Users\adobe\Downloads\temp.txt", , 2048), SVSFlagsAsync
'     b.Open Inputfile, SSFMOpenForRead, False
'     b.Seek 2048, SSSPTRelativeToCurrentPosition
'     a.SpeakStream b, SVSFlagsAsync
     
'     b.Read(
'     c.Seek(1024,SSSPTRelativeToStart)
     
'     b.Seek
'    cs.Vox.SpeakStream oFileOpen, SVSFlagsAsync
'    SpEnd = False
'    Do Until SpEnd = True
'    Sleep 25
'    DoEvents
'    Loop
'    b.Close
'    Set oVoice = Nothing
'    Set oFileOpen = Nothing
'    Set cs = Nothing
End Function

'https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/readtext-method?view=sql-server-ver15
'https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/streamreadenum?view=sql-server-ver15


Private Function GetString(ByVal FilePath As String, Optional ByVal sCharset As String = "gb2312", Optional ByVal sPosition As Long = 0) As String 'ByRef bContent() As Byte
    Const adTypeBinary As Byte = 1
    Const adTypeText As Byte = 2
    Const adModeRead As Byte = 1
    Const adModeWrite As Byte = 2
    Const adModeReadWrite As Byte = 3
    Dim oStream As Object
    '----------------------从指定的位置读取信息
    Set oStream = CreateObject("ADODB.Stream")
    With oStream
        .Mode = 3
        .type = adTypeText
        .CharSet = sCharset
        .Open
        .LoadFromFile (FilePath)
        .Position = sPosition
         GetString = .ReadText()
        .Close
    End With
    Set oStream = Nothing
End Function
