VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cWMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'https://docs.microsoft.com/en-us/windows/win32/multimedia/multimedia-command-strings
'https://docs.microsoft.com/zh-cn/previous-versions/dd743680(v=vs.85)
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, _
ByVal lpstrReturnString As String, _
ByVal uReturnLength As Long, _
ByVal hwndCallback As Long) As Long
Private Declare Function GetShortPathNameA Lib "kernel32.dll" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private sAlias As String
Private sFilename As String
Private nLength As Single
Private nPosition As Single
Private sStatus As Single
Private bWait As Boolean
Private iRepeat As Byte

Friend Function wmOpen(ByVal FilePath As String) As Boolean
    Dim nReturn As Long
    Dim sType As String
    '注意文件的内容要和文件的扩展名相匹配(假如是mp3文件,如果改成wav.是无法播放的)
    If Len(sAlias) > 0 Then wmClose
    wmOpen = False
    Select Case UCase$(Right$(FilePath, 3))
    Case "WAV"
        sType = "Waveaudio"
    Case "AVI"
        sType = "AviVideo"
    Case "MID"
        sType = "Sequencer"
    Case "MP3"
        sType = "MPegVideo"
    Case Else
        Exit Function
    End Select
    Randomize
    sAlias = Right$(FilePath, 3) & Minute(Now) & Second(Now) & Int(1000 * Rnd + 1)
    If InStr(FilePath, " ") Then FilePath = Chr(34) & FilePath & Chr(34)
    nReturn = mciSendString("Open " & FilePath & " ALIAS " & sAlias & " TYPE " & sType & " wait", "", 0, 0)
    If nReturn = 0 Then wmOpen = True
End Function

Friend Sub wmClose()
    Dim nReturn As Long
    
    If sAlias = "" Then Exit Sub
    nReturn = mciSendString("Close " & sAlias, "", 0, 0)
    sAlias = ""
    sFilename = ""
End Sub

Friend Sub wmPause()
    Dim nReturn As Long
    
    If sAlias = "" Then Exit Sub
    nReturn = mciSendString("Pause " & sAlias, "", 0, 0)
End Sub

Friend Function wmPlay(Optional ByVal IsRepeat As Boolean) As Boolean
    Dim nReturn As Long
    
    If sAlias = "" Then Exit Function
    wmPlay = False
    If bWait Then
        nReturn = mciSendString("Play " & sAlias & " wait", "", 0, 0)
    Else
        If IsRepeat = True Then
            nReturn = mciSendString("Play " & sAlias & " repeat", "", 0, 0) '重复播放
        Else
            nReturn = mciSendString("Play " & sAlias, "", 0, 0)
        End If
    End If
    If nReturn = 0 Then wmPlay = True
End Function

Friend Sub wmResume() '对应 paused
    Dim nReturn As Long
    
    If sAlias = "" Then Exit Sub
    If Status <> "paused" Then Exit Sub
    If bWait Then
        nReturn = mciSendString("resume " & sAlias & " wait", "", 0, 0)
    Else
        nReturn = mciSendString("resume " & sAlias, "", 0, 0)
    End If
End Sub

Friend Sub wmVolume(ByVal sValue As String) ''调节播放音量,以1000为基准,但是不能调大,只能调小
    Dim nReturn As Long
    
    If sAlias = "" Then Exit Sub
    If Status <> "playing" Then Exit Sub
    If bWait Then
        nReturn = mciSendString("setaudio " & sAlias & " volume to " & sValue & " wait", "", 0, 0)
    Else
        nReturn = mciSendString("setaudio " & sAlias & " volume to " & sValue, "", 0, 0)
    End If
End Sub

Friend Sub wmSpeed(ByVal sValue As String) '调节速度,以1000为基准,可调快或慢
    Dim nReturn As Long
    
    If sAlias = "" Then Exit Sub
    If Status <> "playing" Then Exit Sub
    If bWait Then
        nReturn = mciSendString("set " & sAlias & " speed " & sValue & " wait", "", 0, 0)
    Else
        nReturn = mciSendString("set " & sAlias & " speed " & sValue, "", 0, 0)
    End If
End Sub

Friend Sub wmStop()
    Dim nReturn As Long
    
    If sAlias = "" Then Exit Sub
    nReturn = mciSendString("Stop " & sAlias, "", 0, 0)
End Sub

Friend Sub wmSeek(ByVal nPosition As Single)
    Dim nReturn As Long
    
    nReturn = mciSendString("seek " & sAlias & " to " & nPosition, "", 0, 0)
End Sub

Friend Sub wmRestart() '重新开始-对应stop
    Dim nReturn As Long
    
    If sAlias = "" Then Exit Sub
    If bWait Then
        nReturn = mciSendString("seek " & sAlias & " to start" & " wait", 0&, 0, 0)
    Else
        nReturn = mciSendString("seek " & sAlias & " to start", 0&, 0, 0)
    End If
    wmPlay
End Sub

Friend Sub wmMove(ByVal iValue As Integer) '移动播放位置按照百分比来确定位置
    Dim nReturn As Long
    Dim iLength As Long
    Dim iStart As Long
    
    If sAlias = "" Then Exit Sub
    If Status <> "playing" Then Exit Sub
    iLength = Length
    If iLength = 0 Then Exit Sub
    iStart = iValue / 100 * iLength
    If bWait Then
        nReturn = mciSendString("seek " & sAlias & " to start" & " wait", iStart, 0, 0)
        nReturn = mciSendString("play " & sAlias & " from " & CStr(iStart) & " to " & CStr(iLength) & " wait", 0&, 0, 0)
    Else
        nReturn = mciSendString("seek " & sAlias & " to start", 0&, 0, 0)
        nReturn = mciSendString("play " & sAlias & " from " & CStr(iStart) & " to " & CStr(iLength), 0&, 0, 0)
    End If
End Sub

Friend Function wmMute(ByVal IsOff As Boolean) As Boolean '静音
    Dim nReturn As Long
    Dim OnOff As String
    Dim sAudio As String * 255
    Dim strTemp As String
    Dim nLength As Integer
    
    wmMute = False
    If sAlias = "" Then Exit Function
    If Status <> "playing" Then Exit Function
    nReturn = mciSendString("Status " & sAlias & " audio", sAudio, 255, 0)
    nLength = InStr(sAudio, Chr$(0))
    strTemp = Left$(sAudio, nLength - 1)
    If nReturn <> 0 Then Exit Function
    If IsOff = True Then
        If strTemp = "off" Then Exit Function
        OnOff = "off"
    Else
        If strTemp = "on" Then Exit Function
        OnOff = "on"
    End If
    nReturn = mciSendString("setaudio " & sAlias & " " & OnOff, vbNull, 0, 0)
    If nReturn = 0 Then wmMute = True
End Function
'-------------------------------------------------------class
Property Get FileName() As String
    FileName = sFilename
End Property

Property Let FileName(ByVal sTheFile As String)
    wmOpen sTheFile
End Property
'-------------------------------https://docs.microsoft.com/en-us/windows/win32/multimedia/the-wait-flag
Property Get Wait() As Boolean
    Wait = bWait
End Property

Property Let Wait(bWaitValue As Boolean) '延时执行
    bWait = bWaitValue
End Property

Property Get Length() As Single '获取文件的长度
    Dim nReturn As Long, nLength As Integer
    Dim sLength As String * 255
    
    If sAlias = "" Then
        Length = 0
        Exit Property
    End If
    nReturn = mciSendString("Status " & sAlias & " length", sLength, 255, 0)
    nLength = InStr(sLength, Chr$(0))
    Length = Val(Left$(sLength, nLength - 1))
End Property

Property Let Position(ByVal nPosition As Single)
    wmSeek nPosition
End Property

Property Get Position() As Single
    Dim nReturn As Integer, nLength As Integer
    Dim sPosition As String * 255
    
    If sAlias = "" Then Exit Property
    nReturn = mciSendString("Status " & sAlias & " position", sPosition, 255, 0)
    nLength = InStr(sPosition, Chr$(0))
    Position = Val(Left$(sPosition, nLength - 1))
End Property

Property Get Volume() As Single
    Dim nReturn As Integer, nLength As Integer
    Dim sVolume As String * 255
    
    If sAlias = "" Then Exit Property
    If Status <> "playing" Then Exit Property
    nReturn = mciSendString("Status " & sAlias & " volume", sVolume, 255, 0)
    nLength = InStr(sVolume, Chr$(0))
    Volume = Val(Left$(sVolume, nLength - 1))
End Property

Property Get Status() As String
    Dim nReturn As Integer, nLength As Integer
    Dim sStatus As String * 255
    
    If sAlias = "" Then Exit Property
    nReturn = mciSendString("Status " & sAlias & " mode", sStatus, 255, 0)
    nLength = InStr(sStatus, Chr$(0))
    Status = Left$(sStatus, nLength - 1)
End Property

Private Sub Class_Terminate()
    If Len(sAlias) = 0 Then Exit Sub
    wmClose
End Sub
