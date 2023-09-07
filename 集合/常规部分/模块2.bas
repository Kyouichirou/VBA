Attribute VB_Name = "Ä£¿é2"
Option Explicit

Const FILENAME1 = "C:\Users\adobe\Downloads\First.wav"
Const FILENAME2 = "C:\Users\adobe\Downloads\Second.wav"

Dim V As SpeechLib.SpVoice
Dim S1 As SpeechLib.SpFileStream
Dim S2 As SpeechLib.SpFileStream

Private Sub Command1_Click()

    Dim varT(3) As Variant      'text to be spoken
    Dim varP(3) As Variant      'positions in output stream
    Dim varD(3) As Variant      'audio data chunks
    Dim varStart As Variant
    Dim ii As Integer



    varT(0) = "one": varT(1) = "point": varT(2) = "five": varT(3) = "SAPI"

    'Create WAV file of "one point five SAPI"
    'Speak the words into a single filestream object,
    'and remember the end-of-stream position of each word.

    S1.Open FILENAME1, SSFMCreateForWrite
    Set V.AudioOutputStream = S1
    V.Speak "Write them back into the second file in reverse orde"
    For ii = 0 To UBound(varT)
        V.Speak varT(ii)
        varP(ii) = S1.Seek(0, SSSPTRelativeToCurrentPosition)
    Next ii
    S1.Close

    'Read the words from the first file into the variant array;
    'Write them back into the second file in reverse order.

    S1.Open FILENAME1, SSFMOpenForRead
    S2.Open FILENAME2, SSFMCreateForWrite

    varStart = 0
    For ii = 0 To UBound(varT)
        S1.Read varD(ii), varP(ii) - varStart
        varStart = varP(ii)
    Next ii

    For ii = UBound(varT) To 0 Step -1
        S2.Write varD(ii)
    Next ii

    S2.Close
    S1.Close

    'After using AudioOutputStream, reset voice AudioOutput property
    Set V.AudioOutput = V.GetAudioOutputs("").item(0)




End Sub

Private Sub Command2_Click()


    S1.Open FILENAME1, SSFMOpenForRead
    S2.Open FILENAME2, SSFMOpenForRead

    'Use first male voice to announce the results
'    Set V.Voice = V.GetVoices("gender=male").item(0)

    V.Speak "This is the first sound file", SVSFlagsAsync
    V.SpeakStream S1, SVSFlagsAsync

    V.Speak "This is the second sound file", SVSFlagsAsync
    V.SpeakStream S2, SVSFlagsAsync

    Do
        DoEvents
    Loop Until V.WaitUntilDone(1)

    S1.Close
    S2.Close

End Sub


Private Sub Form_Load()

    Set V = New SpeechLib.SpVoice
    Set S1 = New SpFileStream      'Create stream1
    Set S2 = New SpFileStream       'Create stream2

End Sub

Private Sub ShowErrMsg()

    ' Declare identifiers:
    Const NL = vbNewLine
    Dim t As String

    t = "Desc: " & Err.Description & NL
    t = t & "Err #: " & Err.Number
    MsgBox t, vbExclamation, "Run-Time Error"
    End

End Sub

