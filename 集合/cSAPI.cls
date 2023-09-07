VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents SpV As SpeechLib.SpVoice '创建sapi检测事件
Attribute SpV.VB_VarHelpID = -1

Private Sub Class_Initialize()
Set SpV = New SpeechLib.SpVoice
SpEnd = False
End Sub

Private Sub Class_Terminate()
Set SpV = Nothing
End Sub

Private Sub vox_EndStream(ByVal StreamNumber As Long, ByVal StreamPosition As Variant)
SpEnd = True
End Sub

Private Sub vox_StartStream(ByVal StreamNumber As Long, ByVal StreamPosition As Variant)
SpStart = True
End Sub
