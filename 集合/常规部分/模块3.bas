Attribute VB_Name = "Ä£¿é3"
'Option Explicit

'Private WithEvents V As SpeechLib.SpVoice

Private Sub Command1_Click()
Set V = New SpVoice
strx = "1. 2. 3. 4. 5. 6. 7. 8. 9. 10. 11. 12. 13. 14. 15. 16. 17. 18. 19. 20. 23. 24. 25. 26"
    V.Speak strx, SVSFlagsAsync
100
    V.Skip "Sentence", 5
'    GoTo 100
End Sub

Private Sub Command2_Click()
    
End Sub

Private Sub Form_Load()

    

    HScroll1.Min = -5
    HScroll1.Max = 5
    HScroll1.Value = -2

    

End Sub

Private Sub HScroll1_Change()
    If HScroll1.Value > 0 Then
        Command2.Caption = " Skip forward " & HScroll1.Value & " sentences"
    Else
        Command2.Caption = " Skip backward " & Abs(HScroll1.Value) & " sentences"
    End If
End Sub

Private Sub V_Word(ByVal StreamNumber As Long, ByVal StreamPosition As Variant, _
                   ByVal CharacterPosition As Long, ByVal Length As Long)
    Text1.SetFocus
    Text1.SelStart = CharacterPosition
    Text1.SelLength = Length
End Sub

