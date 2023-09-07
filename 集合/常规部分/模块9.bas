Attribute VB_Name = "Ä£¿é9"
Option Explicit

'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'
'Sub SwapPtr(ByRef sA As String, sB As String)
'  Dim lTmp As Long
'  CopyMemory lTmp, ByVal VarPtr(sA), 4
'  CopyMemory ByVal VarPtr(sA), ByVal VarPtr(sB), 4
'  CopyMemory ByVal VarPtr(sB), lTmp, 4
'End Sub
'
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'
'Sub SwapStrPtr3(sA As String, sB As String)
'  Dim lTmp As Long
'  Dim pTmp As Long, psA As Long, psB As Long
'  pTmp = StrPtr(sA): psA = VarPtr(sA): psB = VarPtr(sB)
'  CopyMemory ByVal psA, ByVal psB, 4
'  CopyMemory ByVal psB, pTmp, 4
'End Sub

#If VB7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
#Else
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
#End If


Sub SwapString(ByRef aString As String, ByRef bString As String)
  Dim lTmp As Long
  Dim sPtr As Long, aPtr As Long, bPtr As Long
  
  sPtr = StrPtr(aString)
  aPtr = VarPtr(aString): bPtr = VarPtr(bString)
  CopyMemory ByVal aPtr, ByVal bPtr, 4
  CopyMemory ByVal bPtr, sPtr, 4
End Sub

Sub kfdkf()
Dim a As String
Dim b As String
a = "CopyMemory ByVal VarPtr(sB), lTmp, 4"
b = "def"
SwapString a, b
Debug.Print a
Debug.Print b
End Sub
