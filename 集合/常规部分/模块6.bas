Attribute VB_Name = "Ä£¿é6"

Option Explicit

Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr

Declare PtrSafe Sub Sleep Lib "kernel32" _
    (ByVal dwMilliseconds As Long)

Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr


Declare PtrSafe Function SetForegroundWindow Lib "user32" _
    (ByVal hwnd As LongPtr) As Long

Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long



Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" _
    (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long



Public Const BM_CLICK = &HF5
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE

Public Sub Download(ByRef oBrowser As InternetExplorer, _
                     ByRef sFilename As String, _
                     ByRef bReplace As Boolean)

    If sFilename = "" Then
        Call Save(oBrowser)
    Else
        Call SaveAs(oBrowser, sFilename, bReplace)
    End If

End Sub

'https://stackoverflow.com/questions/26038165/automate-saveas-dialouge-for-ie9-vba
Public Sub Save(ByRef oBrowser As InternetExplorer)

    Dim AutomationObj As IUIAutomation
    Dim WindowElement As IUIAutomationElement
    Dim Button As IUIAutomationElement
    Dim hwnd As LongPtr

    Set AutomationObj = New CUIAutomation

    hwnd = oBrowser.hwnd
    hwnd = FindWindowEx(hwnd, 0, "Frame Notification Bar", vbNullString)
    If hwnd = 0 Then Exit Sub

    Set WindowElement = AutomationObj.ElementFromHandle(ByVal hwnd)
    Dim iCnd As IUIAutomationCondition
    Set iCnd = AutomationObj.CreatePropertyCondition(UIA_NamePropertyId, "Save")

    Set Button = WindowElement.FindFirst(TreeScope_Subtree, iCnd)
    Dim InvokePattern As IUIAutomationInvokePattern
    Set InvokePattern = Button.GetCurrentPattern(UIA_InvokePatternId)
    InvokePattern.Invoke

End Sub

Sub SaveAs(ByRef oBrowser As Object, _
                     sFilename As String, _
                     bReplace As Boolean)

    'https://msdn.microsoft.com/en-us/library/system.windows.automation.condition.truecondition(v=vs.110).aspx?cs-save-lang=1&cs-lang=vb#code-snippet-1
    Dim AllElements As IUIAutomationElementArray
    Dim Element As IUIAutomationElement
    Dim InvokePattern As IUIAutomationInvokePattern
    Dim iCnd As IUIAutomationCondition
    Dim AutomationObj As IUIAutomation
    Dim FrameElement As IUIAutomationElement
    Dim bFileExists As Boolean
    Dim hwnd As LongPtr

    'create the automation object
    Set AutomationObj = New CUIAutomation

    WaitSeconds 3

    'get handle from the browser
    hwnd = oBrowser.hwnd

    'get the handle to the Frame Notification Bar
    hwnd = FindWindowEx(hwnd, 0, "DUIViewWndClassName", vbNullString)
'    If hWnd = 0 Then Exit Sub
100
hwnd = FindWindowEx(hwnd, 0, "±£´æÍøÒ³", vbNullString)
GoTo 100
    'obtain the element from the handle
    Set FrameElement = AutomationObj.ElementFromHandle(ByVal hwnd)

    'Get split buttons elements
    Set iCnd = AutomationObj.CreatePropertyCondition(UIA_ControlTypePropertyId, UIA_SplitButtonControlTypeId)
    Set AllElements = FrameElement.FindAll(TreeScope_Subtree, iCnd)

    'There should be only 2 split buttons only
    If AllElements.Length = 2 Then

        'Get the second split button which when clicked shows the other three Save, Save As, Save and Open
        Set Element = AllElements.GetElement(1)

        'click the second spin button to display Save, Save as, Save and open options
        Set InvokePattern = Element.GetCurrentPattern(UIA_InvokePatternId)
        InvokePattern.Invoke

        'Tab across from default Open to Save, down twice to click Save as
        'Displays Save as window
        SendKeys "{TAB}"
        SendKeys "{DOWN}"
        SendKeys "{ENTER}"

        'Enter Data into the save as window


        Call SaveAsFilename(sFilename)

        bFileExists = SaveAsSave
        If bFileExists Then
            Call File_Already_Exists(bReplace)
        End If
    End If
End Sub

Private Sub SaveAsFilename(FileName As String)

    Dim hwnd As LongPtr
    Dim TimeOut As Date
    Dim fullfilename As String
    Dim AutomationObj As IUIAutomation
    Dim WindowElement As IUIAutomationElement


    'Find the Save As window, waiting a maximum of 10 seconds for it to appear
    TimeOut = Now + TimeValue("00:00:10")
    Do
        hwnd = FindWindow("#32770", "Save As")
        DoEvents
        Sleep 200
    Loop Until hwnd Or Now > TimeOut

    If hwnd Then

        SetForegroundWindow hwnd

        'create the automation object
        Set AutomationObj = New CUIAutomation

        'obtain the element from the handle
        Set WindowElement = AutomationObj.ElementFromHandle(ByVal hwnd)

        'Set the filename into the filename control only when one is provided, else use the default filename
        If FileName <> "" Then Call SaveAsSetFilename(FileName, AutomationObj, WindowElement)

    End If

End Sub

'Set the filename to the Save As Dialog
Private Sub SaveAsSetFilename(ByRef sFilename As String, ByRef AutomationObj As IUIAutomation, _
                                ByRef WindowElement As IUIAutomationElement)

    Dim Element As IUIAutomationElement
    Dim ElementArray As IUIAutomationElementArray
    Dim iCnd As IUIAutomationCondition

    'Set the filename control
    Set iCnd = AutomationObj.CreatePropertyCondition(UIA_AutomationIdPropertyId, "FileNameControlHost")
    Set ElementArray = WindowElement.FindAll(TreeScope_Subtree, iCnd)

    If ElementArray.Length <> 0 Then
        Set Element = ElementArray.GetElement(0)
        'should check that it is enabled

        'Update the element
        Element.SetFocus

        ' Delete existing content in the control and insert new content.
        SendKeys "^{HOME}" ' Move to start of control
        SendKeys "^+{END}" ' Select everything
        SendKeys "{DEL}" ' Delete selection
        SendKeys sFilename
    End If

End Sub



'Get the window text
Private Function Get_Window_Text(hwnd As LongPtr) As String

    'Returns the text in the specified window

    Dim Buffer As String
    Dim Length As Long
    Dim Result As Long

    SetForegroundWindow hwnd
    Length = SendMessage(hwnd, WM_GETTEXTLENGTH, 0, 0)
    Buffer = Space(Length + 1) '+1 for the null terminator
    Result = SendMessage(hwnd, WM_GETTEXT, Len(Buffer), ByVal Buffer)


    Get_Window_Text = Left(Buffer, Length)

End Function

'Click Save on the Save As Dialog
Private Function SaveAsSave() As Boolean

    'Click the Save button in the Save As dialogue, returning True if the ' already exists'
    'window appears, otherwise False

    Dim hWndButton As LongPtr
    Dim hWndSaveAs As LongPtr
    Dim hWndConfirmSaveAs As LongPtr
    Dim TimeOut As Date


    'Find the Save As window, waiting a maximum of 10 seconds for it to appear
    TimeOut = Now + TimeValue("00:00:10")
    Do
        hWndSaveAs = FindWindow("#32770", "Save As")
        DoEvents
        Sleep 200
    Loop Until hWndSaveAs Or Now > TimeOut

    If hWndSaveAs Then

        SetForegroundWindow hWndSaveAs

        'Get the child Save button
        hWndButton = FindWindowEx(hWndSaveAs, 0, "Button", "&Save")
    End If

    If hWndButton Then

        'Click the Save button


        Sleep 100
        SetForegroundWindow hWndButton
        PostMessage hWndButton, BM_CLICK, 0, 0
    End If


    'Set function return value depending on whether or not the ' already exists' popup window exists
    Sleep 500
    hWndConfirmSaveAs = FindWindow("#32770", "Confirm Save As")

    If hWndConfirmSaveAs Then
        SaveAsSave = True
    Else
        SaveAsSave = False
    End If

End Function

'Addresses the case when saving the file when it already exists.
'The file can be overwritten if Replace boolean is set to True
Private Sub File_Already_Exists(Replace As Boolean)

    'Click Yes or No in the ' already exists. Do you want to replace it?' window

    Dim hWndSaveAs As LongPtr
    Dim hWndConfirmSaveAs As LongPtr
    Dim AutomationObj As IUIAutomation
    Dim WindowElement As IUIAutomationElement
    Dim Element As IUIAutomationElement
    Dim iCnd As IUIAutomationCondition
    Dim InvokePattern As IUIAutomationInvokePattern


    hWndConfirmSaveAs = FindWindow("#32770", "Confirm Save As")

    Set AutomationObj = New CUIAutomation
    Set WindowElement = AutomationObj.ElementFromHandle(ByVal hWndConfirmSaveAs)

    If hWndConfirmSaveAs Then

        If Replace Then
            Set iCnd = AutomationObj.CreatePropertyCondition(UIA_NamePropertyId, "Yes")
        Else
            Set iCnd = AutomationObj.CreatePropertyCondition(UIA_NamePropertyId, "No")
        End If

        Set Element = WindowElement.FindFirst(TreeScope_Subtree, iCnd)
        Set InvokePattern = Element.GetCurrentPattern(UIA_InvokePatternId)
        InvokePattern.Invoke
    End If

End Sub


Public Sub WaitSeconds(intSeconds As Integer)
  On Error GoTo Errorh

  Dim datTime As Date

  datTime = DateAdd("s", intSeconds, Now)

  Do
    Sleep 100
    DoEvents
  Loop Until Now >= datTime

exitsub:
  Exit Sub

Errorh:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , "WaitSeconds"
  Resume exitsub
End Sub

