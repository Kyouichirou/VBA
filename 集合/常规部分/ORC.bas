Attribute VB_Name = "ORC"
Option Explicit
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Type imageBase64
    base64Text As String
    imageWidth As Long
    imageHeight As Long
End Type

Function GetTextFromSinglePicture(inPicPath As String, Optional cT As String) As String  'ͼƬ����Ϣ���룬�������xml�ı���
    Dim xmlDoc As New MSXML2.DOMDocument60
    Dim xmlNode As MSXML2.IXMLDOMNode
    Dim xmlEle As MSXML2.IXMLDOMElement
    Dim picBase64 As imageBase64
    Dim onenoteApp As Object
    Dim sectionID As String
    Dim pageID As String
    Dim pageXmlText As String
    Dim iCNT As Integer
    Dim onenoteFullName As String
    
    '������ʱ�ıʼǱ�
    With New Scripting.FileSystemObject
        onenoteFullName = .GetSpecialFolder(TemporaryFolder) & "\" & .GetBaseName(.GetTempName) & ".one"
        '�жϺ���ֵ�Ƿ�����
        If .fileexists(inPicPath) = False Then
            GetTextFromSinglePicture = "! Error File Path !"
            Exit Function
        End If
    End With
    Set onenoteApp = CreateObject("OneNote.Application")
    If onenoteApp Is Nothing Then
        GetTextFromSinglePicture = "! Error in Openning OneNote !"
        GoTo clear_variable_before_exit
    End If
    Set xmlEle = CreateNotePageContentElement(2, inPicPath) '������ʱ�ıʼǱ�,��ȡsectionID
    Set xmlEle = AddNodeInfo(xmlEle)
    With onenoteApp
        .OpenHierarchy onenoteFullName, "", sectionID, cftSection
        .CreateNewPage sectionID, pageID, npsBlankPageNoTitle '�����µ�ҳ�棬��ȡpageID
        .GetPageContent pageID, pageXmlText, , xs2013 '��ȡҳ���XML��ʽ
    End With
    If xmlDoc.LoadXML(pageXmlText) = False Then '���뵽XML�н��д���,��ͼƬ��ʽ���뵽XML��
        GetTextFromSinglePicture = "! Error in Loading Xml !"
        GoTo clear_variable_before_exit
    End If
    With xmlDoc.getElementsByTagName("one:Page").item(0)
        .appendChild xmlEle
    End With
    onenoteApp.UpdatePageContent xmlDoc.DocumentElement.XML, , xs2013 '����Page����
    Sleep 1000 'OneNoteʶ��ͼƬ��Ҫʱ�䣬���¿�ʼ��ѯ�����1��*10�� '���ﻹ��Ҫ����
    iCNT = 10
    
re_getPageContent:
    onenoteApp.GetPageContent pageID, pageXmlText, , xs2013
    xmlDoc.LoadXML pageXmlText
    Set xmlEle = xmlDoc.DocumentElement.getElementsByTagName("one:OCRText").item(0)
    If xmlEle Is Nothing Then
        If iCNT > 0 Then
            Sleep 1000
            iCNT = iCNT - 1
            GoTo re_getPageContent
        Else
            GetTextFromSinglePicture = "! Waiting OneNote Time Expired !"
        End If
    Else
        GetTextFromSinglePicture = xmlEle.Text
        cT = xmlEle.Text
    End If
    
clear_variable_before_exit:
    If Not onenoteApp Is Nothing Then
        If Len(pageID) > 0 Then
            onenoteApp.DeleteHierarchy pageID, , True
        End If
        Set onenoteApp = Nothing
    End If
    Kill onenoteFullName
End Function

Function CreateNotePageContentElement(ContentType As Integer, paraContent As String) As MSXML2.IXMLDOMElement
    Dim xmlEle As MSXML2.IXMLDOMElement
    Dim xmlNode As MSXML2.IXMLDOMElement
                
    Dim ns As String
    ns = "one:"
    With New MSXML2.DOMDocument60
        Select Case ContentType
            Case 1 '�ı�
                Set xmlNode = .createElement(ns & "T")
                xmlNode.Text = paraContent
            Case 2 'ͼƬ
                Dim picBase64 As imageBase64
                picBase64 = getBase64(paraContent)
    
                '����һ��ͼƬXML��Ϣ
                Set xmlNode = .createElement(ns & "Image")
                xmlNode.setAttribute "format", "jpg"
                xmlNode.setAttribute "originalPageNumber", 0
                
                Set xmlEle = .createElement(ns & "Position")
                xmlEle.setAttribute "x", 0
                xmlEle.setAttribute "y", 0
                xmlEle.setAttribute "z", 0
                xmlNode.appendChild xmlEle
                
                Set xmlEle = .createElement(ns & "Size")
                xmlEle.setAttribute "width", picBase64.imageWidth
                xmlEle.setAttribute "height", picBase64.imageHeight
                xmlNode.appendChild xmlEle
                
                Set xmlEle = .createElement(ns & "Data")
                xmlEle.Text = picBase64.base64Text
                xmlNode.appendChild xmlEle
        End Select
    End With
    Set CreateNotePageContentElement = xmlNode
End Function

Function AddNodeInfo(ContentElement As MSXML2.IXMLDOMElement) As MSXML2.IXMLDOMElement
    Dim xmlEle As MSXML2.IXMLDOMElement
    Dim xmlNode As MSXML2.IXMLDOMElement
    Dim ns As String
    ns = "one:"
    Set xmlNode = ContentElement
    With New MSXML2.DOMDocument60
        Set xmlEle = .createElement(ns & "OE")
        xmlEle.appendChild xmlNode
        Set xmlNode = xmlEle
        
        Set xmlEle = .createElement(ns & "OEChildren")
        xmlEle.appendChild xmlNode
        Set xmlNode = xmlEle
        
        Set xmlEle = .createElement(ns & "Outline")
        xmlEle.appendChild xmlNode
        Set xmlNode = xmlEle
    End With

    Set AddNodeInfo = xmlNode

End Function

Function getBase64(inBmpFile As String) As imageBase64
    Dim xmlEle As MSXML2.IXMLDOMElement
    
    With New MSXML2.DOMDocument60
        Set xmlEle = .createElement("Base64Data") 'https://blog.csdn.net/webxiaoma/article/details/70053444
    End With
    xmlEle.DataType = "bin.base64"
    With New ADODB.Stream
        .type = adTypeBinary
        .Open
        .LoadFromFile inBmpFile
        xmlEle.nodeTypedValue = .Read()
        .Close
    End With
    getBase64.base64Text = xmlEle.Text
    With CreateObject("WIA.ImageFile")
        .LoadFile inBmpFile
        getBase64.imageHeight = .Height
        getBase64.imageWidth = .Width
    End With
End Function

'https://abbyy.technology/en:kb:code-sample:scripting_languages
'����ʹ��abbyy��ʵ��orc
Sub OCR_Pictures_To_Text(ByVal outputfile As String) 'תͼƬ-txt
    Dim vFNi As Variant
    Dim sFNi As Variant
    Dim sFNo As String
    Dim oTS As TextStream
    Dim t As String
    Dim sTmp As String
    
    vFNi = Application.GetOpenFilename("*.jpg,*.jpeg", , , , True)
    If VarType(vFNi) = vbBoolean Then Exit Sub
    sFNo = outputfile
    Open sFNo For Binary As #1 '������Ըĳ�ado
    Close #1
    If sFNo = "False" Then Exit Sub
    With New Scripting.FileSystemObject
        Set oTS = .CreateTextFile(sFNo)
    End With
    For Each sFNi In vFNi
        sTmp = GetTextFromSinglePicture(CStr(sFNi), t)
        oTS.Write sTmp
    Next
    oTS.Close
End Sub
