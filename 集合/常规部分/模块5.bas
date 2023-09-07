Attribute VB_Name = "模块5"
'
'Function GetCookie(strUrl)
'    With CreateObject("WinHttp.WinHttpRequest.5.1")
'        .Open "GET", strUrl, False
'        .setRequestHeader "REFERER", strUrl
'        .setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)"
'        .setRequestHeader "Accept", "text/xml,application/xml,application/xhtml+xml,text/html;q=0.9,text/plain;q=0.8,image/png,*/*;q=0.5"
'        .setRequestHeader "Accept-Language", "en-us,en;q=0.5"
'        .setRequestHeader "Accept-Charset", "ISO-8859-1,utf-8;q=0.7,*;q=0.7"
'        .send
'        strCookie = .getResponseHeader("Cookie")
'        strCookie = Split(strCookie, ";")(0)
'        GetCookie = strCookie
'    End With
'End Function
'
'Sub Demo()
'Debug.Print GetCookie("http://data.10jqka.com.cn/funds/ggzjl/")
'End Sub

Sub SeachBook() 'http://www.yuedu88.com/ 对应此网站

    Dim Urlx As String
    Dim fl As Object, FilePath As String
    Dim bookt As Object
    Dim strHtml As String, textl As Object, HtmF As Object
    Dim XmlH As Object
    Dim rt As Long
'    Set HtmF = CreateObject("htmlfile")             'https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms753804(v=vs.85)
'    Set XmlH = CreateObject("Msxml2.ServerXMLHTTP") '("WinHttp.WinHttpRequest.5.1") 'https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms766431(v=vs.85)
'    strx = "三体"
'    strx = ThisWorkbook.Application.EncodeURL(strx)
'    urlx = "http://www.yuedu88.com/search.php?q=" & strx
'    urlx = "https://www.jisilu.cn/data/etf/#index"
'    urlx = "https://www.jisilu.cn/data/etf/etf_list/?___jsl=LST___t=" & TimeStamp & "&rp=25&page=1"
'    urlx = "http://wgeo.weather.com.cn/ip/?_=" & timestamp '"https://www.xicidaili.com/wn/"
'    Debug.Print urlx
'    HtmF.DesignMode = "on" ' 开启编辑模式
'    Dim strx As String
'    strx = ""
'    Urlx = "http://data.10jqka.com.cn/funds/ggzjl/field/zdf/order/desc/page/10/ajax/1/free/1/"
'    With XmlH
'        .Open "GET", Urlx, False
'        .send
''        a = .getAllResponseHeaders
'        strHtml = .responseText
''        a = .getAllResponseHeaders
''        Debug.Print a
''        a = .getResponseHeader("Cookie")
''         Debug.Print strHtml
''        JSDemo_1 strHtml
'    End With
'    HtmF.Write strHtml ' 写入数据
Set IE = CreateObject("InternetExplorer.Application") '创建ie浏览器对象
    IE.Visible = False
    IE.Navigate "http://data.10jqka.com.cn/funds/ggzjl/"
    For i = 1 To 10
'    Urlx = "http://data.10jqka.com.cn/funds/ggzjl/field/zdf/order/desc/page/" & CStr(i) & "/ajax/1/free/1/"
'    ie.IEIsInPrivateBrowsing = True
'ie.Navigate "http://data.10jqka.com.cn/"
'ie.Refresh2 3
'    ie.Navigate Urlx
'    Debug.Print ie.Document.Cookie
    
    Next
    GoTo 100
100
    IE.Quit
    
    
''Debug.Print strHtml
'    Set bookt = HtmF.getElementsByTagName("table")
'    Set bookt = HtmF.getElementById("table")
'
'
'    For Each textl In bookt
'
'
'    Next
'    For Each Textl In BookT.Children 'ChildNodes

'    Set x = Textl.getElementsByClassName("calibre4")
'    If Textl.NodeType = 3 Then Debug.Print Textl.Data
'    If i < 573 Then Debug.Print Textl.innertext: i = i + 1
''    Next
'100
'Set bookt = HtmF.Cookie
''''    Set BookT = HtmF.getElementById("pagebar") '根据id提取数据
''
'''     Set BookT = htmlf.getAttribute("pagebar")
''
'    Set bookt = HtmF.getelementsbyclassname("m_table") 'getElementsByClassName '根据span class来访问
'
'
'
'
'

'    For Each item In bookt.item(0).Children.item(1).Children
''    Set oA = item.getelementsbyclassname("title")  'getElementsByClassName
''    Set x = oA(0).getElementsByTagName("span")
''    Set a = oA(0).getElementsByTagName("a")
''    Set b = item.getelementsbyclassname("rating-info")
''    Set c = b(0).getElementsByTagName("span") '结果对应的超链接
''    Set w = item.getelementsbyclassname("pic")
''    Set o = w(0).getElementsByTagName("a")
''    Set p = w(0).getElementsByTagName("img")
'    For Each itemx In item.Children
'    Debug.Print itemx.innertext
'    Next
''    GoTo 100
'    Next
'
''        For Each item In HtmF.all
''        Debug.Print item.tagName, item.innertext, item.innerhtml
'        Debug.Print item.Children.item(0).Children.item("description").Content
''        Set x = item.getElementsByTagName("a")
''            If item.tagName = "pagebar" Then
'100
''       Debug.Print item.tagName
''
''            End If
''        Next
''        GoTo 100
''
''        Next
'
'Set BookT = HtmF.getElementsByClassName("intro")
'
'
''GoTo 100
'For Each Textl In BookT
'Set BookT = HtmF.getElementsByClassName("intro")
'Set x = Textl.getElementsByTagName("img")
'Author = x(0).href
'200
'
''GoTo 200
'Next
'   For Each Textl In BookT.Children '每个节点(item) '第二个节点是搜索结果所在
'   GoTo 100
'For Each Textl In BookT
'GoTo 100
'Debug.Print BookT.item(0).Children.Length
'GoTo 100
'k = Textl.Children.Length '此方法可获取内包含的元素的多少
''If k > 0 Then
''k = k - 1
'    Set x = Textl.getElementsByTagName("a")
'    j = x.Length
'    If j > 0 Then
'    j = j - 1
'    For l = 0 To j
'    Debug.Print x(l).href, x(l).Text
'    Next
'    End If
''End If
'    Next
End Sub

Sub JSDemo_1(ByVal strx As String)
Dim Urlx As String
Dim fso As New FileSystemObject
Dim objJason As Object
Dim objSc As Object
Dim objItem As Object
Dim iRow As Integer


'Range("A:D").Clear
'[A1:D1] = Array("c", "p.x", "p.y", "p.z")
Set objSc = CreateObject("ScriptControl")
objSc.Language = "JScript"
'Dim fl As TextStream
'Set fl = fso.OpenTextFile("C:\Users\Lian\Desktop\new  31.txt", ForReading, False, TristateUseDefault)
'strx = fl.ReadAll
'strx = Replace(strx, Chr(34) & "c" & Chr(34), Chr(34) & "cc" & Chr(34))
100
strx = Replace(strx, "page", "ipage")
strx = Replace(strx, "rows", "irows")
strx = Replace(strx, "cell", "icell")
Set objJason = objSc.eval("s=" & strx)
'iRow = 2
'For Each objItem In objJason.Rows.item
'Debug.Print objJason.irows.item(0).id
'objJason.Properties
'For Each x In objJason.Properties
'Debug.Print x.Name
'Next
i = 0
For Each objItem In objJason.irows
'GoTo 100
With objItem
'Debug.Print objItem.Count
 Set objt = .icell
 Debug.Print objt.urls
i = i + 1
'Set objX = objJason.irows
'Sample .icell
''Exit For
'For Each x In objt.Properties
'x.Name
'Next
'Cells(iRow, "A") = .cc
'Cells(iRow, "B") = .p.x
'Cells(iRow, "C") = .p.Y
'Cells(iRow, "D") = .p.z
'Debug.Print .icell.fee
'
''For Each x In objt
'Debug.Print objJason.Items.Count
'
''
''iRow = iRow + 1
'Debug.Print CallByName(objItem.icell, "amount", VbGet)
End With
Next objItem
Set objSc = Nothing
Set objJason = Nothing
Set objItem = Nothing
End Sub

Private Function ReplaceTest(ByVal rpText As String) As String

'  Dim objinfo As InterfaceInfo
  Dim regEx, str1               ' 建立变量。
  str1 = Cells(1, "f").Value
  Set regEx = CreateObject("VBScript.RegExp")             ' 建立正则表达式。
  regEx.Pattern = "cell"              ' 设置模式。
  regEx.IgnoreCase = True               ' 设置是否区分大小写。
  ReplaceTest = regEx.Replace(str1, replStr)         ' 作替换。
End Function

Public Sub Sample(ByVal obj As Object)
'TypeLib Information(tlbinf32.dll)要参照
  Dim ta As TLI.TLIApplication
  Dim ci As TLI.ConstantInfo
  Dim mi As TLI.MemberInfo
   
  Set ta = New TLI.TLIApplication
100
Set a = ta.InterfaceInfoFromObject(obj)
'
k = a.Members.Count
i = 1
For i = 1 To k

Debug.Print a.Members(i).Name

Next

End

'
'  With ta.InterfaceInfoFromObject(obj)
'    For Each x In .Members
''    ci.Name
'Dim a As InterfaceInfo
''Debug.Print x.value
'
'GoTo 100
''Debug.Print ci.Name
''      For Each mi In ci.AttributeMask
''        Debug.Print ci.Name, mi.Name, mi.value
''      Next
'    Next
'  End With
End Sub





Private Sub Form_Load()
Dim oTil As Object
Set oTil = CreateObject("TLI.TLIApplication")
 Dim oTLB As Object, i As Long

 Set oTLB = oTil.InterfaceInfoFromObject(oTil)

 Debug.Print oTLB.Name

 For i = 1 To oTLB.Members.Count
' Select Case oTLB.Members(I).InvokeKind
 strx = oTLB.Members(i).InvokeKind
 Select Case strx
 Case INVOKE_CONST
 Debug.Print " 常数:" & oTLB.Members(i).Name
 Case INVOKE_EVENTFUNC
 Debug.Print " 事件:" & oTLB.Members(i).Name
 Case INVOKE_FUNC
 Debug.Print " 方法:" & oTLB.Members(i).Name
 Case INVOKE_PROPERTYGET
 Debug.Print "属性(Get):" & oTLB.Members(i).Name
 Case INVOKE_PROPERTYPUT
 Debug.Print "属性(Let):" & oTLB.Members(i).Name
 Case INVOKE_PROPERTYPUTREF
 Debug.Print "属性(Set):" & oTLB.Members(i).Name
 Case INVOKE_UNKNOWN
 Debug.Print " 未知:" & oTLB.Members(i).Name
 End Select
 Next
End Sub






