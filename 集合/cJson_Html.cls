VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cJson_html"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private oWindow As Object
Private oHtml As Object
'-------------------------------------------------------通过html来解析json数据
'------------https://www.cnblogs.com/qingmingsang/articles/5216125.html

Friend Function Json_oParse(ByVal strText As String) As Object ''https://www.w3school.com.cn/js/js_json_jsonp.asp
    Dim i As Long
    Dim strTemp As String * 1
    Dim callBack As String
    
    strText = Trim$(strText)
    i = Len(strText)
    If i = 0 Then Exit Function
    strTemp = Left$(strText, 1)
    If strTemp = "[" Or strTemp = "{" Then
        strText = "var o=" & strText
    ElseIf strTemp = "(" Then          '处理一些不规范的json数据
        i = i - InStrRev(strText, ")") '-----------------);"这种情况 '去掉前后的括号, 通常出现在callback jsonp返回的数据, 例如: callback({****});
        strText = Mid$(strTemp, 2, Len(strText) - 1 - i)
        strText = Trim$(strText)
        strText = "var o=" & strText
    Else
        callBack = Left$(strText, i - InStr(1, strText, "(", vbBinaryCompare) - 1)
        oWindow.eval "function " & callBack & "(o){json=o}", "JScript"
    End If
    oWindow.eval strText, "JScript" '这里并没有完成彻底的转移
    Set Json_oParse = oWindow '------注意必须保证,在执行完毕才能销毁对象, 否则对象的属性将会消失
End Function

Private Sub Class_Initialize()
    Set oHtml = CreateObject("htmlfile")
    Set oWindow = oHtml.parentwindow
End Sub

Private Sub Class_Terminate() '注意必须保证,在执行完毕才能销毁对象
    Set oHtml = Nothing
    Set oWindow = Nothing
End Sub
