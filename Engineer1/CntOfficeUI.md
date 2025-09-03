&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub CntOfficeUI()`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' On Error GoTo ErrorHandler`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim procName As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`procName = "N/A"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim iCol As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`iCol = 3`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim oXML As Object`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set oXML = New DOMDocument`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Set oXML = New DOMDocument60`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim strUrl As String, resultStr As String`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim strFilePath As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`strFilePath = "C:\Users\" & Environ$("username") & "\AppData\Local\Microsoft\Office\Excel.officeUI"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox strFilePath`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`oXML.Load (strFilePath)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim oXmlNodes As IXMLDOMNodeList`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Set oXmlNodes = oXML.SelectNodes("//customUI/ribbon/tabs")`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set oXmlNodes = oXML.SelectNodes("//mso:customUI/mso:ribbon/mso:tabs")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox oXmlNodes.Length`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim node As IXMLDOMNode`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim xx As String`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`buttonNum = 0`  
&nbsp;&nbsp;&nbsp;&nbsp;`groupNum = 0`  
&nbsp;&nbsp;&nbsp;&nbsp;`tabNum = 0`  
&nbsp;&nbsp;&nbsp;&nbsp;`menuNum = 0`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`For Each node In oXmlNodes`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`xx = `[`ListNodes`](ListNodes)`(node, procName, iCol, False)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If xx = "exit" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox "exit"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Next node`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Set oXML = Nothing`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox tabNum`  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox groupNum`  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox buttonNum`  
&nbsp;&nbsp;&nbsp;&nbsp;`' ErrorHandler:`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 30`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`End Sub`  


# BeCaller
- CntOfficeUI{S}(21)->[[ListNodes]]{F}
- CntOfficeUI{S}(28)->[[MyMsgBox]]{S}

