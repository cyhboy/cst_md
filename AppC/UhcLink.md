&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub UhcLink()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`'On Error GoTo ErrorHandler`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim n As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`n = Selection.count`  
&nbsp;&nbsp;&nbsp;&nbsp;`If n > 1 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`n = Selection.SpecialCells(xlCellTypeVisible).count`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`If n > 1 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Dim curCell As Range`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`For Each curCell In Selection`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If curCell.EntireColumn.Hidden = False And curCell.EntireRow.Hidden = False Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`curCell.Select`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox subName`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`RobotRun`](RobotRun)` "UhcLink"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Next curCell`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim linkStr As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`linkStr = Cells(currentRow, 10)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim link As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`link = `[`CutStringByStartAndEnd`](CutStringByStartAndEnd)`(linkStr, """", """")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`'On Error Resume Next`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim iee As Object`  
&nbsp;&nbsp;&nbsp;&nbsp;`While iee Is Nothing`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Set iee = CreateObject("InternetExplorer.Application")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Sleep 2000`  
&nbsp;&nbsp;&nbsp;&nbsp;`Wend`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`iee.Visible = False`  
&nbsp;&nbsp;&nbsp;&nbsp;`'iee.navigate "about:blank"`  
&nbsp;&nbsp;&nbsp;&nbsp;`iee.navigate link`  
&nbsp;&nbsp;&nbsp;&nbsp;`While iee.Busy`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`DoEvents`  
&nbsp;&nbsp;&nbsp;&nbsp;`Wend`  
&nbsp;&nbsp;&nbsp;&nbsp;`Application.Wait Now + TimeSerial(0, 0, 5)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim elm As Object`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set elm = iee.Document.getElementById("post-date")`  
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 12) = elm.innerText`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set elm = iee.Document.getElementById("post_view_count")`  
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 7) = elm.innerText`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set elm = iee.Document.getElementById("post_comment_count")`  
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 11) = elm.innerText`  
&nbsp;&nbsp;&nbsp;&nbsp;`iee.Quit`  
`'ErrorHandler:`  
`'    If Err.Number <> 0 Then`  
`'        MyMsgBox Err.Number & " " & Err.Description, 30`  
`'    End If`  
`End Sub`  


# BeCaller
- UhcLink{S}(13)->[[RobotRun]]{S}
- UhcLink{S}(23)->[[CutStringByStartAndEnd]]{F}

