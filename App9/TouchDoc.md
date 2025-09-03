&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub TouchDoc()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`'On Error GoTo ErrorHandler`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim filePath As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`filePath = Cells(currentRow, 9) & Cells(currentRow, 11)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim wa As New Word.Application`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim wd As Word.Document`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim objSelection As Word.Selection`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`wa.Visible = False`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Set wd = wa.Documents.Open(filePath)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'    Dim objComment As Word.comment`  
`'    For Each objComment In wd.Comments`  
`'        MsgBox objComment.Author`  
`'        objComment.Author = ""`  
`'        objComment.Initial = ""`  
`'    Next`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Set objSelection = wa.Selection`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`objSelection.Font.Bold = True`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`objSelection.Font.Size = "22"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`objSelection.TypeText ("I am new here" & vbCrLf)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`wd.Save`  
&nbsp;&nbsp;&nbsp;&nbsp;`'Sleep 3000`  
&nbsp;&nbsp;&nbsp;&nbsp;`wd.Close`  
&nbsp;&nbsp;&nbsp;&nbsp;`'wd.Close savechanges:=True`  
&nbsp;&nbsp;&nbsp;&nbsp;`wa.Quit`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set wa = Nothing`  
`ErrorHandler:`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 10`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`End Sub`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  


# BeCaller
- TouchDoc{S}(24)->[[MyMsgBox]]{S}

