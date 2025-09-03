&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub WriteTxt2Code(text As String, path As String)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' On Error GoTo ErrorHandler`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim ff As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`ff = FreeFile()`  
&nbsp;&nbsp;&nbsp;&nbsp;`Open path For Output As #ff`  
&nbsp;&nbsp;&nbsp;&nbsp;`Print #ff, text`  
&nbsp;&nbsp;&nbsp;&nbsp;`Close #ff`  
`ErrorHandler:`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 30`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`End Sub`  


# BeCaller
- WriteTxt2Code{S}(12)->[[MyMsgBox]]{S}

