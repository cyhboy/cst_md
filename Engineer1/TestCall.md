&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub TestCall(proc As String)`  
`'    If testing Then`  
`'        Exit Sub`  
`'    End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrorHandler`  
&nbsp;&nbsp;&nbsp;&nbsp;`Application.Run "'cst.xlam'!" & proc`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`ErrorHandler:`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Err.Clear`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'Application.Run "'cst.xlam'!AIA." & proc`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 30`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`End Sub`  


# BeCaller
- TestCall{S}(11)->[[MyMsgBox]]{S}

