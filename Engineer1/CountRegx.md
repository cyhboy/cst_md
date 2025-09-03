&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Function CountRegx(text As String, patt As String) As Long`  
`'    If testing Then`  
`'        Exit Function`  
`'    End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrorHandler`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim RE As New RegExp`  
&nbsp;&nbsp;&nbsp;&nbsp;`RE.Pattern = patt`  
&nbsp;&nbsp;&nbsp;&nbsp;`RE.Global = True`  
&nbsp;&nbsp;&nbsp;&nbsp;`RE.IgnoreCase = False`  
&nbsp;&nbsp;&nbsp;&nbsp;`RE.multiLine = True`  
&nbsp;&nbsp;&nbsp;&nbsp;`'Retrieve all matches`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim Matches As MatchCollection`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set Matches = RE.Execute(text)`  
&nbsp;&nbsp;&nbsp;&nbsp;`'Return the corrected count of matches`  
&nbsp;&nbsp;&nbsp;&nbsp;`CountRegx = Matches.count`  
`ErrorHandler:`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 30`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`End Function`  


# BeCaller
- CountRegx]]{F}(13)->[[MyMsgBox]]{S}

