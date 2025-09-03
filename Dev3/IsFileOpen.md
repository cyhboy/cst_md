&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Function IsFileOpen(sFileName As String)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`On Error Resume Next`  
&nbsp;&nbsp;&nbsp;&nbsp;`Open sFileName For Binary Access Read Lock Read As #1`  
&nbsp;&nbsp;&nbsp;&nbsp;`Close #1`  
&nbsp;&nbsp;&nbsp;&nbsp;`IsFileOpen = IIf(Err.Number > 0, True, False)`  
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo 0`  
`End Function`  

