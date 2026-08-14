&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Private Function NormalizePath(ByVal p As String) As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`p = Trim(p)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Len(p) = 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`NormalizePath = p`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Right$(p, 1) <> "\" Then p = p & "\"`  
&nbsp;&nbsp;&nbsp;&nbsp;`NormalizePath = p`  
`End Function`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  

