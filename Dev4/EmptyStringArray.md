&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Function EmptyStringArray() As String()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`EmptyStringArray = VBA.Strings.Split(vbNullString)`  
`End Function`  

