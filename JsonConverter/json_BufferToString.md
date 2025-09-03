&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Private Function json_BufferToString(ByRef json_Buffer As String, ByVal json_BufferPosition As Long) As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`If json_BufferPosition > 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`json_BufferToString = VBA.Left$(json_Buffer, json_BufferPosition)`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`End Function`  

