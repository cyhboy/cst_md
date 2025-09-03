&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Private Function json_Peek(json_String As String, ByVal json_Index As Long, Optional json_NumberOfCharacters As Long = 1) As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`' "Peek" at the next number of characters without incrementing json_Index (ByVal instead of ByRef)`  
&nbsp;&nbsp;&nbsp;&nbsp;`json_SkipSpaces json_String, json_Index`  
&nbsp;&nbsp;&nbsp;&nbsp;`json_Peek = VBA.Mid$(json_String, json_Index, json_NumberOfCharacters)`  
`End Function`  

