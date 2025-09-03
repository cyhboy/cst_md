&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Private Function json_ParseErrorMessage(json_String As String, ByRef json_Index As Long, ErrorMessage As String)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Provide detailed parse error message, including details of where and what occurred`  
&nbsp;&nbsp;&nbsp;&nbsp;`'`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Example:`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Error parsing JSON:`  
&nbsp;&nbsp;&nbsp;&nbsp;`' {"abcde":True}`  
&nbsp;&nbsp;&nbsp;&nbsp;`'          ^`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim json_StartIndex As Long`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim json_StopIndex As Long`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Include 10 characters before and after error (if possible)`  
&nbsp;&nbsp;&nbsp;&nbsp;`json_StartIndex = json_Index - 10`  
&nbsp;&nbsp;&nbsp;&nbsp;`json_StopIndex = json_Index + 10`  
&nbsp;&nbsp;&nbsp;&nbsp;`If json_StartIndex <= 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`json_StartIndex = 1`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`If json_StopIndex > VBA.Len(json_String) Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`json_StopIndex = VBA.Len(json_String)`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`json_ParseErrorMessage = "Error parsing JSON:" & VBA.vbNewLine & _`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`VBA.Mid$(json_String, json_StartIndex, json_StopIndex - json_StartIndex + 1) & VBA.vbNewLine & _`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`VBA.Space$(json_Index - json_StartIndex) & "^" & VBA.vbNewLine & _`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ErrorMessage`  
`End Function`  

