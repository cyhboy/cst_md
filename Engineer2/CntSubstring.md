&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Function CntSubstring(text As String, subStr As String, Optional ignoreFlag As Boolean = False) As Long`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If ignoreFlag Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`CntSubstring = (Len(UCase(text)) - Len(Replace(UCase(text), UCase(subStr), ""))) / Len(UCase(subStr))`  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`CntSubstring = (Len(text) - Len(Replace(text, subStr, ""))) / Len(subStr)`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`End Function`  

