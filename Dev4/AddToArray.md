&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Function AddToArray(arr As Variant, str As String)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`ReDim Preserve arr(LBound(arr) To UBound(arr) + 1)`  
&nbsp;&nbsp;&nbsp;&nbsp;`arr(UBound(arr)) = str`  
&nbsp;&nbsp;&nbsp;&nbsp;`AddToArray = arr`  
`End Function`  

