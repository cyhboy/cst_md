&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub AddToArrayG(str As String)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If `[`TypeName`](TypeName)`(garr) <> "String()" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`garr = `[`EmptyStringArray`](EmptyStringArray)`()`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`ReDim Preserve garr(LBound(garr) To UBound(garr) + 1)`  
&nbsp;&nbsp;&nbsp;&nbsp;`garr(UBound(garr)) = str`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`End Sub`  


# BeCaller
- AddToArrayG{S}(5)->[[TypeName]]{F}
- AddToArrayG{S}(6)->[[EmptyStringArray]]{F}

