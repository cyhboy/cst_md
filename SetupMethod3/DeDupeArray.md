&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Function DeDupeArray(vArray As Variant) As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim oDict As Object`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim i As Long`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set oDict = CreateObject("Scripting.Dictionary")`  
&nbsp;&nbsp;&nbsp;&nbsp;`For i = LBound(vArray) To UBound(vArray)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`oDict(vArray(i)) = True`  
&nbsp;&nbsp;&nbsp;&nbsp;`Next`  
&nbsp;&nbsp;&nbsp;&nbsp;`DeDupeArray = oDict.Keys()`  
`End Function`  

