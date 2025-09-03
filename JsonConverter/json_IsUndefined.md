&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Private Function json_IsUndefined(ByVal json_Value As Variant) As Boolean`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Empty / Nothing -> undefined`  
&nbsp;&nbsp;&nbsp;&nbsp;`Select Case VBA.VarType(json_Value)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Case VBA.vbEmpty`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`json_IsUndefined = True`  
&nbsp;&nbsp;&nbsp;&nbsp;`Case VBA.vbObject`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Select Case VBA.TypeName(json_Value)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Case "Empty", "Nothing"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`json_IsUndefined = True`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End Select`  
&nbsp;&nbsp;&nbsp;&nbsp;`End Select`  
`End Function`  

