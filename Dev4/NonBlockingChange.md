&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub NonBlockingChange()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Update`  
&nbsp;&nbsp;&nbsp;&nbsp;`Application.OnTime Now + `[`TimeValue`](TimeValue)`("0:00:01"), "NonBlockingChange"`  
`End Sub`  


# BeCaller
- NonBlockingChange{S}(6)->[[TimeValue]]{F}

