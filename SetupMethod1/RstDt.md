&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub RstDt()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Range("L2") <> "" And Range("L3") <> "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Range("L2").Select`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Range(Selection, Selection.End(xlDown)).Select`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Selection.NumberFormat = "yyyy-mm-dd hh:mm:ss"`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`End Sub`  

