&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Function RangeToArray()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim SourceRange As Range`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set SourceRange = Selection.CurrentRegion`  
&nbsp;&nbsp;&nbsp;&nbsp;`RangeToArray = SourceRange.Value`  
`End Function`  

