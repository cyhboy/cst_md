&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Function GetResult()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`GetResult = Trim(ReadLineByFile("C:\BAK\interaction.log"))`  
`End Function`  

