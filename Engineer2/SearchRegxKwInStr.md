&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Function SearchRegxKwInStr(str As String, regxKw As String, Optional multiLine As Boolean = False, Optional ignoreC As Boolean = False)`  
&nbsp;&nbsp;&nbsp;&nbsp;`' SearchRegxKwInStr`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim reg As New RegExp`  
&nbsp;&nbsp;&nbsp;&nbsp;`With reg`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.Global = True`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.IgnoreCase = ignoreC`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.multiLine = multiLine`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.Pattern = regxKw`  
&nbsp;&nbsp;&nbsp;&nbsp;`End With`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim mc As MatchCollection`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim dynamicStr1 As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`dynamicStr1 = ""`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set mc = reg.Execute(str)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If mc.count > 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`dynamicStr1 = mc.Item(0).SubMatches.Item(0)`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`SearchRegxKwInStr = dynamicStr1`  
`End Function`  

