&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Function SeedFileGet(filePath As String) As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim seedFilePath As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fileName As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim ext As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(filePath, ".") > 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ext = Right(filePath, Len(filePath) - InStrRev(filePath, ".") + 1)`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(filePath, "_") > 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`fileName = Left(filePath, InStrRev(filePath, "_") - 1)`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`SeedFileGet = fileName & ext`  
`End Function`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  

