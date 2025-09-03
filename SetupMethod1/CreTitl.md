&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub CreTitl()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim titlAry As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`titlAry = Array("Hostname", "FQDN", "User", "Password", "Folder", "IP", "Port", "Memo", "Local Folder", "Command", "Specify File", "Last Update", "Demand", "CO", "Sequence", "Executor", "Status", "#", "#", "#", "#", "#", "#")`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim i As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`For i = 0 To UBound(titlAry)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If Trim(Cells(1, i + 1)) = "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(1, i + 1) = titlAry(i)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Next i`  
`End Sub`  

