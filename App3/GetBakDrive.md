&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Function GetBakDrive() As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim bakDrive As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fso As Object`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set fso = CreateObject("Scripting.FileSystemObject")`  
&nbsp;&nbsp;&nbsp;&nbsp;`bakDrive = fso.GetDriveName(ActiveWorkbook.path) & "\BAK"`  
&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(bakDrive, ":") = 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`bakDrive = "C:" & "\BAK"`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(bakDrive, "D:") > 0 Or InStr(bakDrive, "d:") > 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`bakDrive = "C:" & "\BAK"`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set fso = Nothing`  
&nbsp;&nbsp;&nbsp;&nbsp;`GetBakDrive = bakDrive`  
`End Function`  

