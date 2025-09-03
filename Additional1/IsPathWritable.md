&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Function IsPathWritable(ByVal fPath As String) As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fName As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`'Dim localFName As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim ff As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim Counter As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`Counter = 1`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If (Right(fPath, 1) <> "\") Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`fPath = fPath & "\"`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`IsPathWritable = "Invalid"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrHandler`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Do`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`fName = fPath & "TempFile" & Counter & ".tmp"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'localFName = "C:\Temp\" & "TempFile" & Counter & ".tmp"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Counter = Counter + 1`  
&nbsp;&nbsp;&nbsp;&nbsp;`Loop Until Dir(fName) = ""`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo CantWrite`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`ff = FreeFile()`  
&nbsp;&nbsp;&nbsp;&nbsp;`Open fName For Output Access Write As #ff`  
&nbsp;&nbsp;&nbsp;&nbsp;`Print #ff, "TESTWRITE"`  
&nbsp;&nbsp;&nbsp;&nbsp;`Close #ff`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`'FileCopy FName, localFName`  
&nbsp;&nbsp;&nbsp;&nbsp;`'FileCopy localFName, FName`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo CantDelete`  
&nbsp;&nbsp;&nbsp;&nbsp;`Kill fName`  
&nbsp;&nbsp;&nbsp;&nbsp;`IsPathWritable = "Modifiable"`  
&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`CantDelete:`  
&nbsp;&nbsp;&nbsp;&nbsp;`IsPathWritable = "Writeable"`  
&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`CantWrite:`  
&nbsp;&nbsp;&nbsp;&nbsp;`IsPathWritable = "Readable"`  
&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`ErrHandler:`  
&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
`End Function`  

