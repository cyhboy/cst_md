&nbsp;&nbsp;&nbsp;&nbsp;
`Public Function InUse(filePath As String) As Boolean`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim ff As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`On Error Resume Next`
&nbsp;&nbsp;&nbsp;&nbsp;`Open filePath For Binary Access Read Lock Read As #ff`
&nbsp;&nbsp;&nbsp;&nbsp;`Close #ff`
&nbsp;&nbsp;&nbsp;&nbsp;`InUse = IIf(Err.Number > 0, True, False)`
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo 0`
`End Function`

