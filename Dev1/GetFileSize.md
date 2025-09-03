&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Function GetFileSize(fileName As String)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`On Error Resume Next`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim oFolder, ofPName As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`With CreateObject("Shell.Application")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Set oFolder = .Namespace(Left(fileName, InStrRev(fileName, "\") - 1))`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Set ofPName = oFolder.ParseName(Right(fileName, Len(fileName) - InStrRev(fileName, "\")))`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`GetFileSize = oFolder.GetDetailsOf(ofPName, 1)`  
&nbsp;&nbsp;&nbsp;&nbsp;`End With`  
`End Function`  

