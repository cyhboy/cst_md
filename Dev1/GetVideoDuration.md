&nbsp;&nbsp;&nbsp;&nbsp;
`Public Function GetVideoDuration(filename As String)`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrorHandler`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim oFolder, ofPName As Variant`
&nbsp;&nbsp;&nbsp;&nbsp;`With CreateObject("Shell.Application")`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Set oFolder = .Namespace(Left(filename, InStrRev(filename, "\") - 1))`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Set ofPName = oFolder.ParseName(Right(filename, Len(filename) - InStrRev(filename, "\")))`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`GetVideoDuration = CDbl(TimeValue(oFolder.GetDetailsOf(ofPName, 27))) * 24# * 60#`
&nbsp;&nbsp;&nbsp;&nbsp;`End With`
`ErrorHandler:`
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`GetVideoDuration = ""`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
`End Function`

