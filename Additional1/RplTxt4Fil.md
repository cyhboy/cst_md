&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub RplTxt4Fil(sFileName As String, orgTxt As String, newTxt As String)`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`'Call by RplTxt4Fld`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim sBuf As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim sTemp As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim ff As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`ff = FreeFile`
&nbsp;&nbsp;&nbsp;&nbsp;`Open sFileName For Input As #ff`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Do Until EOF(ff)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Line Input #ff, sBuf`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`sTemp = sTemp & sBuf & vbCrLf`
&nbsp;&nbsp;&nbsp;&nbsp;`Loop`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`sTemp = Left(sTemp, Len(sTemp) - Len(vbCrLf))`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Close #ff`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`sTemp = Replace(sTemp, orgTxt, newTxt)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`ff = FreeFile`
&nbsp;&nbsp;&nbsp;&nbsp;`Open sFileName For Output As #ff`
&nbsp;&nbsp;&nbsp;&nbsp;`Print #ff, sTemp`
&nbsp;&nbsp;&nbsp;&nbsp;`Close #ff`
&nbsp;&nbsp;&nbsp;&nbsp;
`End Sub`

