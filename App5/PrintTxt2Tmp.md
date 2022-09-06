&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub PrintTxt2Tmp(text As String, path As String)`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim ff As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`ff = FreeFile`
&nbsp;&nbsp;&nbsp;&nbsp;`Open path For Append As #ff`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Print #ff, text`
&nbsp;&nbsp;&nbsp;&nbsp;`Close #ff`
`End Sub`

