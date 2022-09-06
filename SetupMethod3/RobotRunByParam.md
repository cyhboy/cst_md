&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub RobotRunByParam(comm As String)`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`' On Error GoTo ErrorHandler`
&nbsp;&nbsp;&nbsp;&nbsp;`' Application.Run comm`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim comms As Variant`
&nbsp;&nbsp;&nbsp;&nbsp;`comms = Split(comm, "_")`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim i As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`For i = 0 To UBound(comms)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Application.Run comms(i)`
&nbsp;&nbsp;&nbsp;&nbsp;`Next i`
&nbsp;&nbsp;&nbsp;&nbsp;`' ErrorHandler:`
&nbsp;&nbsp;&nbsp;&nbsp;`'    If Err.Number <> 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;`'        MyMsgBox Err.Number & " " & Err.Description & " " & comm, 30`
&nbsp;&nbsp;&nbsp;&nbsp;`'    End If`
`End Sub`

