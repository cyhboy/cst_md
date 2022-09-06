&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub CpFil2Fil(filePath1 As String, filePath2 As String, displayFlag As Boolean, overrideFlag As Boolean)`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrorHandler`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fso As Object`
&nbsp;&nbsp;&nbsp;&nbsp;`Set fso = CreateObject("Scripting.FileSystemObject")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If Not fso.fileexists(filePath2) Or overrideFlag Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`fso.copyfile filePath1, filePath2`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Set fso = Nothing`
&nbsp;&nbsp;&nbsp;&nbsp;`If displayFlag Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` filePath1 & " to " & filePath2 & " copied", 5`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
`ErrorHandler:`
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 30`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
`End Sub`


# BeCaller
- CpFil2Fil{S}(13)->[[MyMsgBox]]{S}
- CpFil2Fil{S}(17)->[[MyMsgBox]]{S}

