&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub MvFil2Fil(filePath1 As String, filePath2 As String, displayFlag As Boolean)`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'On Error GoTo ErrorHandler`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fso As Object`
&nbsp;&nbsp;&nbsp;&nbsp;`Set fso = CreateObject("Scripting.FileSystemObject")`
&nbsp;&nbsp;&nbsp;&nbsp;`fso.MoveFile filePath1, filePath2`
&nbsp;&nbsp;&nbsp;&nbsp;`Set fso = Nothing`
&nbsp;&nbsp;&nbsp;&nbsp;`If displayFlag Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` filePath1 & " to " & filePath2 & " moved", 5`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`'ErrorHandler:`
&nbsp;&nbsp;&nbsp;&nbsp;`'    If Err.Number <> 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;`'        MyMsgBox Err.Number & " " & Err.Description, 30`
&nbsp;&nbsp;&nbsp;&nbsp;`'    End If`
`End Sub`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;


# BeCaller
- MvFil2Fil{S}(10)->[[MyMsgBox]]{S}

