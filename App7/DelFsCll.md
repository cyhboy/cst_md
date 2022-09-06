&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub DelFsCll()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`MyQuestionBox`](MyQuestionBox)` "delete file in cell? ", "No", "Yes", 10`
&nbsp;&nbsp;&nbsp;&nbsp;`If confirmation = "No" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrorHandler`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim path As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim parameter As String`
&nbsp;&nbsp;&nbsp;&nbsp;`path = "C:\AppFiles\cmdutils\Recycle.exe -f "`
&nbsp;&nbsp;&nbsp;&nbsp;`'path = "Recycle.exe "`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim cell As Object`
&nbsp;&nbsp;&nbsp;&nbsp;`For Each cell In Selection.Cells`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If cell.EntireColumn.Hidden = False And cell.EntireRow.Hidden = False Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`parameter = " " & """" & Replace(cell.Value, Chr(10), """ """) & """"`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`ShellRun`](ShellRun)` path & parameter, False`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`Next cell`
&nbsp;&nbsp;&nbsp;&nbsp;
`ErrorHandler:`
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 30`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
`End Sub`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;


> [!Getting information]
> Ribbon path please refer to ==**Health Check >> Tools >> Cll >> DelFsCll**==


# BeCaller
- DelFsCll{S}(5)->[[MyQuestionBox]]{S}
- DelFsCll{S}(17)->[[ShellRun]]{S}
- DelFsCll{S}(22)->[[MyMsgBox]]{S}

