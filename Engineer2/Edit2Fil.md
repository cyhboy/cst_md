&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub Edit2Fil()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim module As String`
&nbsp;&nbsp;&nbsp;&nbsp;`module = Cells(currentRow, 1)`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim subb As String`
&nbsp;&nbsp;&nbsp;&nbsp;`subb = Cells(currentRow, 2)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim path As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim parameter As String`
&nbsp;&nbsp;&nbsp;&nbsp;`path = """" & `[`GetAppDrive`](GetAppDrive)`() & "\EditPlus\editplus.exe"" "`
&nbsp;&nbsp;&nbsp;&nbsp;`parameter = "C:\SANDBOX\VB_SPACE\VBA_PROJECT\" & Format(Now, "yyyyMMdd") & "\" & module & "\" & subb & ".vb"`
&nbsp;&nbsp;&nbsp;&nbsp;[`ShellRun`](ShellRun)` path & parameter, False`
`End Sub`


> [!Getting information]
> Ribbon path please refer to ==**Engineer >> Project >> Edit2Fil**==


# BeCaller
- Edit2Fil{S}(13)->[[GetAppDrive]]{F}
- Edit2Fil{S}(15)->[[ShellRun]]{S}

