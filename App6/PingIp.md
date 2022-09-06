&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub PingIp()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim path As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim parameter As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`path = "cmd.exe /K Ping -a "`
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'Ping -a 130.29.48.148`
&nbsp;&nbsp;&nbsp;&nbsp;`parameter = Cells(currentRow, 6)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`ShellRun`](ShellRun)` path & parameter, False`
`End Sub`


> [!Getting information]
> Ribbon path please refer to ==**Extra >> Server && Local Extra >> PingIp**==


# BeCaller
- PingIp{S}(11)->[[ShellRun]]{S}

