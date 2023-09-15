&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub DelWb()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`MyQuestionBox`](MyQuestionBox)` "delete activated workbook in row? ", "No", "Yes", 10`
&nbsp;&nbsp;&nbsp;&nbsp;`If confirmation = "No" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim path As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim parameter As String`
&nbsp;&nbsp;&nbsp;&nbsp;`'path = "cmd.exe /C C:\AppFiles\cmdutils\Recycle -f "`
&nbsp;&nbsp;&nbsp;&nbsp;`path = "C:\AppFiles\cmdutils\Recycle.exe -f "`
&nbsp;&nbsp;&nbsp;&nbsp;`'path = "Recycle.exe "`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`parameter = ActiveWorkbook.FullName`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`ActiveWorkbook.Close`
&nbsp;&nbsp;&nbsp;&nbsp;[`ShellRun`](ShellRun)` path & parameter, False`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim exeName As String: exeName = `[`ExtractEXE`](ExtractEXE)`(path)`
&nbsp;&nbsp;&nbsp;&nbsp;`While True = `[`IsExeRunning`](IsExeRunning)`(exeName)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Sleep 3000`
&nbsp;&nbsp;&nbsp;&nbsp;`Wend`
`End Sub`


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Common >> DelWb**==


# BeCaller
- DelWb{S}(5)->[[MyQuestionBox]]{S}
- DelWb{S}(14)->[[ShellRun]]{S}
- DelWb{S}(15)->[[ExtractEXE]]{F}
- DelWb{S}(16)->[[IsExeRunning]]{F}

