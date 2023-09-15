&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub ChOp()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim path As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim parameter As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`
&nbsp;&nbsp;&nbsp;&nbsp;`'path = "firefox.exe "`
&nbsp;&nbsp;&nbsp;&nbsp;`'path = "chrome.exe "`
&nbsp;&nbsp;&nbsp;&nbsp;`path = "C:\AppFiles\chrome-win\chrome.exe "`
&nbsp;&nbsp;&nbsp;&nbsp;`parameter = Cells(currentRow, 10)`
&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(parameter, "http") > 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`parameter = `[`CutStrByStartEnd`](CutStrByStartEnd)`(parameter, "http", "$$", True)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(parameter, Chr(10)) > 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`parameter = `[`CutStrByStartEnd`](CutStrByStartEnd)`(parameter, "http", Chr(10), True)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(parameter, """") > 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`parameter = `[`CutStrByStartEnd`](CutStrByStartEnd)`(parameter, "http", """", True)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`Else`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`parameter = ""`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox parameter`
&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox path & """" & parameter & """"`
&nbsp;&nbsp;&nbsp;&nbsp;[`ShellRunStd`](ShellRunStd)` path & """" & parameter & """"`
`End Sub`


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Auto >> ChOp**==


# BeCaller
- ChOp{S}(12)->[[CutStrByStartEnd]]{F}
- ChOp{S}(14)->[[CutStrByStartEnd]]{F}
- ChOp{S}(17)->[[CutStrByStartEnd]]{F}
- ChOp{S}(22)->[[ShellRunStd]]{S}

