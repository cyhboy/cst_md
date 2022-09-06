&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub ShellRunMax(cmd As String)`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`'vbHide 0 Window is hidden and focus is passed to the hidden window. The vbHide constant is not applicable on Macintosh platforms.`
&nbsp;&nbsp;&nbsp;&nbsp;`'vbNormalFocus 1 Window has focus and is restored to its original size and position.`
&nbsp;&nbsp;&nbsp;&nbsp;`'vbMinimizedFocus 2 Window is displayed as an icon with focus.`
&nbsp;&nbsp;&nbsp;&nbsp;`'vbMaximizedFocus 3 Window is maximized with focus.`
&nbsp;&nbsp;&nbsp;&nbsp;`'vbNormalNoFocus 4 Window is restored to its most recent size and position. The currently active window remains active.`
&nbsp;&nbsp;&nbsp;&nbsp;`'vbMinimizedNoFocus 6 Window is displayed as an icon. The currently active window remains active.`
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrorHandler`
&nbsp;&nbsp;&nbsp;&nbsp;`Shell cmd, vbMaximizedFocus`
`ErrorHandler:`
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 30`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
`End Sub`


# BeCaller
- ShellRunMax{S}(9)->[[MyMsgBox]]{S}

