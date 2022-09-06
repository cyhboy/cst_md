&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub PowerShellRun(cmd As String, hold As Boolean)`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`On Error Resume Next`
&nbsp;&nbsp;&nbsp;&nbsp;`' Shell "powershell.exe -ExecutionPolicy Unrestricted -File " & cmd, vbNormal`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox "powershell.exe -ExecutionPolicy Unrestricted -noexit " & """" & "& " & cmd`
&nbsp;&nbsp;&nbsp;&nbsp;`' Shell "cmd.exe /c powershell.exe -ExecutionPolicy Unrestricted -noexit " & """" & "& " & cmd, vbNormal`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Shell "powershell.exe -ExecutionPolicy Unrestricted -noexit " & """" & "& " & cmd, vbNormal`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`' KillExeRunning ExtractEXE("powershell.exe"), 2`
&nbsp;&nbsp;&nbsp;&nbsp;`' KillExeRunning ExtractEXE("conhost.exe"), 1`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'    Dim exeName As String: exeName = ExtractEXE("cmd.exe")`
&nbsp;&nbsp;&nbsp;&nbsp;`'    While True = IsExeRunning(exeName)`
&nbsp;&nbsp;&nbsp;&nbsp;`'        Sleep 3000`
&nbsp;&nbsp;&nbsp;&nbsp;`'    Wend`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`MsgBox "Met a unexpected case: " & Err.Number`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
`End Sub`

