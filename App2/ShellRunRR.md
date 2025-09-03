&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Function ShellRunRR(cmd As String)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim path As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`'path = "cmd.exe /C " & cmd & " > " & cmdLogFile`  
&nbsp;&nbsp;&nbsp;&nbsp;`path = "cmd"`  
&nbsp;&nbsp;&nbsp;&nbsp;`'path = Split(cmd, " ")(0)`  
&nbsp;&nbsp;&nbsp;&nbsp;`'path = Replace(path, "2 >", "2>")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim param As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`param = "/C " & Replace(cmd, """", "")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`MsgBox path`  
&nbsp;&nbsp;&nbsp;&nbsp;`MsgBox param`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`'Exit Function`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`'Shell path, vbNormalFocus`  
&nbsp;&nbsp;&nbsp;&nbsp;`'Shell path, vbHide`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox `[`Redirect`](Redirect)`("cmd", "/c dir")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`ShellRunRR = `[`Redirect`](Redirect)`(path, param)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`End Function`  


# BeCaller
- ShellRunRR]]{F}(11)->[[Redirect]]{F}

