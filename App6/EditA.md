&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub EditA()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim path As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`path = `[`GetAppDrive`](GetAppDrive.md)`() & "\EditPlus\editplus.exe -e"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If silentMode = False Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyQuestionBox`](MyQuestionBox.md)` "please select button as editor", "vscode", "editplus", "", 5`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Select Case confirmation`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Case "editplus"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`path = `[`GetAppDrive`](GetAppDrive.md)`() & "\EditPlus\editplus.exe -e"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Case "vscode"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`path = `[`GetAppDrive`](GetAppDrive.md)`() & "\VSCode\code.exe"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Case Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`path = `[`GetAppDrive`](GetAppDrive.md)`() & "\EditPlus\editplus.exe -e"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End Select`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`' path = """" & `[`GetAppDrive`](GetAppDrive.md)`() & "\EditPlus\editplus.exe"`  
&nbsp;&nbsp;&nbsp;&nbsp;[`FileEditParam`](FileEditParam.md)` False, False, path`  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Document >> EditA**==


# BeCaller
- EditA{S}(6)->[[GetAppDrive]]{F}
- EditA{S}(8)->[[MyQuestionBox]]{S}
- EditA{S}(11)->[[GetAppDrive]]{F}
- EditA{S}(13)->[[GetAppDrive]]{F}
- EditA{S}(15)->[[GetAppDrive]]{F}
- EditA{S}(18)->[[FileEditParam]]{S}

