&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Function IsExeRunning(exeName As String) As Boolean`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrorHandler`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim flag As Boolean`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim strComputer As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim objWMI As Object, objProcessSet As Object, objProcess As Object`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim strUserName As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim strUserDomain As String`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`strComputer = "."`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set objProcessSet = objWMI.ExecQuery("SELECT Name FROM Win32_Process WHERE Name = '" & exeName & "'")`  
&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox objProcessSet.count`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox Environ$("username")`  
&nbsp;&nbsp;&nbsp;&nbsp;`For Each objProcess In objProcessSet`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`objProcess.GetOwner strUserName, strUserDomain`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox strUserName`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If strUserName = Environ$("username") Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`flag = True`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit For`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox "Process " & objProcess.Name & " is owned by " & strUserDomain & "\" & strUserName & "."`  
&nbsp;&nbsp;&nbsp;&nbsp;`Next objProcess`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`'If objProcessSet.count > 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;`'    flag = True`  
&nbsp;&nbsp;&nbsp;&nbsp;`'Else`  
&nbsp;&nbsp;&nbsp;&nbsp;`'    flag = False`  
&nbsp;&nbsp;&nbsp;&nbsp;`'End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`'    For Each Process In objProcessSet`  
&nbsp;&nbsp;&nbsp;&nbsp;`'        If Process.Name = exeName Then`  
&nbsp;&nbsp;&nbsp;&nbsp;`'            flag = True`  
&nbsp;&nbsp;&nbsp;&nbsp;`'            Exit For`  
&nbsp;&nbsp;&nbsp;&nbsp;`'        End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`'    Next Process`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`ErrorHandler:`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set objProcessSet = Nothing`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set objWMI = Nothing`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`IsExeRunning = True`  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`IsExeRunning = flag`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`End Function`  

