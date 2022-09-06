&nbsp;&nbsp;&nbsp;&nbsp;
`Public Function CntExeRunning(exeName As String) As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`'On Error GoTo ErrorHandler`
&nbsp;&nbsp;&nbsp;&nbsp;`On Error Resume Next`
&nbsp;&nbsp;&nbsp;&nbsp;`'Dim flag As Boolean`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim cnt As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`'cnt = 0`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim strComputer As String`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim objWMI As Object`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim objProcessSet As Object`
&nbsp;&nbsp;&nbsp;&nbsp;`'Dim objProcess As Object`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim strUserName As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim strUserDomain As String`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`strComputer = "."`
&nbsp;&nbsp;&nbsp;&nbsp;`Set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")`
&nbsp;&nbsp;&nbsp;&nbsp;`Set objProcessSet = objWMI.ExecQuery("SELECT Name FROM Win32_Process WHERE Name = '" & exeName & "'")`
&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox objProcessSet.count`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`cnt = objProcessSet.count`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'ErrorHandler:`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'Do nothing as always error`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'MyMsgBox Err.Number & " " & Err.Description, 10`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'cnt = 0`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'MyMsgBox cnt & "", 10`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Set objProcessSet = Nothing`
&nbsp;&nbsp;&nbsp;&nbsp;`Set objWMI = Nothing`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`CntExeRunning = cnt`
`End Function`

