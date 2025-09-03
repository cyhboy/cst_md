&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub SetRemoveEnvVar()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim nameParam As String, valueParam As String, userParam As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`nameParam = Cells(currentRow, 9)`  
&nbsp;&nbsp;&nbsp;&nbsp;`valueParam = Cells(currentRow, 10)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`userParam = Cells(currentRow, 1)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim objWMI As Object`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim objVar As Object`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Set objWMI = GetObject("winmgmts://./root/cimv2:Win32_Environment")`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set objVar = objWMI.SpawnInstance_`  
&nbsp;&nbsp;&nbsp;&nbsp;`objVar.Name = nameParam`  
&nbsp;&nbsp;&nbsp;&nbsp;`objVar.VariableValue = valueParam`  
&nbsp;&nbsp;&nbsp;&nbsp;`objVar.UserName = userParam`  
&nbsp;&nbsp;&nbsp;&nbsp;`'objVar.SystemVariable      = False`  
&nbsp;&nbsp;&nbsp;&nbsp;`'objVar.Caption      = "GUANGZHOU\asnphpb\JAVA_HOME"`  
&nbsp;&nbsp;&nbsp;&nbsp;`'objVar.Description      = "GUANGZHOU\asnphpb\JAVA_HOME"`  
&nbsp;&nbsp;&nbsp;&nbsp;`objVar.Put_`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Set objVar = Nothing`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set objWMI = Nothing`  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Misc >> New Group >> Commander >> SetRemoveEnvVar**==

