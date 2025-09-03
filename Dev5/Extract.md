&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub Extract(zipFileName As String, fileTargetPath As String)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim oApp As Object`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Ensure target path ends with a separator`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Right(fileTargetPath, 1) <> Application.PathSeparator Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`fileTargetPath = fileTargetPath & Application.PathSeparator`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Create a Shell.Application object`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set oApp = CreateObject("Shell.Application")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Copy items from the zip file to the target directory`  
&nbsp;&nbsp;&nbsp;&nbsp;`oApp.Namespace(CVar(fileTargetPath)).CopyHere oApp.Namespace(CVar(zipFileName)).Items`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Clean up`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set oApp = Nothing`  
`End Sub`  

