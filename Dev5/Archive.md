&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub Archive(itemPath As String, zipFileName As String)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim oApp As Object`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Create a Shell.Application object`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set oApp = CreateObject("Shell.Application")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Create a new zip file (it must be created as an empty file first)`  
`'    Open zipFileName For Output As #1`  
`'    Close #1`  
&nbsp;&nbsp;&nbsp;&nbsp;`Open zipFileName For Output As #1`  
&nbsp;&nbsp;&nbsp;&nbsp;`Close #1`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Copy the file to the zip folder`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim zipfolder As Object`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim zipitem As Object`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Set zipitem = oApp.Namespace(CVar(Left(itemPath, InStrRev(itemPath, "\") - 1))).Items.Item(Dir(itemPath))`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set zipfolder = oApp.Namespace(CVar(zipFileName))`  
&nbsp;&nbsp;&nbsp;&nbsp;`zipfolder.CopyHere zipitem`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`'Keep script waiting until Compressing is done`  
&nbsp;&nbsp;&nbsp;&nbsp;`On Error Resume Next`  
&nbsp;&nbsp;&nbsp;&nbsp;`Do Until zipfolder.Items.count = 1`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Application.Wait (Now + `[`TimeValue`](TimeValue)`("0:00:01"))`  
&nbsp;&nbsp;&nbsp;&nbsp;`Loop`  
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo 0`  
&nbsp;&nbsp;&nbsp;&nbsp;`'Delete the temporary xls file`  
&nbsp;&nbsp;&nbsp;&nbsp;`Kill itemPath`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Cleanup`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set oApp = Nothing`  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Misc >> New Group >> Commander >> Archive**==


# BeCaller
- Archive{S}(16)->[[TimeValue]]{F}

