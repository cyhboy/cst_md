&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub XcopyFolder(ByVal sourceFolder As String, ByVal destFolder As String)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`sourceFolder = `[`NormalizePath`](NormalizePath)`(sourceFolder)`  
&nbsp;&nbsp;&nbsp;&nbsp;`destFolder = `[`NormalizePath`](NormalizePath)`(destFolder)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fso As Object`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set fso = CreateObject("Scripting.FileSystemObject")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If Not fso.FolderExists(sourceFolder) Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Err.Raise vbObjectError + 1, , "sourceFolder not exist: " & sourceFolder`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If Not fso.FolderExists(destFolder) Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`fso.CreateFolder destFolder`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;[`CopyFolderContents`](CopyFolderContents)` fso, sourceFolder, destFolder`  
`End Sub`  


# BeCaller
- XcopyFolder{S}(5)->[[NormalizePath]]{F}
- XcopyFolder{S}(6)->[[NormalizePath]]{F}
- XcopyFolder{S}(15)->[[CopyFolderContents]]{S}

