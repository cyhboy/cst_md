&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub MvFilesAndSubfolders(sourceFolder As String, targetFolder As String, wildStr As String)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fso As Object`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim source As Object`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim myFile As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim mySubfolder As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Create a FileSystemObject`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set fso = CreateObject("Scripting.FileSystemObject")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Get the source folder`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set source = fso.GetFolder(sourceFolder)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Loop through each file in the source folder`  
&nbsp;&nbsp;&nbsp;&nbsp;`For Each myFile In source.Files`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If myFile.Name Like wildStr Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' Move the file to the destination folder`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`fso.MoveFile myFile.path, targetFolder & myFile.Name`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Next myFile`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Loop through each subfolder in the source folder`  
&nbsp;&nbsp;&nbsp;&nbsp;`For Each mySubfolder In source.SubFolders`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' Move the subfolder to the destination folder`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`fso.MoveFolder mySubfolder.path, targetFolder & mySubfolder.Name`  
&nbsp;&nbsp;&nbsp;&nbsp;`Next mySubfolder`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Clean up`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set fso = Nothing`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set source = Nothing`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`MsgBox "files and subfolders have been moved."`  
`End Sub`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  

