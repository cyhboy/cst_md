&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Function SelectFolder(myPath As String) As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fldr As FileDialog`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim sItem As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set fldr = Application.FileDialog(msoFileDialogFolderPicker)`  
&nbsp;&nbsp;&nbsp;&nbsp;`With fldr`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.Title = "Select a Folder"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.AllowMultiSelect = False`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.InitialFileName = `[`GetParentFolderName`](GetParentFolderName)`(GetParentFolderName(myPath))`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'.InitialFileName = `[`GetParentFolderName`](GetParentFolderName)`(myPath)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If .Show <> -1 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`GoTo NextCode`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`sItem = .SelectedItems.Item(1)`  
&nbsp;&nbsp;&nbsp;&nbsp;`End With`  
`NextCode:`  
&nbsp;&nbsp;&nbsp;&nbsp;`SelectFolder = sItem`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set fldr = Nothing`  
`End Function`  


# BeCaller
- SelectFolder]]{F}(11)->[[GetParentFolderName]]{F}

