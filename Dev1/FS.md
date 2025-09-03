&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub FS()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim videoPath As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim videoFileName As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim videoFullFilename As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;&nbsp;&nbsp;&nbsp;`videoPath = Cells(currentRow, 9)`  
&nbsp;&nbsp;&nbsp;&nbsp;`videoFileName = Cells(currentRow, 11)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If Trim(videoPath) = "" Or Trim(videoFileName) = "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`videoFullFilename = videoPath & videoFileName`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 8) = `[`GetFileSize`](GetFileSize)`(videoFullFilename)`  
`End Sub`  


# BeCaller
- FS{S}(16)->[[GetFileSize]]{F}

