&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub Xcp()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim sourceFolder As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim destFolder As String`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`sourceFolder = Cells(currentRow, 9)`  
&nbsp;&nbsp;&nbsp;&nbsp;`destFolder = Cells(currentRow + 1, 9)`  
&nbsp;&nbsp;&nbsp;&nbsp;[`XcopyFolder`](XcopyFolder.md)` sourceFolder, destFolder`  
`End Sub`  


# BeCaller
- Xcp{S}(11)->[[XcopyFolder]]{S}

