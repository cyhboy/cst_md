&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub ZIP()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim folder_path As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim file_path As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim file_path_zip As String`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`folder_path = Cells(currentRow, 9)`  
&nbsp;&nbsp;&nbsp;&nbsp;`file_path = Cells(currentRow, 11)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;[`Archive`](Archive)` folder_path & file_path, folder_path & file_path & ".zip"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`End Sub`  


# BeCaller
- ZIP{S}(12)->[[Archive]]{S}

