&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub Tch_legacy()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fileName As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`fileName = Cells(currentRow, 11)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim localFolder As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`localFolder = Cells(currentRow, 9)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim txtValue As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`txtValue = Cells(currentRow, 10)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fso As Object`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Dim txtFile As Object`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set fso = CreateObject("Scripting.FileSystemObject")`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim filePath1 As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`filePath1 = "C:\AppFiles\linux.txt"`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim result As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Not fso.FileExists(localFolder & fileName) Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`result = fso.copyfile(filePath1, localFolder & fileName)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If result = "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`MsgBox "copied linux file done"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`MsgBox "target file is existing, copy canceled"`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set fso = Nothing`  
`End Sub`  

