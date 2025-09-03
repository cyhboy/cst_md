&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub RnJpgByMemo()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim memo As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`memo = Cells(currentRow, 8)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim sourceFile As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim targetFile As String`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`sourceFile = `[`GetLstFilenameByKw`](GetLstFilenameByKw)`("C:\BAK", ".jpg", 1)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fso As Object`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set fso = CreateObject("Scripting.FileSystemObject")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim arrFile As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`arrFile = Right(sourceFile, Len(sourceFile) - InStrRev(sourceFile, "\"))`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox arrFile`  
&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(arrFile, ".") > 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`arrFile = Left(arrFile, InStr(arrFile, ".") - 1)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`targetFile = Replace(sourceFile, arrFile, memo)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`fso.MoveFile sourceFile, targetFile`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Set fso = Nothing`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`MsgBox sourceFile & " to " & targetFile & " move successfully"`  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Misc >> New Group >> RnJpgByMemo**==


# BeCaller
- RnJpgByMemo{S}(11)->[[GetLstFilenameByKw]]{F}

