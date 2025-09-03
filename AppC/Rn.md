&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub Rn()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim path As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim parameter As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim sourceFile As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`sourceFile = Cells(currentRow, 9) & Cells(currentRow, 11)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fso As Object`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fileObject As Object`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set fso = CreateObject("Scripting.FileSystemObject")`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set fileObject = fso.GetFile(sourceFile)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim arrFile As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`arrFile = Right(sourceFile, Len(sourceFile) - InStrRev(sourceFile, "\"))`  
&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox arrFile`  
&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(arrFile, ".") > 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`arrFile = Left(sourceFile, InStrRev(sourceFile, ".") - 1)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim dateStr As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`dateStr = Format(fileObject.DateLastModified, "yyyy-MM-dd")`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim targetFile As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`targetFile = Replace(sourceFile, arrFile, arrFile & "_" & dateStr)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`fso.MoveFile sourceFile, targetFile`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Set fileObject = Nothing`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set fso = Nothing`  
&nbsp;&nbsp;&nbsp;&nbsp;`MsgBox sourceFile & " to " & targetFile & " move successfully"`  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Extra >> Common Extra >> Rn**==

