&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub Wr2Cod()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrorHandler`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim resultStr As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim codeStr As String`
&nbsp;&nbsp;&nbsp;&nbsp;`codeStr = Cells(currentRow, 10)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If Trim(codeStr) = "" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim comment As String`
&nbsp;&nbsp;&nbsp;&nbsp;`comment = Cells(currentRow, 8)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`resultStr = comment & vbCrLf & codeStr`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim localFolder As String`
&nbsp;&nbsp;&nbsp;&nbsp;`localFolder = Cells(currentRow, 9)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If localFolder = "" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`localFolder = "C:\TMP\"`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 9) = localFolder`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`Fold`](Fold)
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim speicalFile As String`
&nbsp;&nbsp;&nbsp;&nbsp;`speicalFile = Cells(currentRow, 11)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If Trim(speicalFile) = "" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`speicalFile = "tmp_" & Format(Now, "yyyyMMdd-HHmmss") & ".sql"`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 11) = speicalFile`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim filePath As String`
&nbsp;&nbsp;&nbsp;&nbsp;`filePath = localFolder & speicalFile`
&nbsp;&nbsp;&nbsp;&nbsp;[`WriteTxt2Code`](WriteTxt2Code)` resultStr, filePath`
&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` "coding finish", 3`
`ErrorHandler:`
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 30`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
`End Sub`


# BeCaller
- Wr2Cod{S}(23)->[[Fold]]{S}
- Wr2Cod{S}(32)->[[WriteTxt2Code]]{S}
- Wr2Cod{S}(33)->[[MyMsgBox]]{S}
- Wr2Cod{S}(36)->[[MyMsgBox]]{S}

