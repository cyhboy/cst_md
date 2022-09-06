&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub RunAppRA()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 19) = "'" & Cells(currentRow, 18)`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim parameter As String`
&nbsp;&nbsp;&nbsp;&nbsp;`parameter = Cells(currentRow, 10)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim arr As Variant`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`arr = Split(parameter, Chr(10))`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim path As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim i As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`For i = 0 To UBound(arr)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`path = path & arr(i) & "&"`
&nbsp;&nbsp;&nbsp;&nbsp;`Next i`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`While Right(path, 1) = "&"`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`path = Left(path, Len(path) - 1)`
&nbsp;&nbsp;&nbsp;&nbsp;`Wend`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If Not Cells(currentRow, 9).HasFormula Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If Dir(Cells(currentRow, 9), vbDirectory) <> vbNullString Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Dim cdPath As String`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`cdPath = Cells(currentRow, 9)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`path = "cd " & cdPath & "&" & path`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox path`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 18) = "'" & `[`ShellRunResult`](ShellRunResult)`(path, "C:\BAK\cmd.log", True)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 12) = `[`LastModDate`](LastModDate)`("C:\BAK\cmd.log")`
`End Sub`
&nbsp;&nbsp;&nbsp;&nbsp;


# BeCaller
- RunAppRA{S}(27)->[[ShellRunResult]]{F}
- RunAppRA{S}(28)->[[LastModDate]]{F}

