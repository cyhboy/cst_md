&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub Touch()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`'On Error GoTo ErrorHandler`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim filename As String`
&nbsp;&nbsp;&nbsp;&nbsp;`filename = Cells(currentRow, 11)`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim localFolder As String`
&nbsp;&nbsp;&nbsp;&nbsp;`localFolder = Cells(currentRow, 9)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim filePath As String`
&nbsp;&nbsp;&nbsp;&nbsp;`filePath = localFolder & filename`
&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(filename, ".doc") > 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`TouchDoc`](TouchDoc)
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim wb As Workbook`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim appExcel As New Application`
&nbsp;&nbsp;&nbsp;&nbsp;`appExcel.Visible = False`
&nbsp;&nbsp;&nbsp;&nbsp;`appExcel.EnableEvents = False`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Set wb = appExcel.Workbooks.Open(filePath)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`wb.Save`
&nbsp;&nbsp;&nbsp;&nbsp;`wb.Close savechanges:=True`
&nbsp;&nbsp;&nbsp;&nbsp;`appExcel.Quit`
&nbsp;&nbsp;&nbsp;&nbsp;`Set appExcel = Nothing`
&nbsp;&nbsp;&nbsp;&nbsp;
`ErrorHandler:`
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 10`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
`End Sub`


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Document >> Touch**==


# BeCaller
- Touch{S}(14)->[[TouchDoc]]{S}
- Touch{S}(28)->[[MyMsgBox]]{S}

