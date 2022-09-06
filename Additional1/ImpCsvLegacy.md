&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub ImpCsvLegacy()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrorHandler`
&nbsp;&nbsp;&nbsp;&nbsp;`Application.ScreenUpdating = False`
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`CleanRgn`](CleanRgn)
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim filePath As String`
&nbsp;&nbsp;&nbsp;&nbsp;`filePath = `[`FnLstFil`](FnLstFil)`("C:\BAK\*.csv")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim ff As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`ff = FreeFile()`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Open filePath For Input As #ff`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim rowNo As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim lineStr As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim lineItems As Variant`
&nbsp;&nbsp;&nbsp;&nbsp;`rowNo = 0`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim colNo As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim cllVal As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Do Until EOF(ff)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Line Input #ff, lineStr`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(lineStr, "'") > 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'lineStr = Left(lineStr, Len(lineStr) - 1)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'lineStr = Right(lineStr, Len(lineStr) - 1)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'lineItems = Split(lineStr, "','")`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`lineItems = Split(lineStr, ",")`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Else`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`lineItems = Split(lineStr, ",")`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`For colNo = 0 To UBound(lineItems)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`cllVal = lineItems(colNo)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If StartsWith(cllVal, "'") And EndsWith(cllVal, "'") Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`cllVal = Left(cllVal, Len(cllVal) - 1)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`cllVal = Right(cllVal, Len(cllVal) - 1)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If IsNumeric(cllVal) Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`cllVal = "'" & cllVal`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(rowNo + 1, colNo + 1).Value = cllVal`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Next colNo`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`rowNo = rowNo + 1`
&nbsp;&nbsp;&nbsp;&nbsp;`Loop`
&nbsp;&nbsp;&nbsp;&nbsp;`Close #ff`
&nbsp;&nbsp;&nbsp;&nbsp;`Application.ScreenUpdating = True`
`ErrorHandler:`
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 30`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
`End Sub`


# BeCaller
- ImpCsvLegacy{S}(7)->[[CleanRgn]]{S}
- ImpCsvLegacy{S}(9)->[[FnLstFil]]{F}
- ImpCsvLegacy{S}(43)->[[MyMsgBox]]{S}

