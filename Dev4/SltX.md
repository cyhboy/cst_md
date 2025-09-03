&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub SltX(colNum As Integer)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`'ActiveSheet.Cells(1, colNum).Select`  
&nbsp;&nbsp;&nbsp;&nbsp;[`FVC`](FVC)` colNum`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim rng As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim visibleCnt As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set rng = ActiveSheet.AutoFilter.Range`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' count the visible cells. here i just want the row count.`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Application.WorksheetFunction.CountA(rng.offset(1, 0).Resize(rng.Rows.count - 1).SpecialCells(xlCellTypeVisible))`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`visibleCnt = rng.offset(1, 0).Resize(rng.Rows.count - 1).SpecialCells(xlCellTypeVisible).Rows.count`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`If visibleCnt > 1 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ActiveSheet.Range(Selection, Selection.End(xlDown)).Select`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'    Call UnSltTitle`  
&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox Selection.count`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`End Sub`  


# BeCaller
- SltX{S}(5)->[[FVC]]{S}

