&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub TopVisible(colNum As Integer, rate As Double, Optional maxNum As Integer = 60)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim rng As Range`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim visibleTotal As Double`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim visibleAvg As Double`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim visibleVar As Double`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim visibleStd As Double`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim visibleCnt As Long`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim visibleTop As Double`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim topCnt As Integer`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;[`SltX`](SltX)` colNum`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set rng = Selection.Cells`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`visibleTotal = Application.WorksheetFunction.Sum(rng.SpecialCells(xlCellTypeVisible))`  
&nbsp;&nbsp;&nbsp;&nbsp;`visibleAvg = Application.WorksheetFunction.Average(rng.SpecialCells(xlCellTypeVisible))`  
&nbsp;&nbsp;&nbsp;&nbsp;`visibleVar = Application.WorksheetFunction.Var(rng.SpecialCells(xlCellTypeVisible))`  
&nbsp;&nbsp;&nbsp;&nbsp;`visibleStd = Application.WorksheetFunction.StDev(rng.SpecialCells(xlCellTypeVisible))`  
&nbsp;&nbsp;&nbsp;&nbsp;`visibleCnt = Application.WorksheetFunction.count(rng.SpecialCells(xlCellTypeVisible))`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If visibleCnt <= 10 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`topCnt = `[`Fix`](Fix)`(visibleCnt * rate)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If topCnt > maxNum + 1 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`topCnt = maxNum + 1`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`visibleTop = Application.WorksheetFunction.Large(rng.SpecialCells(xlCellTypeVisible), topCnt)`  
&nbsp;&nbsp;&nbsp;&nbsp;`' print to the immediate window`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Debug.Print visibleTotal`  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox visibleTotal`  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox visibleAvg`  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox visibleVar`  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox visibleStd`  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox visibleCnt`  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox visibleTop`  
&nbsp;&nbsp;&nbsp;&nbsp;`Filt colNum, ">" & visibleTop`  
`End Sub`  


# BeCaller
- TopVisible{S}(13)->[[SltX]]{S}
- TopVisible{S}(23)->[[Fix]]{F}

