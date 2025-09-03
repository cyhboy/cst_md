&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub DrawTbl()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Range("A1").Select`  
&nbsp;&nbsp;&nbsp;&nbsp;`Range(Selection, Selection.End(xlToRight)).Select`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`InsCnoSlt`](InsCnoSlt)  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Selection.Font.Bold = True`  
&nbsp;&nbsp;&nbsp;&nbsp;`With Selection.Interior`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.Pattern = xlSolid`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.PatternColorIndex = xlAutomatic`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.Color = RGB(150, 150, 150)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.TintAndShade = 0`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.PatternTintAndShade = 0`  
&nbsp;&nbsp;&nbsp;&nbsp;`End With`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If Range("A2") <> "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Range(Selection, Selection.End(xlDown)).Select`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Selection.Borders(xlDiagonalDown).LineStyle = xlNone`  
&nbsp;&nbsp;&nbsp;&nbsp;`Selection.Borders(xlDiagonalUp).LineStyle = xlNone`  
&nbsp;&nbsp;&nbsp;&nbsp;`With Selection.Borders(xlEdgeLeft)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.LineStyle = xlContinuous`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.ColorIndex = 0`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.TintAndShade = 0`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.Weight = xlThin`  
&nbsp;&nbsp;&nbsp;&nbsp;`End With`  
&nbsp;&nbsp;&nbsp;&nbsp;`With Selection.Borders(xlEdgeTop)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.LineStyle = xlContinuous`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.ColorIndex = 0`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.TintAndShade = 0`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.Weight = xlThin`  
&nbsp;&nbsp;&nbsp;&nbsp;`End With`  
&nbsp;&nbsp;&nbsp;&nbsp;`With Selection.Borders(xlEdgeBottom)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.LineStyle = xlContinuous`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.ColorIndex = 0`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.TintAndShade = 0`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.Weight = xlThin`  
&nbsp;&nbsp;&nbsp;&nbsp;`End With`  
&nbsp;&nbsp;&nbsp;&nbsp;`With Selection.Borders(xlEdgeRight)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.LineStyle = xlContinuous`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.ColorIndex = 0`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.TintAndShade = 0`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.Weight = xlThin`  
&nbsp;&nbsp;&nbsp;&nbsp;`End With`  
&nbsp;&nbsp;&nbsp;&nbsp;`With Selection.Borders(xlInsideVertical)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.LineStyle = xlContinuous`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.ColorIndex = 0`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.TintAndShade = 0`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.Weight = xlThin`  
&nbsp;&nbsp;&nbsp;&nbsp;`End With`  
&nbsp;&nbsp;&nbsp;&nbsp;`With Selection.Borders(xlInsideHorizontal)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.LineStyle = xlContinuous`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.ColorIndex = 0`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.TintAndShade = 0`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.Weight = xlThin`  
&nbsp;&nbsp;&nbsp;&nbsp;`End With`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`With Selection`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.HorizontalAlignment = xlGeneral`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.VerticalAlignment = xlTop`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.WrapText = False`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.Orientation = 0`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.AddIndent = False`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.IndentLevel = 0`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.ShrinkToFit = False`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.ReadingOrder = xlContext`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.MergeCells = False`  
&nbsp;&nbsp;&nbsp;&nbsp;`End With`  
&nbsp;&nbsp;&nbsp;&nbsp;`With Selection`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.HorizontalAlignment = xlGeneral`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.VerticalAlignment = xlTop`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.WrapText = True`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.Orientation = 0`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.AddIndent = False`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.IndentLevel = 0`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.ShrinkToFit = False`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.ReadingOrder = xlContext`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.MergeCells = False`  
&nbsp;&nbsp;&nbsp;&nbsp;`End With`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Rows("1:1").Select`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Range("A2") <> "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Range(Selection, Selection.End(xlDown)).Select`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Selection.RowHeight = 25`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`End Sub`  


# BeCaller
- DrawTbl{S}(7)->[[InsCnoSlt]]{S}

