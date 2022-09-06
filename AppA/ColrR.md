&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub ColrR()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim lastCol, firstRow, lastRow As Long`
&nbsp;&nbsp;&nbsp;&nbsp;`lastCol = ActiveSheet.[A1].End(xlToRight).Column`
&nbsp;&nbsp;&nbsp;&nbsp;`firstRow = Selection.Cells(1, 1).Row`
&nbsp;&nbsp;&nbsp;&nbsp;`lastRow = Selection.Cells(Selection.Rows.count, 1).Row`
&nbsp;&nbsp;&nbsp;&nbsp;`Range(Cells(firstRow, 1), Cells(lastRow, lastCol)).Select`
&nbsp;&nbsp;&nbsp;&nbsp;`With Selection.Interior`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox .ColorIndex`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If IsNull(.ColorIndex) Or IsEmpty(.ColorIndex) Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.ColorIndex = 1`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If .ColorIndex >= 56 Or .ColorIndex < 1 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.ColorIndex = 1`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.ColorIndex = .ColorIndex + 1`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`While False = `[`IsBackgroudColor`](IsBackgroudColor)`(.Color) And .ColorIndex < 56`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.ColorIndex = .ColorIndex + 1`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Wend`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox .ColorIndex`
&nbsp;&nbsp;&nbsp;&nbsp;`End With`
`End Sub`


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Auto >> ColrR**==


# BeCaller
- ColrR{S}(18)->[[IsBackgroudColor]]{F}

