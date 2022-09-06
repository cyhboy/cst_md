&nbsp;&nbsp;&nbsp;&nbsp;
`Public Function IsBackgroudColor(colorValue As Long) As Boolean`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim redVal As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim greenVal As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim blueVal As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`redVal = colorValue Mod 256`
&nbsp;&nbsp;&nbsp;&nbsp;`greenVal = (colorValue \ 256) Mod 256`
&nbsp;&nbsp;&nbsp;&nbsp;`blueVal = colorValue \ 65536`
&nbsp;&nbsp;&nbsp;&nbsp;`If redVal + greenVal + blueVal >= 255 * 3 * 0.8 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`IsBackgroudColor = True`
&nbsp;&nbsp;&nbsp;&nbsp;`Else`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`IsBackgroudColor = False`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
`End Function`
&nbsp;&nbsp;&nbsp;&nbsp;

