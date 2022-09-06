&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub Colr()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`If ActiveSheet.Tab.ColorIndex >= 18 Or ActiveSheet.Tab.ColorIndex < 1 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ActiveSheet.Tab.ColorIndex = 1`
&nbsp;&nbsp;&nbsp;&nbsp;`Else`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ActiveSheet.Tab.ColorIndex = CInt(ActiveSheet.Tab.ColorIndex + 1)`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
`End Sub`


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Document >> Colr**==

