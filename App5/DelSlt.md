&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub DelSlt(Optional control As IRibbonControl)`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim curCell As Range`
&nbsp;&nbsp;&nbsp;&nbsp;`For Each curCell In Selection`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If curCell.EntireColumn.Hidden = False And curCell.EntireRow.Hidden = False Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`curCell.Font.Strikethrough = Not curCell.Font.Strikethrough`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`Next curCell`
`End Sub`


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Common >> DelSlt**==

