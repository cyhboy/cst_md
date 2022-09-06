&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub CleanRgn()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`ActiveCell.CurrentRegion.Select`
&nbsp;&nbsp;&nbsp;&nbsp;`If Selection.Rows.count > 1 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`UnSltTitle`](UnSltTitle)
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Selection.ClearContents`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
`End Sub`


> [!Getting information]
> Ribbon path please refer to ==**Extra >> Common Extra >> CleanRgn**==


# BeCaller
- CleanRgn{S}(7)->[[UnSltTitle]]{S}

