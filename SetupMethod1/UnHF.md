&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub UnHF()`
&nbsp;&nbsp;&nbsp;&nbsp;`' unhide and unfilter`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentws As Worksheet`
&nbsp;&nbsp;&nbsp;&nbsp;`Set currentws = ActiveWorkbook.ActiveSheet`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`currentws.Cells.Select`
&nbsp;&nbsp;&nbsp;&nbsp;`Selection.EntireColumn.Hidden = False`
&nbsp;&nbsp;&nbsp;&nbsp;`Selection.EntireRow.Hidden = False`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If currentws.AutoFilterMode = True Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`currentws.Rows("1:1").Select`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`currentws.AutoFilterMode = False`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`currentws.Range("A1").Select`
&nbsp;&nbsp;&nbsp;&nbsp;`If currentws.Cells(1, 1) <> "" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Selection.AutoFilter`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
`End Sub`

