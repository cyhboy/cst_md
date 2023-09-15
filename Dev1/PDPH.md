&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub PDPH()`
&nbsp;&nbsp;&nbsp;&nbsp;`' pandas post handler`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`UnHF`](UnHF)
&nbsp;&nbsp;&nbsp;&nbsp;`Cells.Replace What:="_x000D_", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False`
`End Sub`


# BeCaller
- PDPH{S}(5)->[[UnHF]]{S}

