&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub PDPH()`  
&nbsp;&nbsp;&nbsp;&nbsp;`' pandas post handler`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`UnHF`](UnHF)  
&nbsp;&nbsp;&nbsp;&nbsp;`Cells.Replace What:="_x000D_", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False`  
&nbsp;&nbsp;&nbsp;&nbsp;`Cells.Replace What:="xa0", Replacement:=WorksheetFunction.Unichar(160), LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Auto >> RnVA >> PDPH**==


# BeCaller
- PDPH{S}(5)->[[UnHF]]{S}

