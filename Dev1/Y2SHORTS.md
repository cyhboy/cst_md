&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub Y2SHORTS()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`UnHF`](UnHF)  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim findOut_youtube As Range`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set findOut_youtube = Cells.Find(What:="www.youtube.com", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Not findOut_youtube Is Nothing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`RDR`](RDR)` 17`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim findOut As Range`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set findOut = Cells.Find(What:="C:\Shorts\", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Not findOut Is Nothing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells.Replace What:="C:\Shorts\", Replacement:="C:\Users\" & Environ$("username") & "\Desktop\youtube\", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False`  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells.Replace What:="C:\Users\" & Environ$("username") & "\Desktop\youtube\", Replacement:="C:\Shorts\", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Common >> DLF >> Y2SHORTS**==


# BeCaller
- Y2SHORTS{S}(5)->[[UnHF]]{S}
- Y2SHORTS{S}(9)->[[RDR]]{S}

