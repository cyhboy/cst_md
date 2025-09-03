&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub Y2N()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`UnHF`](UnHF)  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim findOut_douyin As Range`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set findOut_douyin = Cells.Find(What:="www.douyin.com", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Not findOut_douyin Is Nothing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`OBO`](OBO)  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`RDR`](RDR)` 6, 13, 15`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim findOut As Range`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set findOut = Cells.Find(What:="D:\network\", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If Not findOut Is Nothing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells.Replace What:="D:\network\", Replacement:="C:\Users\" & Environ$("username") & "\Desktop\youtube\", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False`  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells.Replace What:="C:\Users\" & Environ$("username") & "\Desktop\youtube\", Replacement:="D:\network\", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Common >> DLF >> Y2N**==


# BeCaller
- Y2N{S}(5)->[[UnHF]]{S}
- Y2N{S}(9)->[[OBO]]{S}
- Y2N{S}(10)->[[RDR]]{S}

