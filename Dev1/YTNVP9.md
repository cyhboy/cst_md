&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub YTNVP9()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Call UnHF`  
&nbsp;&nbsp;&nbsp;&nbsp;`Cells.Replace What:="bestvideo[ext=mp4]+", Replacement:="bestvideo[ext=mp4][vcodec!~=vp09][vcodec!~=av01]+", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False`  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Auto >> RnVA >> YTNVP9**==

