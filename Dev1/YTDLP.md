&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub YTDLP()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`UnHF`](UnHF)  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim findOut As Range`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set findOut = Cells.Find(What:="yt-dlp --proxy ""socks5://127.0.0.1:7890"" --cookies", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If Not findOut Is Nothing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells.Replace What:="yt-dlp --proxy ""socks5://127.0.0.1:7890"" --cookies", Replacement:="yt-dlp --cookies", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False`  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells.Replace What:="youtube-dl", Replacement:="yt-dlp", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells.Replace What:="yt-dlp --cookies", Replacement:="yt-dlp --proxy ""socks5://127.0.0.1:7890"" --cookies", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Auto >> RnVA >> YTDLP**==


# BeCaller
- YTDLP{S}(5)->[[UnHF]]{S}

