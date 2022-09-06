&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub OpX()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`UnHF`](UnHF)
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim rplStr As String`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim findOut As Range`
&nbsp;&nbsp;&nbsp;&nbsp;`Set findOut = Cells.Find(What:="youtube-dl --cookies", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If Not findOut Is Nothing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyQuestionBox`](MyQuestionBox)` "keep original video/audio file or not?", "Yes", "No", 5`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If confirmation = "Yes" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`rplStr = "youtube-dl -k --cookies"`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyQuestionBox`](MyQuestionBox)` "how about generate audio file after download?", "Yes", "No", 5`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If confirmation = "Yes" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`rplStr = Replace(rplStr, " --cookies", " -x --audio-format flac --cookies")`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyQuestionBox`](MyQuestionBox)` "how about apply compression to all audio file?", "Yes", "No", 5`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If confirmation = "Yes" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`rplStr = Replace(rplStr, "--audio-format flac ", "")`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells.Replace What:="youtube-dl --cookies", Replacement:=rplStr, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False`
&nbsp;&nbsp;&nbsp;&nbsp;`Else`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells.Replace What:="youtube-dl * --cookies", Replacement:="youtube-dl --cookies", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
`End Sub`
&nbsp;&nbsp;&nbsp;&nbsp;


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Auto >> DLF >> OpX**==


# BeCaller
- OpX{S}(5)->[[UnHF]]{S}
- OpX{S}(10)->[[MyQuestionBox]]{S}
- OpX{S}(14)->[[MyQuestionBox]]{S}
- OpX{S}(17)->[[MyQuestionBox]]{S}

