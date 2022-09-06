&nbsp;&nbsp;&nbsp;&nbsp;
`Public Function MatchRegx(text As String, patt As String, Optional ignoreC As Boolean = False) As Boolean`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`'Set up regular expression object`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim RE As New RegExp`
&nbsp;&nbsp;&nbsp;&nbsp;`RE.Pattern = patt`
&nbsp;&nbsp;&nbsp;&nbsp;`RE.Global = True`
&nbsp;&nbsp;&nbsp;&nbsp;`RE.IgnoreCase = ignoreC`
&nbsp;&nbsp;&nbsp;&nbsp;`RE.multiLine = True`
&nbsp;&nbsp;&nbsp;&nbsp;`'Retrieve all matches`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim Matches As MatchCollection`
&nbsp;&nbsp;&nbsp;&nbsp;`Set Matches = RE.Execute(text)`
&nbsp;&nbsp;&nbsp;&nbsp;`'Return the corrected count of matches`
&nbsp;&nbsp;&nbsp;&nbsp;`If Matches.count > 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`MatchRegx = True`
&nbsp;&nbsp;&nbsp;&nbsp;`Else`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`MatchRegx = False`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
`End Function`

