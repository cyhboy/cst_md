&nbsp;&nbsp;&nbsp;&nbsp;
`Public Function FnGetFileLine(filePath As String, theLineNo As Long) As String`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`Const ForReading = 1, ForWriting = 2`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fso As Object`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim ts As Object`
&nbsp;&nbsp;&nbsp;&nbsp;`Set fso = CreateObject("Scripting.FileSystemObject")`
&nbsp;&nbsp;&nbsp;&nbsp;`Set ts = fso.OpenTextFile(filePath, ForReading)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim lineNo As Long`
&nbsp;&nbsp;&nbsp;&nbsp;`lineNo = 0`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim theLine As String`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Do Until ts.AtEndOfStream`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`lineNo = lineNo + 1`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If lineNo = theLineNo Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`theLine = ts.readline`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`GoTo FoundHandler`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`Loop`
`FoundHandler:`
&nbsp;&nbsp;&nbsp;&nbsp;`Set ts = Nothing`
&nbsp;&nbsp;&nbsp;&nbsp;`Set fso = Nothing`
&nbsp;&nbsp;&nbsp;&nbsp;`FnGetFileLine = theLine`
`End Function`

