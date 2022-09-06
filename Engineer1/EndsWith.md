&nbsp;&nbsp;&nbsp;&nbsp;
`Public Function EndsWith(str As String, ending As String) As Boolean`
`'    If testing Then`
`'        Exit Function`
`'    End If`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim endingLen As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`endingLen = Len(ending)`
&nbsp;&nbsp;&nbsp;&nbsp;`EndsWith = (Right(UCase(str), endingLen) = UCase(ending))`
`End Function`

