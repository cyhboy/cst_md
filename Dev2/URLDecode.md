&nbsp;&nbsp;&nbsp;&nbsp;
`Public Function URLDecode(ByVal strEncodedURL As String) As String`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim str As String`
&nbsp;&nbsp;&nbsp;&nbsp;`str = strEncodedURL`
&nbsp;&nbsp;&nbsp;&nbsp;`If Len(str) > 0 Then`
`'        str = Replace(str, "&amp", " & ")`
`'        str = Replace(str, "&#03", Chr(39))`
`'        str = Replace(str, "&quo", Chr(34))`
`'        str = Replace(str, "+", " ")`
`'        str = Replace(str, "%2B", "+")`
`'        str = Replace(str, "%2A", "*")`
`'        str = Replace(str, "%40", "@")`
`'        str = Replace(str, "%2D", "-")`
`'        str = Replace(str, "%5F", "_")`
`'        str = Replace(str, "%2E", ".")`
`'        str = Replace(str, "%2F", "/")`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`str = Replace(str, "%5C", "\")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`URLDecode = str`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
`End Function`

