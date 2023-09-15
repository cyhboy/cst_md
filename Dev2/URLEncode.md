&nbsp;&nbsp;&nbsp;&nbsp;
`Public Function URLEncode(ByVal strDecodedURL As String) As String`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim str As String`
&nbsp;&nbsp;&nbsp;&nbsp;`str = strDecodedURL`
&nbsp;&nbsp;&nbsp;&nbsp;`If Len(str) > 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`str = Replace(str, "\", "%5C")`
&nbsp;&nbsp;&nbsp;&nbsp;
`'        str = Replace(str, " & ", "&amp")`
`'        str = Replace(str, Chr(39), "&#03")`
`'        str = Replace(str, Chr(34), "&quo")`
`'        str = Replace(str, "+", "%2B")`
`'        str = Replace(str, " ", "+")`
`'        str = Replace(str, "*", "%2A")`
`'        str = Replace(str, "@", "%40")`
`'        str = Replace(str, "-", "%2D")`
`'        str = Replace(str, "_", "%5F")`
&nbsp;&nbsp;&nbsp;&nbsp;
`'        str = Replace(str, ".", "%2E")`
`'        str = Replace(str, "/", "%2F")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`URLEncode = str`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
`End Function`

