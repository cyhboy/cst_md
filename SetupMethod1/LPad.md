&nbsp;&nbsp;&nbsp;&nbsp;
`Public Function LPad(str As String, strLen As Integer, padStr As String) As String`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim n: n = 0`
&nbsp;&nbsp;&nbsp;&nbsp;`If strLen > Len(str) Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`n = strLen - Len(str)`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`'LPad = String(n, padStr) & str`
&nbsp;&nbsp;&nbsp;&nbsp;`LPad = Replace(Space(n), " ", padStr) & str`
`End Function`

