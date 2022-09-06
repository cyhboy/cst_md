&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub EditA()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim path As String`
&nbsp;&nbsp;&nbsp;&nbsp;`path = """" & `[`GetAppDrive`](GetAppDrive)`() & "\EditPlus\editplus.exe" & """ -e"`
&nbsp;&nbsp;&nbsp;&nbsp;`' path = """" & `[`GetAppDrive`](GetAppDrive)`() & "\EditPlus\editplus.exe"`
&nbsp;&nbsp;&nbsp;&nbsp;[`FileEditParam`](FileEditParam)` False, False, path`
`End Sub`


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Document >> EditA**==


# BeCaller
- EditA{S}(6)->[[GetAppDrive]]{F}
- EditA{S}(7)->[[FileEditParam]]{S}

