&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub ScpDl()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`ScpDlParam`](ScpDlParam)` True`
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`Xftp`](Xftp)
&nbsp;&nbsp;&nbsp;&nbsp;
`'    If "On" = ReadRegAR() Then`
`'        Dim exer As String`
`'        exer = Cells(currentRow, 16)`
`'        If InStr(exer, "ScpDl") = 0 Then`
`'            Cells(currentRow, 16) = Trim(exer & " " & "ScpDl")`
`'        End If`
`'    End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 17) = "Success"`
`End Sub`


> [!Getting information]
> Ribbon path please refer to ==**Misc >> New Group >> Collaborate >> ScpDl**==


# BeCaller
- ScpDl{S}(7)->[[ScpDlParam]]{S}
- ScpDl{S}(8)->[[Xftp]]{S}

