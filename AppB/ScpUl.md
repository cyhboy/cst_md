&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub ScpUl()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;[`ScpUlParam`](ScpUlParam)` True`  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`Xftp`](Xftp)  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'    If "On" = ReadRegAR() Then`  
`'        Dim currentRow As Integer`  
`'        currentRow = ActiveCell.Row`  
`'        Dim exer As String`  
`'        exer = Cells(currentRow, 16)`  
`'        If InStr(exer, "ScpUl") = 0 Then`  
`'            Cells(currentRow, 16) = Trim(exer & " " & "ScpUl")`  
`'        End If`  
`'    End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Misc >> New Group >> Collaborate >> ScpUl**==


# BeCaller
- ScpUl{S}(5)->[[ScpUlParam]]{S}
- ScpUl{S}(6)->[[Xftp]]{S}

