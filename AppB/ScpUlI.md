&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub ScpUlI()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;[`ScpUlParam`](ScpUlParam)` True`  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`XftpI`](XftpI)  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'    If "On" = ReadRegAR() Then`  
`'        Dim currentRow As Integer`  
`'        currentRow = ActiveCell.Row`  
`'        Dim exer As String`  
`'        exer = Cells(currentRow, 16)`  
`'        If InStr(exer, "ScpUlI") = 0 Then`  
`'            Cells(currentRow, 16) = Trim(exer & " " & "ScpUlI")`  
`'        End If`  
`'    End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Health Check >> Tools >> ScpUlI**==


# BeCaller
- ScpUlI{S}(5)->[[ScpUlParam]]{S}
- ScpUlI{S}(6)->[[XftpI]]{S}

