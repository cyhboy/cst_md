&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub ScpDlI()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;[`ScpDlParam`](ScpDlParam)` True`  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`XftpI`](XftpI)  
&nbsp;&nbsp;&nbsp;&nbsp;`If "On" = `[`ReadRegAR`](ReadRegAR)`() Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Dim exer As String`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`exer = Cells(currentRow, 16)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(exer, "ScpDlI") = 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 16) = Trim(exer & " " & "ScpDlI")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Health Check >> Tools >> ScpDlI**==


# BeCaller
- ScpDlI{S}(5)->[[ScpDlParam]]{S}
- ScpDlI{S}(6)->[[XftpI]]{S}
- ScpDlI{S}(7)->[[ReadRegAR]]{F}

