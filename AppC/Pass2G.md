&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub Pass2G()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim pass As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`pass = `[`GenPass`](GenPass)`()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If 1 = 2 And True = `[`IsHistPass`](IsHistPass)`(pass) Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`Pass2G`](Pass2G)  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 7) = pass`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Misc >> New Group >> Commander >> Pass2G**==


# BeCaller
- Pass2G{S}(8)->[[GenPass]]{F}
- Pass2G{S}(9)->[[IsHistPass]]{F}
- Pass2G{S}(10)->[[Pass2G]]{S}

