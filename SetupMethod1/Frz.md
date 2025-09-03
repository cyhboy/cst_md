&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub Frz()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`UnHF`](UnHF)  
&nbsp;&nbsp;&nbsp;&nbsp;`On Error Resume Next`  
&nbsp;&nbsp;&nbsp;&nbsp;`With ActiveWindow`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If .FreezePanes Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.FreezePanes = False`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.SplitColumn = 9`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.SplitRow = 1`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.FreezePanes = True`  
&nbsp;&nbsp;&nbsp;&nbsp;`End With`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`'Range("A1").Select`  
&nbsp;&nbsp;&nbsp;&nbsp;`'Selection.AutoFilter`  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**NLB >> BkShtCll >> Frz**==


# BeCaller
- Frz{S}(5)->[[UnHF]]{S}

