&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub RunAppK()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`' RunAppParam False, False, True`  
&nbsp;&nbsp;&nbsp;&nbsp;`If theKeep Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`RunAppParam`](RunAppParam)` True, False, True`  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`RunAppParam`](RunAppParam)` True, False, False`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Auto >> RA >> RunAppK**==


# BeCaller
- RunAppK{S}(6)->[[RunAppParam]]{S}
- RunAppK{S}(8)->[[RunAppParam]]{S}

