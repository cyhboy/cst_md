&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub Tpc()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'    If Application.WindowState <> xlMinimized Then`  
`'        Application.WindowState = xlMinimized`  
`'    End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Call UnHF`  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`Y2SONGS`](Y2SONGS)  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;[`OBO`](OBO)` 3`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;[`Cntif`](Cntif)` 23, 9`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`FiltW`](FiltW)  
&nbsp;&nbsp;&nbsp;&nbsp;[`SltX`](SltX)` 9`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`silentMode = True`  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`Fold`](Fold)  
&nbsp;&nbsp;&nbsp;&nbsp;`silentMode = False`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`UnHF`](UnHF)  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`FiltF`](FiltF)  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`FiltQ`](FiltQ)  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`FiltM`](FiltM)  
&nbsp;&nbsp;&nbsp;&nbsp;[`TopVisible`](TopVisible)` 3, 0.5, 10`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Auto >> Tp >> Tpc**==


# BeCaller
- Tpc{S}(5)->[[Y2SONGS]]{S}
- Tpc{S}(6)->[[OBO]]{S}
- Tpc{S}(7)->[[Cntif]]{S}
- Tpc{S}(8)->[[FiltW]]{S}
- Tpc{S}(9)->[[SltX]]{S}
- Tpc{S}(11)->[[Fold]]{S}
- Tpc{S}(13)->[[UnHF]]{S}
- Tpc{S}(14)->[[FiltF]]{S}
- Tpc{S}(15)->[[FiltQ]]{S}
- Tpc{S}(16)->[[FiltM]]{S}
- Tpc{S}(17)->[[TopVisible]]{S}

