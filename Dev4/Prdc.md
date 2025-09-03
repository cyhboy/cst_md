&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub Prdc()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Application.WindowState = xlNormal`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Application.WindowState <> xlMinimized Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Application.WindowState = xlMinimized`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`UnHF`](UnHF)  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`FiltW`](FiltW)  
&nbsp;&nbsp;&nbsp;&nbsp;[`SltX`](SltX)` 9`  
&nbsp;&nbsp;&nbsp;&nbsp;`'Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Call FVC`  
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
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Call Slt9`  
&nbsp;&nbsp;&nbsp;&nbsp;[`SltX`](SltX)` 9`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`RunAppRA`](RunAppRA)  
&nbsp;&nbsp;&nbsp;&nbsp;`' Call Slt9`  
&nbsp;&nbsp;&nbsp;&nbsp;[`SltX`](SltX)` 9`  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`Robp`](Robp)  
&nbsp;&nbsp;&nbsp;&nbsp;`' Call Slt9`  
&nbsp;&nbsp;&nbsp;&nbsp;[`SltX`](SltX)` 9`  
&nbsp;&nbsp;&nbsp;&nbsp;[`RunAppRA`](RunAppRA)` "next"`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Call Slt9`  
&nbsp;&nbsp;&nbsp;&nbsp;[`SltX`](SltX)` 9`  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`DelFs`](DelFs)  
&nbsp;&nbsp;&nbsp;&nbsp;`' Call Slt9`  
&nbsp;&nbsp;&nbsp;&nbsp;[`SltX`](SltX)` 9`  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`Robp`](Robp)  
&nbsp;&nbsp;&nbsp;&nbsp;`' Call FVC`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`UnHF`](UnHF)  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`FiltW`](FiltW)  
&nbsp;&nbsp;&nbsp;&nbsp;[`SltX`](SltX)` 9`  
&nbsp;&nbsp;&nbsp;&nbsp;`silentMode = True`  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`Lrc`](Lrc)  
&nbsp;&nbsp;&nbsp;&nbsp;`silentMode = False`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`UnHF`](UnHF)  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`FiltF`](FiltF)  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`FiltQ`](FiltQ)  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`FiltM`](FiltM)  
&nbsp;&nbsp;&nbsp;&nbsp;[`TopVisible`](TopVisible)` 3, 0.5, 10`  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Auto >> Prdc**==


# BeCaller
- Prdc{S}(8)->[[UnHF]]{S}
- Prdc{S}(9)->[[FiltW]]{S}
- Prdc{S}(10)->[[SltX]]{S}
- Prdc{S}(12)->[[Fold]]{S}
- Prdc{S}(14)->[[UnHF]]{S}
- Prdc{S}(15)->[[FiltF]]{S}
- Prdc{S}(16)->[[FiltQ]]{S}
- Prdc{S}(17)->[[FiltM]]{S}
- Prdc{S}(18)->[[TopVisible]]{S}
- Prdc{S}(19)->[[SltX]]{S}
- Prdc{S}(20)->[[RunAppRA]]{S}
- Prdc{S}(21)->[[SltX]]{S}
- Prdc{S}(22)->[[Robp]]{S}
- Prdc{S}(23)->[[SltX]]{S}
- Prdc{S}(24)->[[RunAppRA]]{S}
- Prdc{S}(25)->[[SltX]]{S}
- Prdc{S}(26)->[[DelFs]]{S}
- Prdc{S}(27)->[[SltX]]{S}
- Prdc{S}(28)->[[Robp]]{S}
- Prdc{S}(29)->[[UnHF]]{S}
- Prdc{S}(30)->[[FiltW]]{S}
- Prdc{S}(31)->[[SltX]]{S}
- Prdc{S}(33)->[[Lrc]]{S}
- Prdc{S}(35)->[[UnHF]]{S}
- Prdc{S}(36)->[[FiltF]]{S}
- Prdc{S}(37)->[[FiltQ]]{S}
- Prdc{S}(38)->[[FiltM]]{S}
- Prdc{S}(39)->[[TopVisible]]{S}

