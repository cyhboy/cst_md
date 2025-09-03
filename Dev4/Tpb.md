&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub Tpb()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If coding Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Dim mcode As String`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`mcode = `[`Proc2FilFun`](Proc2FilFun)`("Tpb")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`MsgBox mcode`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`UnHF`](UnHF)  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;[`OBO`](OBO)` 3`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;[`Cntif`](Cntif)` 23, 9`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`FiltW`](FiltW)  
&nbsp;&nbsp;&nbsp;&nbsp;[`SltX`](SltX)` 9`  
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
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Auto >> Tp >> Tpb**==


# BeCaller
- Tpb{S}(7)->[[Proc2FilFun]]{F}
- Tpb{S}(11)->[[UnHF]]{S}
- Tpb{S}(12)->[[OBO]]{S}
- Tpb{S}(13)->[[Cntif]]{S}
- Tpb{S}(14)->[[FiltW]]{S}
- Tpb{S}(15)->[[SltX]]{S}
- Tpb{S}(17)->[[Fold]]{S}
- Tpb{S}(19)->[[UnHF]]{S}
- Tpb{S}(20)->[[FiltF]]{S}
- Tpb{S}(21)->[[FiltQ]]{S}
- Tpb{S}(22)->[[FiltM]]{S}
- Tpb{S}(23)->[[TopVisible]]{S}

