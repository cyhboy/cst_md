&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub FiltQ()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Call UnHF`  
&nbsp;&nbsp;&nbsp;&nbsp;[`SltX`](SltX)` 17`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim myArr As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`myArr = `[`RangeToArray`](RangeToArray)`()`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim exArr As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`exArr = `[`EmptyStringArray`](EmptyStringArray)`()`  
`'    exArr = `[`AddToArray`](AddToArray)`(exArr, "3b4ZpPCLVm8")`  
`'    exArr = `[`AddToArray`](AddToArray)`(exArr, "FUTLi1wsmsg")`  
`'    exArr = `[`AddToArray`](AddToArray)`(exArr, "OSwhvUpjfa4")`  
`'    exArr = `[`AddToArray`](AddToArray)`(exArr, "lvnMkP8iQ6I")`  
`'    exArr = `[`AddToArray`](AddToArray)`(exArr, "JSu6PFOjogk")`  
`'    exArr = `[`AddToArray`](AddToArray)`(exArr, "iuMewtge6c4")`  
`'    exArr = `[`AddToArray`](AddToArray)`(exArr, "tXrCnZ9N45E")`  
`'    exArr = `[`AddToArray`](AddToArray)`(exArr, "qYKZ_AWkoFE")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`exArr = `[`AddToArray`](AddToArray)`(exArr, "4dT8TO8-aIQ")`  
&nbsp;&nbsp;&nbsp;&nbsp;`exArr = `[`AddToArray`](AddToArray)`(exArr, "m_29ZoE58MM")`  
&nbsp;&nbsp;&nbsp;&nbsp;`exArr = `[`AddToArray`](AddToArray)`(exArr, "NC4nxOrfPdM")`  
&nbsp;&nbsp;&nbsp;&nbsp;`exArr = `[`AddToArray`](AddToArray)`(exArr, "F-muzcZZUbU")`  
&nbsp;&nbsp;&nbsp;&nbsp;`exArr = `[`AddToArray`](AddToArray)`(exArr, "a3ftowBldYA")`  
&nbsp;&nbsp;&nbsp;&nbsp;`exArr = `[`AddToArray`](AddToArray)`(exArr, "GEO8TtuNLTM")`  
&nbsp;&nbsp;&nbsp;&nbsp;`exArr = `[`AddToArray`](AddToArray)`(exArr, "GJBI7GyuCpw")`  
&nbsp;&nbsp;&nbsp;&nbsp;`exArr = `[`AddToArray`](AddToArray)`(exArr, "tjr3SrCg3pc")`  
&nbsp;&nbsp;&nbsp;&nbsp;`exArr = `[`AddToArray`](AddToArray)`(exArr, "1NRgyL9P-S0")`  
&nbsp;&nbsp;&nbsp;&nbsp;`exArr = `[`AddToArray`](AddToArray)`(exArr, "-RO0rDsbLQ4")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim inArr As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`inArr = `[`EmptyStringArray`](EmptyStringArray)`()`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim e As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`For Each e In myArr`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If False = `[`IsInArray`](IsInArray)`(CStr(e), exArr) Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`inArr = `[`AddToArray`](AddToArray)`(inArr, CStr(e))`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Next e`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;[`Filt`](Filt)` 17, Join(inArr, "|")`  
`End Sub`  


# BeCaller
- FiltQ{S}(5)->[[SltX]]{S}
- FiltQ{S}(7)->[[RangeToArray]]{F}
- FiltQ{S}(9)->[[EmptyStringArray]]{F}
- FiltQ{S}(10)->[[AddToArray]]{F}
- FiltQ{S}(11)->[[AddToArray]]{F}
- FiltQ{S}(12)->[[AddToArray]]{F}
- FiltQ{S}(13)->[[AddToArray]]{F}
- FiltQ{S}(14)->[[AddToArray]]{F}
- FiltQ{S}(15)->[[AddToArray]]{F}
- FiltQ{S}(16)->[[AddToArray]]{F}
- FiltQ{S}(17)->[[AddToArray]]{F}
- FiltQ{S}(18)->[[AddToArray]]{F}
- FiltQ{S}(19)->[[AddToArray]]{F}
- FiltQ{S}(21)->[[EmptyStringArray]]{F}
- FiltQ{S}(24)->[[IsInArray]]{F}
- FiltQ{S}(25)->[[AddToArray]]{F}
- FiltQ{S}(28)->[[Filt]]{S}

