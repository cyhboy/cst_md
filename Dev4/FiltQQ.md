&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub FiltQQ(filtStr As String)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`UnHF`](UnHF)  
`'    SltX 17`  
`'    Dim myArr As Variant`  
`'    myArr = RangeToArray()`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'    Dim inArr As Variant`  
`'    inArr = EmptyStringArray()`  
`'    Dim e As Variant`  
`'    For Each e In myArr`  
`'        If IsInArray(CStr(e), garr) Then`  
`'            inArr = AddToArray(inArr, CStr(e))`  
`'        End If`  
`'    Next e`  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox (Join(garr, "|"))`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Filt 17, Join(garr, "|")`  
&nbsp;&nbsp;&nbsp;&nbsp;[`Filt`](Filt)` 17, filtStr`  
&nbsp;&nbsp;&nbsp;&nbsp;`ActiveWorkbook.Save`  
`End Sub`  


# BeCaller
- FiltQQ{S}(5)->[[UnHF]]{S}
- FiltQQ{S}(6)->[[Filt]]{S}

