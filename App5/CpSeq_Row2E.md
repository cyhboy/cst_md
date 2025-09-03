&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub CpSeq_Row2E()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrorHandler`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`CpSeq`](CpSeq)  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`Row2E`](Row2E)  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`ErrorHandler:`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 10`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Common >> CpSeq >> CpSeq_Row2E**==


# BeCaller
- CpSeq_Row2E{S}(6)->[[CpSeq]]{S}
- CpSeq_Row2E{S}(7)->[[Row2E]]{S}
- CpSeq_Row2E{S}(10)->[[MyMsgBox]]{S}

