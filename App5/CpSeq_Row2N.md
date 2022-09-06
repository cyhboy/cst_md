&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub CpSeq_Row2N()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrorHandler`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`CpSeq`](CpSeq)
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`Row2N`](Row2N)
&nbsp;&nbsp;&nbsp;&nbsp;
`ErrorHandler:`
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 10`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
`End Sub`


# BeCaller
- CpSeq_Row2N{S}(6)->[[CpSeq]]{S}
- CpSeq_Row2N{S}(7)->[[Row2N]]{S}
- CpSeq_Row2N{S}(10)->[[MyMsgBox]]{S}

