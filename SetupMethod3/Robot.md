&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub Robot(Optional control As IRibbonControl)`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'On Error GoTo ErrorHandler`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim n As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`n = Selection.count`
&nbsp;&nbsp;&nbsp;&nbsp;`If n > 1 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`n = Selection.SpecialCells(xlCellTypeVisible).count`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`If n > 1 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Dim curCell As Range`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`For Each curCell In Selection`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If curCell.EntireColumn.Hidden = False And curCell.EntireRow.Hidden = False Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`curCell.Select`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox subName`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`RobotRunByParam`](RobotRunByParam)` "Robot"`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Next curCell`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim mycll As Excel.Range`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Set mycll = ActiveCell`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim comms As String`
&nbsp;&nbsp;&nbsp;&nbsp;`'Dim interval As Long`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`
&nbsp;&nbsp;&nbsp;&nbsp;`comms = Cells(currentRow, 16)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim arr As Variant`
&nbsp;&nbsp;&nbsp;&nbsp;`arr = Split(comms, " ")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim i As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim tmp_comm As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim comm As String`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`For i = 0 To UBound(arr)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`tmp_comm = comm`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`comm = arr(i)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(comm, "#") = 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(comms, "CpSeqO") = 0 And InStr(comms, "CpSeq") = 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`mycll.Select`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`RobotRunByParam`](RobotRunByParam)` comm`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`Next i`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'ErrorHandler:`
&nbsp;&nbsp;&nbsp;&nbsp;`'    If Err.Number <> 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;`'       MyMsgBox Err.Number & " " & Err.Description, 30`
&nbsp;&nbsp;&nbsp;&nbsp;`'    End If`
`End Sub`


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Auto >> Robot**==


# BeCaller
- Robot{S}(15)->[[RobotRunByParam]]{S}
- Robot{S}(38)->[[RobotRunByParam]]{S}

