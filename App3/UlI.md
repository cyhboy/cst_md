&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub UlI()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'    On Error GoTo ErrorHandler`  
`'    Dim n As Integer`  
`'    n = Selection.count`  
`'    If n > 1 Then`  
`'        n = Selection.SpecialCells(xlCellTypeVisible).count`  
`'    End If`  
`'    If n > 1 Then`  
`'        Dim curCell As Range`  
`'        For Each curCell In Selection`  
`'            If curCell.EntireColumn.Hidden = False And curCell.EntireRow.Hidden = False Then`  
`'                curCell.Select`  
`'                'MsgBox subName`  
`'                RobotRun "UlI"`  
`'            End If`  
`'        Next curCell`  
`'        Exit Sub`  
`'    End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Call FilDt`  
&nbsp;&nbsp;&nbsp;&nbsp;[`UlParam`](UlParam.md)` False`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'ErrorHandler:`  
`'    If Err.Number <> 0 Then`  
`'        MyMsgBox Err.Number & " " & Err.Description, 30`  
`'    End If`  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Visual >> Execution >> UlI**==


# BeCaller
- UlI{S}(5)->[[UlParam]]{S}

