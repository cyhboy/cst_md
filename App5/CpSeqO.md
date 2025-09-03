&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub CpSeqO(Optional control As IRibbonControl)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Application.ScreenUpdating = False`  
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrorHandler`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim count As Integer, countNew As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`'count = Range("A65536").End(xlUp).Row`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim endCount As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`'endCount = 32767`  
&nbsp;&nbsp;&nbsp;&nbsp;`endCount = 3276`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`For count = endCount To 1 Step -1`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If Cells(count, 1) <> "" And Cells(count, 1).EntireColumn.Hidden = False And Cells(count, 1).EntireRow.Hidden = False Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit For`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Next count`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`For countNew = count To endCount`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If Cells(countNew, 1) = "" And Cells(countNew, 1).EntireColumn.Hidden = False And Cells(countNew, 1).EntireRow.Hidden = False Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit For`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Next countNew`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`'If Cells(count, 1) <> vbNullString Then  'skip blank lines`  
&nbsp;&nbsp;&nbsp;&nbsp;`'ActiveSheet.Rows(count + 1).Insert`  
&nbsp;&nbsp;&nbsp;&nbsp;`ActiveSheet.Rows(count).Copy Destination:=ActiveSheet.Rows(countNew)`  
&nbsp;&nbsp;&nbsp;&nbsp;`'End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`'count = Range("A65536").End(xlUp).Row`  
&nbsp;&nbsp;&nbsp;&nbsp;`Range("H" & countNew).Activate`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 15) = "'" & `[`LPad`](LPad)`(Cells(currentRow - 1, 15) + 1, 4, "0")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`ErrorHandler:`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 30`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Application.ScreenUpdating = True`  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Common >> CpSeq >> CpSeqO**==


# BeCaller
- CpSeqO{S}(24)->[[LPad]]{F}
- CpSeqO{S}(27)->[[MyMsgBox]]{S}

