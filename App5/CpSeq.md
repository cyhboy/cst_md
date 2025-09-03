&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub CpSeq(Optional control As IRibbonControl)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Application.ScreenUpdating = False`  
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrorHandler`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentCol As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`currentCol = ActiveCell.Column`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`'ActiveSheet.Rows(currentRow).Copy Destination:=ActiveSheet.Rows(currentRow + 1)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`ActiveSheet.Rows(currentRow).Select`  
&nbsp;&nbsp;&nbsp;&nbsp;`Selection.Copy`  
&nbsp;&nbsp;&nbsp;&nbsp;`ActiveSheet.Rows(currentRow + 1).Select`  
&nbsp;&nbsp;&nbsp;&nbsp;`Selection.Insert Shift:=xlDown`  
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow + 1, currentCol).Select`  
`ErrorHandler:`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 30`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Application.ScreenUpdating = True`  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Common >> CpSeq >> CpSeq**==


# BeCaller
- CpSeq{S}(18)->[[MyMsgBox]]{S}

