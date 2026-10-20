&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub RplFil()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrorHandler`  
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
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`RobotRun`](RobotRun.md)` "RplFil"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Next curCell`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim orgTxt As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim newTxt As String`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`orgTxt = Cells(currentRow, 24)`  
&nbsp;&nbsp;&nbsp;&nbsp;`newTxt = Cells(currentRow, 25)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If orgTxt = newTxt Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim localPath As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`localPath = Cells(currentRow, 9)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fileName As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`fileName = Cells(currentRow, 11)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim resultFlag As Boolean`  
&nbsp;&nbsp;&nbsp;&nbsp;`resultFlag = `[`RplTxt4Fil`](RplTxt4Fil.md)`(localPath & fileName, orgTxt, newTxt)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`ErrorHandler:`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox.md)` Err.Number & " " & Err.Description, 30`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`End Sub`  


# BeCaller
- RplFil{S}(16)->[[RobotRun]]{S}
- RplFil{S}(35)->[[RplTxt4Fil]]{F}
- RplFil{S}(38)->[[MyMsgBox]]{S}

