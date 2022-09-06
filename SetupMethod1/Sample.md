&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub Sample()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'    Dim n As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`'    n = Selection.Count`
&nbsp;&nbsp;&nbsp;&nbsp;`'    If n > 1 Then`
&nbsp;&nbsp;&nbsp;&nbsp;`'        n = Selection.SpecialCells(xlCellTypeVisible).Count`
&nbsp;&nbsp;&nbsp;&nbsp;`'    End If`
&nbsp;&nbsp;&nbsp;&nbsp;`'    If n > 1 Then`
&nbsp;&nbsp;&nbsp;&nbsp;`'        Dim curCell As Range`
&nbsp;&nbsp;&nbsp;&nbsp;`'        For Each curCell In Selection`
&nbsp;&nbsp;&nbsp;&nbsp;`'            If curCell.EntireColumn.Hidden = False And curCell.EntireRow.Hidden = False Then`
&nbsp;&nbsp;&nbsp;&nbsp;`'                curCell.Select`
&nbsp;&nbsp;&nbsp;&nbsp;`'                'MsgBox subName`
&nbsp;&nbsp;&nbsp;&nbsp;`'                RobotRunByParam "Sample"`
&nbsp;&nbsp;&nbsp;&nbsp;`'            End If`
&nbsp;&nbsp;&nbsp;&nbsp;`'        Next curCell`
&nbsp;&nbsp;&nbsp;&nbsp;`'        Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`'    End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If Trim(Cells(currentRow, 14)) = "" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 14).FormulaR1C1 = "=RIGHT(CELL(""filename"", R1C1),LEN(CELL(""filename"", R1C1))-FIND(""]"",CELL(""filename"", R1C1)))"`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If Trim(Cells(currentRow, 12)) = "" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 12).FormulaR1C1 = "=TODAY()"`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If Trim(Cells(currentRow, 15)) = "" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 15) = "'" & `[`LPad`](LPad)`((currentRow - 1) & "", 4, "0")`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If Trim(Cells(currentRow, 9)) = "" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 9).FormulaR1C1 = "=""C:\Deploy\"" & RC[-1] &  ""\"" & RC[-7] & ""_"" & RC[-8] & ""\"""`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If Trim(Cells(currentRow, 1)) = "" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 1) = "U"`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If Trim(Cells(currentRow, 2)) = "" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 2) = "U"`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If Trim(Cells(currentRow, 8)) = "" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 8) = "Task Memo"`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If Trim(Cells(currentRow, 16)) = "" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 16) = "MyTest"`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
`End Sub`


# BeCaller
- Sample{S}(14)->[[LPad]]{F}

