&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub PlyVA()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
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
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`RobotRunByParam`](RobotRunByParam)` "PlyVA"`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Next curCell`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'GoTo SECOND_STAGE`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim parameter As String`
&nbsp;&nbsp;&nbsp;&nbsp;`parameter = Cells(currentRow, 10)`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim formatCode As String`
&nbsp;&nbsp;&nbsp;&nbsp;`formatCode = `[`CutStrByStartEnd`](CutStrByStartEnd)`(parameter, " best", "http", True)`
&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(parameter, "http") > 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`parameter = `[`CutStrByStartEnd`](CutStrByStartEnd)`(parameter, "http", "$", True)`
&nbsp;&nbsp;&nbsp;&nbsp;`Else`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`parameter = ""`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim cmdStr As String`
&nbsp;&nbsp;&nbsp;&nbsp;`' cmdStr = "conda activate learn"`
&nbsp;&nbsp;&nbsp;&nbsp;`' cmdStr = cmdStr & " && " & "python C:\AppFiles\ipy\plyVA.py """ & parameter & """"`
&nbsp;&nbsp;&nbsp;&nbsp;`' after pyinstaller build the python file`
&nbsp;&nbsp;&nbsp;&nbsp;`cmdStr = "C:\AppFiles\ipy\plyVA\plyVA.exe """ & parameter & """"`
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 18) = "'" & `[`ShellRunResult`](ShellRunResult)`(cmdStr, "C:\BAK\cmd.log", True)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'SECOND_STAGE:`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim jsonStr As String`
&nbsp;&nbsp;&nbsp;&nbsp;`jsonStr = Cells(currentRow, 18)`
&nbsp;&nbsp;&nbsp;&nbsp;`jsonStr = `[`CutStrByStartEnd`](CutStrByStartEnd)`(jsonStr, "{", "$", True)`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim Json As Object`
&nbsp;&nbsp;&nbsp;&nbsp;`Set Json = JsonConverter.ParseJson(jsonStr)`
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 1) = `[`Json`](Json)`("subtitles")`
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 2) = `[`Json`](Json)`("filesizeString")`
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 3) = `[`Json`](Json)`("view_count")`
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 4) = "'" & `[`Json`](Json)`("upload_date")`
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 8) = Replace(Cells(currentRow, 8), formatCode, " " & `[`Json`](Json)`("formatCode") & " ")`
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 10) = Replace(Cells(currentRow, 10), formatCode, " " & `[`Json`](Json)`("formatCode") & " ")`
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 13) = "'" & `[`Json`](Json)`("videoFileName")`
&nbsp;&nbsp;&nbsp;&nbsp;
`End Sub`


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Auto >> PlyVA**==


# BeCaller
- PlyVA{S}(15)->[[RobotRunByParam]]{S}
- PlyVA{S}(25)->[[CutStrByStartEnd]]{F}
- PlyVA{S}(27)->[[CutStrByStartEnd]]{F}
- PlyVA{S}(33)->[[ShellRunResult]]{F}
- PlyVA{S}(36)->[[CutStrByStartEnd]]{F}
- PlyVA{S}(39)->[[Json]]{F}
- PlyVA{S}(40)->[[Json]]{F}
- PlyVA{S}(41)->[[Json]]{F}
- PlyVA{S}(42)->[[Json]]{F}
- PlyVA{S}(45)->[[Json]]{F}

