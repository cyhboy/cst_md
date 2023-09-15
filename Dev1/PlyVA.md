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
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim resultStr As String`
&nbsp;&nbsp;&nbsp;&nbsp;`resultStr = Cells(currentRow, 10)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim formatCode As String`
&nbsp;&nbsp;&nbsp;&nbsp;`formatCode = `[`CutStrByStartEnd`](CutStrByStartEnd)`(parameter, " best", "http", True)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim origFile As String`
&nbsp;&nbsp;&nbsp;&nbsp;`origFile = `[`CutStrByStartEnd`](CutStrByStartEnd)`(parameter, "::ffmpeg -i """, """")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim audioFile As String`
&nbsp;&nbsp;&nbsp;&nbsp;`audioFile = `[`CutStrByStartEnd`](CutStrByStartEnd)`(parameter, " -acodec copy """, """")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(parameter, "http") > 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`parameter = `[`CutStrByStartEnd`](CutStrByStartEnd)`(parameter, "http", "$$", True)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`While InStr(parameter, vbLf) > 0`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox "vbLf"`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`parameter = `[`CutStrByStartEnd`](CutStrByStartEnd)`(parameter, "http", vbLf, True)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Wend`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`While InStr(parameter, vbCr) > 0`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox "vbCr"`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`parameter = `[`CutStrByStartEnd`](CutStrByStartEnd)`(parameter, "http", vbCr, True)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Wend`
&nbsp;&nbsp;&nbsp;&nbsp;`Else`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`parameter = ""`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox parameter`
&nbsp;&nbsp;&nbsp;&nbsp;`'Exit Sub`
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
&nbsp;&nbsp;&nbsp;&nbsp;`jsonStr = `[`CutStrByStartEnd`](CutStrByStartEnd)`(jsonStr, "{", "$$", True)`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim Json As Object`
&nbsp;&nbsp;&nbsp;&nbsp;`Set Json = JsonConverter.ParseJson(jsonStr)`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim cmdResult As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim subtitles As String`
&nbsp;&nbsp;&nbsp;&nbsp;`subtitles = `[`Json`](Json)`("subtitles")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 1) = subtitles`
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 2) = `[`Json`](Json)`("filesizeString")`
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 3) = `[`Json`](Json)`("view_count")`
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 4) = "'" & `[`Json`](Json)`("upload_date")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim rplFormatCode As String`
&nbsp;&nbsp;&nbsp;&nbsp;`rplFormatCode = " " & `[`Json`](Json)`("formatCode") & " "`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If Not (subtitles = "subtitles0[]" Or subtitles = "subtitles0" Or subtitles = "subtitlesErr" Or subtitles = "subtitlesNil") Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`rplFormatCode = rplFormatCode & "--write-sub --sub-lang en,en-US,en-GB,zh,zh-CN,zh-HK,zh-TW,zh-Hans,zh-Hant --convert-subs srt "`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`resultStr = Replace(resultStr, formatCode, rplFormatCode)`
&nbsp;&nbsp;&nbsp;&nbsp;`resultStr = Replace(resultStr, origFile, `[`Json`](Json)`("videoFileName"))`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim extStr As String`
&nbsp;&nbsp;&nbsp;&nbsp;`extStr = `[`Json`](Json)`("videoFileName")`
&nbsp;&nbsp;&nbsp;&nbsp;`extStr = Right(extStr, Len(extStr) - InStrRev(extStr, ".") + 1)`
&nbsp;&nbsp;&nbsp;&nbsp;`If extStr = "webm" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`resultStr = Replace(resultStr, audioFile, Replace(Json("videoFileName"), extStr, ".opus"))`
&nbsp;&nbsp;&nbsp;&nbsp;`Else`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`resultStr = Replace(resultStr, audioFile, Replace(Json("videoFileName"), extStr, ".wma"))`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 10) = resultStr`
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 8) = resultStr`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 13) = "'" & `[`Json`](Json)`("videoFileName")`
&nbsp;&nbsp;&nbsp;&nbsp;
`End Sub`


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Auto >> PlyVA**==


# BeCaller
- PlyVA{S}(15)->[[RobotRunByParam]]{S}
- PlyVA{S}(27)->[[CutStrByStartEnd]]{F}
- PlyVA{S}(29)->[[CutStrByStartEnd]]{F}
- PlyVA{S}(31)->[[CutStrByStartEnd]]{F}
- PlyVA{S}(33)->[[CutStrByStartEnd]]{F}
- PlyVA{S}(35)->[[CutStrByStartEnd]]{F}
- PlyVA{S}(38)->[[CutStrByStartEnd]]{F}
- PlyVA{S}(45)->[[ShellRunResult]]{F}
- PlyVA{S}(48)->[[CutStrByStartEnd]]{F}
- PlyVA{S}(53)->[[Json]]{F}
- PlyVA{S}(55)->[[Json]]{F}
- PlyVA{S}(56)->[[Json]]{F}
- PlyVA{S}(57)->[[Json]]{F}
- PlyVA{S}(59)->[[Json]]{F}
- PlyVA{S}(66)->[[Json]]{F}
- PlyVA{S}(75)->[[Json]]{F}

