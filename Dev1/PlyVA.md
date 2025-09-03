&nbsp;  &nbsp;  &nbsp;  &nbsp;  
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
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox subName`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`RobotRun`](RobotRun)` "PlyVA"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Next curCell`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' GoTo SECOND_STAGE`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim parameter As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`parameter = Cells(currentRow, 10)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim resultStr As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`resultStr = Cells(currentRow, 10)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim formatCode As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`formatCode = `[`CutStrByStartEnd`](CutStrByStartEnd)`(parameter, " best", "http", True)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim orgTitleText As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(parameter, " -metadata title=") > 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`orgTitleText = `[`CutStrByStartEnd`](CutStrByStartEnd)`(parameter, " -metadata title=", " -metadata album=", True, True)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`orgTitleText = " -metadata album="`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim origFile As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`origFile = `[`CutStrByStartEnd`](CutStrByStartEnd)`(parameter, "::ffmpeg -i """, """")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim audioFile As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`audioFile = `[`CutStrByStartEnd`](CutStrByStartEnd)`(parameter, " -acodec copy """, """")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(parameter, "http") > 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`parameter = `[`CutStrByStartEnd`](CutStrByStartEnd)`(parameter, "http", "$$", True)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`While InStr(parameter, vbLf) > 0`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox "vbLf"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`parameter = `[`CutStrByStartEnd`](CutStrByStartEnd)`(parameter, "http", vbLf, True)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Wend`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`While InStr(parameter, vbCr) > 0`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox "vbCr"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`parameter = `[`CutStrByStartEnd`](CutStrByStartEnd)`(parameter, "http", vbCr, True)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Wend`  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`parameter = ""`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim cmdStr As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`' cmdStr = "conda activate learn"`  
&nbsp;&nbsp;&nbsp;&nbsp;`' cmdStr = cmdStr & " && " & "python C:\AppFiles\ipy\plyVA.py """ & parameter & """"`  
&nbsp;&nbsp;&nbsp;&nbsp;`' after pyinstaller build the python file`  
&nbsp;&nbsp;&nbsp;&nbsp;`cmdStr = "C:\AppFiles\ipy\plyVA\plyVA.exe """ & parameter & """"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'    MsgBox cmdStr`  
`'    Exit Sub`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;[`MyQuestionBox`](MyQuestionBox)` "please select video type? ", "mp4", "", "webm", 3`  
&nbsp;&nbsp;&nbsp;&nbsp;`If confirmation = "mp4" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`cmdStr = cmdStr & " " & """" & "mp4" & """"`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If confirmation = "webm" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`cmdStr = cmdStr & " " & """" & "webm" & """"`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'    MyMsgBox cmdStr`  
`'    Exit Sub`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim pyResult As GlobalConfig.results`  
&nbsp;&nbsp;&nbsp;&nbsp;`pyResult = `[`ShellRunResult`](ShellRunResult)`(cmdStr, "C:\BAK\cmd.log", False, False, currentRow)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' MyMsgBox pyResultStr, 15`  
&nbsp;&nbsp;&nbsp;&nbsp;`' MyMsgBox ReadLineByFile("C:\BAK\cmd.log"), 15`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Cells(pyResult.rowNum, 18) = "'" & pyResult.resultStr`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim titleText As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim videoFileName As String`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim jsonStr As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`' jsonStr = Cells(pyResult.rowNum, 18)`  
&nbsp;&nbsp;&nbsp;&nbsp;`jsonStr = pyResult.resultStr`  
&nbsp;&nbsp;&nbsp;&nbsp;`jsonStr = `[`CutStrByStartEnd`](CutStrByStartEnd)`(jsonStr, "{", "$$", True)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim Json As Object`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set Json = JsonConverter.ParseJson(jsonStr)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim cmdResult As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim subtitles As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`subtitles = `[`Json`](Json)`("subtitles")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`titleText = Replace(Replace(Json("titleText"), "<<", "{"), ">>", "}")`  
&nbsp;&nbsp;&nbsp;&nbsp;`videoFileName = Replace(Replace(Json("videoFileName"), "<<", "{"), ">>", "}")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(pyResult.rowNum, 1) = subtitles`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Cells(pyResult.rowNum, 2) = `[`Json`](Json)`("filesizeString")`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Cells(pyResult.rowNum, 3) = `[`Json`](Json)`("view_count")`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Cells(pyResult.rowNum, 4) = "'" & `[`Json`](Json)`("upload_date")`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Cells(currentRow, 6) = titleText`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim rplTitleText As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`rplTitleText = " -metadata title=""" & titleText & """" & " -metadata album="`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim rplFormatCode As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`rplFormatCode = " " & `[`Json`](Json)`("formatCode") & " "`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If Not (subtitles = "subtitles0[]" Or subtitles = "subtitles0" Or subtitles = "subtitlesErr" Or subtitles = "subtitlesNil") Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`rplFormatCode = rplFormatCode & "--write-sub --sub-lang en,en-US,en-GB,zh,zh-CN,zh-HK,zh-TW,zh-Hans,zh-Hant --convert-subs srt "`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`resultStr = Replace(resultStr, formatCode, rplFormatCode)`  
&nbsp;&nbsp;&nbsp;&nbsp;`resultStr = Replace(resultStr, origFile, videoFileName)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`resultStr = Replace(resultStr, orgTitleText, rplTitleText)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim extStr As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`extStr = videoFileName`  
&nbsp;&nbsp;&nbsp;&nbsp;`extStr = Right(extStr, Len(extStr) - InStrRev(extStr, ".") + 1)`  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox extStr`  
&nbsp;&nbsp;&nbsp;&nbsp;`If extStr = ".webm" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox resultStr`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`resultStr = Replace(resultStr, audioFile, Replace(videoFileName, extStr, ".opus"))`  
&nbsp;&nbsp;&nbsp;&nbsp;`ElseIf extStr = ".mp4" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox resultStr`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`resultStr = Replace(resultStr, audioFile, Replace(videoFileName, extStr, ".m4a"))`  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox resultStr`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(pyResult.rowNum, 10) = resultStr`  
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(pyResult.rowNum, 8) = resultStr`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(pyResult.rowNum, 13) = "'" & videoFileName`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Cells(pyResult.rowNum, 17) = "'" & `[`Json`](Json)`("vid")`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Cells(pyResult.rowNum, 19) = "'" & `[`Json`](Json)`("quality")`  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Auto >> PlyVA >> PlyVA**==


# BeCaller
- PlyVA{S}(15)->[[RobotRun]]{S}
- PlyVA{S}(27)->[[CutStrByStartEnd]]{F}
- PlyVA{S}(30)->[[CutStrByStartEnd]]{F}
- PlyVA{S}(35)->[[CutStrByStartEnd]]{F}
- PlyVA{S}(37)->[[CutStrByStartEnd]]{F}
- PlyVA{S}(39)->[[CutStrByStartEnd]]{F}
- PlyVA{S}(41)->[[CutStrByStartEnd]]{F}
- PlyVA{S}(44)->[[CutStrByStartEnd]]{F}
- PlyVA{S}(51)->[[MyQuestionBox]]{S}
- PlyVA{S}(59)->[[ShellRunResult]]{F}
- PlyVA{S}(64)->[[CutStrByStartEnd]]{F}
- PlyVA{S}(69)->[[Json]]{F}
- PlyVA{S}(76)->[[Json]]{F}

