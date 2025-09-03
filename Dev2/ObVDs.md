&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub ObVDs()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim parameter As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`parameter = Cells(currentRow, 10)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim Folder As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Folder = Cells(currentRow, 9)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If True = `[`EndsWith`](EndsWith)`(Folder, "\") Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Folder = Left(Folder, Len(Folder) - 1)`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`' folder = Replace(folder, "\", "\\")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim output_name As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`output_name = Cells(currentRow, 11)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'    If InStr(parameter, "http") > 0 Then`  
`'        parameter = CutStrByStartEnd(parameter, "http", "$$", True)`  
`'        While InStr(parameter, vbLf) > 0`  
`'            ' MsgBox "vbLf"`  
`'            parameter = CutStrByStartEnd(parameter, "http", vbLf, True)`  
`'        Wend`  
`'        While InStr(parameter, vbCr) > 0`  
`'            ' MsgBox "vbCr"`  
`'            parameter = CutStrByStartEnd(parameter, "http", vbCr, True)`  
`'        Wend`  
`'    Else`  
`'        parameter = ""`  
`'    End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox parameter`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Exit Sub`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim cmdStr As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`' cmdStr = "conda activate learn"`  
&nbsp;&nbsp;&nbsp;&nbsp;`' cmdStr = cmdStr & " && " & "python C:\AppFiles\ipy\plyVA.py """ & parameter & """"`  
&nbsp;&nbsp;&nbsp;&nbsp;`' after pyinstaller build the python file`  
&nbsp;&nbsp;&nbsp;&nbsp;`' cmdStr = "C:\AppFiles\ipy\obmp4\obmp4.exe """ & parameter & """" & " " & """" & folder & """" & " " & """" & output_name & """"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`cmdStr = "C:\AppFiles\ipy\obvdbyid\obvdbyid.exe" & " " & """" & Folder & """" & " " & """" & output_name & """"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox cmdStr`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' Exit Sub`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'    Dim logFile As String`  
`'    logFile = "C:\BAK\cmd.log"`  
`'    logFile = Left(logFile, Len(logFile) - 4) & "_" & LPad(CStr(currentRow), 4, "0") & Right(logFile, 4)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim pyResult As GlobalConfig.results`  
&nbsp;&nbsp;&nbsp;&nbsp;`pyResult = `[`ShellRunResult`](ShellRunResult)`(cmdStr, "C:\BAK\cmd.log", False, False, currentRow)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(pyResult.rowNum, 20) = pyResult.resultStr`  
`End Sub`  


# BeCaller
- ObVDs{S}(11)->[[EndsWith]]{F}
- ObVDs{S}(19)->[[ShellRunResult]]{F}

