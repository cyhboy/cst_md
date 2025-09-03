&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub PlyNTs()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'    Dim parameter As String`  
`'    parameter = Cells(currentRow, 10)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'    Dim noteId As String`  
`'    noteId = Cells(currentRow, 15)`  
`'    If Cells(currentRow, 11) <> noteId Then`  
`'        Cells(currentRow, 11) = "'" & noteId`  
`'    End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If Cells(currentRow, 11) <> Cells(currentRow, 13) Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 11) = Cells(currentRow, 13)`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim subFldr As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`subFldr = Cells(currentRow, 11)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If True = `[`EndsWith`](EndsWith)`(subFldr, ".mp4") Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`subFldr = Left(subFldr, Len(subFldr) - 4)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 11) = subFldr`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fldr As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`fldr = Cells(currentRow, 9) & Cells(currentRow, 11)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If True = `[`EndsWith`](EndsWith)`(fldr, "\") Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`fldr = Left(fldr, Len(fldr) - 1)`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`'Exit Sub`  
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
&nbsp;&nbsp;&nbsp;&nbsp;`Dim cmdStr As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`cmdStr = "C:\AppFiles\ipy\notebyid\notebyid.exe" & " " & """" & fldr & """"`  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox cmdStr`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`' ShellRunStd cmdStr`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim pyResult As GlobalConfig.results`  
&nbsp;&nbsp;&nbsp;&nbsp;`pyResult = `[`ShellRunResult`](ShellRunResult)`(cmdStr, "C:\BAK\cmd.log", False, False, currentRow)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(pyResult.rowNum, 20) = pyResult.resultStr`  
`End Sub`  


# BeCaller
- PlyNTs{S}(12)->[[EndsWith]]{F}
- PlyNTs{S}(18)->[[EndsWith]]{F}
- PlyNTs{S}(24)->[[ShellRunResult]]{F}

