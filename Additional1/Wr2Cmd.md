&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub Wr2Cmd()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim specialFile As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`specialFile = ActiveWorkbook.Name`  
&nbsp;&nbsp;&nbsp;&nbsp;`specialFile = Left(specialFile, InStrRev(specialFile, ".") - 1) & ".cmd"`  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox UnicodeToByteArray(specialFile)`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim resultStr As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`resultStr = ActiveWorkbook.FullName`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Exit Sub`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' On Error GoTo ErrorHandler`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim ind As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`ind = Cells(currentRow, 1)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim localFolder As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`localFolder = Cells(currentRow, 9)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If localFolder = "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim localFolderArr As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`localFolderArr = Split(localFolder, "\")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If localFolderArr(3) = "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If localFolderArr(1) = "Songs" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' specialFile = localFolderArr(2) & "_" & "Songs" & ".cmd"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' resultStr = "C:\CHROME_SPACE\hcs\learning\" & localFolderArr(2) & "_" & "Songs" & ".xlsm"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' specialFile = localFolderArr(2) & "_" & "videos" & ".cmd"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' resultStr = "C:\CHROME_SPACE\hcs\learning\" & localFolderArr(2) & "_" & "videos" & ".xlsm"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If ind = "None" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' specialFile = localFolderArr(2) & "_" & "playlists" & ".cmd"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`localFolder = Replace(localFolder, localFolderArr(3), "")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`localFolder = Replace(localFolder, "\\", "\")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' resultStr = "C:\CHROME_SPACE\hcs\learning\" & localFolderArr(2) & "_" & "playlists" & ".xlsm"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ElseIf InStr(ind, "_") > 0 And InStr(ind, "subtitles") <= 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' specialFile = ind & ".cmd"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' resultStr = "C:\CHROME_SPACE\hcs\learning\" & ind & ".xlsm"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' specialFile = localFolderArr(2) & "_" & localFolderArr(3) & ".cmd"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' resultStr = "C:\CHROME_SPACE\hcs\learning\" & localFolderArr(2) & "_" & localFolderArr(3) & ".xlsm"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim filePath As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`filePath = localFolder & specialFile`  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox filePath`  
&nbsp;&nbsp;&nbsp;&nbsp;[`UnhideFile`](UnhideFile)` filePath`  
&nbsp;&nbsp;&nbsp;&nbsp;[`WriteTxt2Code`](WriteTxt2Code)` resultStr, filePath`  
&nbsp;&nbsp;&nbsp;&nbsp;[`HideFile`](HideFile)` filePath`  
`'ErrorHandler:`  
`'    If Err.Number <> 0 Then`  
`'        MyMsgBox Err.Number & " " & Err.Description, 30`  
`'    End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`End Sub`  


# BeCaller
- Wr2Cmd{S}(35)->[[UnhideFile]]{S}
- Wr2Cmd{S}(36)->[[WriteTxt2Code]]{S}
- Wr2Cmd{S}(37)->[[HideFile]]{S}

