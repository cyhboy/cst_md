&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub ScpDlParam(Hold As Boolean)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrorHandler`  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`Fold`](Fold.md)  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim path As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`'path = "cmd.exe /C " & `[`GetAppDrive`](GetAppDrive.md)`() & "\WinSCP\WinSCP.com /script=WinSCP.txt"`  
&nbsp;&nbsp;&nbsp;&nbsp;`path = "cmd.exe /C " & `[`GetAppDrive`](GetAppDrive.md)`() & "\WinSCP\WinSCP.com /command"`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim parameter As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fqdn As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`fqdn = Cells(currentRow, 2)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim uid As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`uid = Cells(currentRow, 3)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim vla As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`vla = Cells(currentRow, 3)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim pass As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`pass = Cells(currentRow, 4)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fileOrFolder As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`fileOrFolder = Cells(currentRow, 5)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim localFolder As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`localFolder = Cells(currentRow, 9)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fileSet As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`fileSet = Cells(currentRow, 11)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fileSetArr As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`fileSetArr = Split(fileSet, Chr(10))`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim port As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`port = Cells(currentRow, 7)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If (port = "") Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`port = "22"`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`If (Len(port) > 5) Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`port = "22"`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Not IsNumeric(port) Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`port = "22"`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim ppkPath As String: ppkPath = ""`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim ppkFile As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`ppkFile = Cells(currentRow, 14)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If True = `[`EndsWith`](EndsWith.md)`(ppkFile, ".ppk") Or ppkFile = "private_key" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Dim ppkFolder As String`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ppkFolder = Replace(Cells(currentRow, 13), ActiveWorkbook.Name, "")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ppkPath = ppkFolder & ppkFile`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`'Dim Length As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`'Length = Len(fileOrFolder)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim index As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`index = InStrRev(fileOrFolder, "/")`  
&nbsp;&nbsp;&nbsp;&nbsp;`fileOrFolder = Left(fileOrFolder, index)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If uid = "" And pass = "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`uid = Environ$("username")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Dim propsMap As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Set propsMap = `[`ReadPropertyInAppFiles`](ReadPropertyInAppFiles.md)`("identity.ini")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`pass = propsMap("AD_PASSWORD")`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`'binary|ascii|automatic`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If ppkPath <> "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`parameter = parameter & " " & """" & "open sftp://" & uid & "@" & fqdn & ":" & port & " -privatekey=" & ppkPath & """"`  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`parameter = parameter & " " & """" & "open sftp://" & uid & ":" & pass & "@" & fqdn & ":" & port & """"`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`'parameter = parameter & " " & """" & "call dzdo /bin/su - " & vla & """"`  
&nbsp;&nbsp;&nbsp;&nbsp;`parameter = parameter & " " & """" & "cd " & fileOrFolder & """"`  
&nbsp;&nbsp;&nbsp;&nbsp;`parameter = parameter & " " & """" & "lcd """"" & localFolder & """"""""`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim i As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`For i = 0 To UBound(fileSetArr)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(fileSetArr(i), ".xls") > 0 Or InStr(fileSetArr(i), ".rdb") > 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`parameter = parameter & " " & """" & "option transfer binary" & """"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`parameter = parameter & " " & """" & "option transfer ascii" & """"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`parameter = parameter & " " & """" & "get """"" & fileSetArr(i) & """"""""`  
&nbsp;&nbsp;&nbsp;&nbsp;`Next i`  
&nbsp;&nbsp;&nbsp;&nbsp;`parameter = parameter & " " & """" & "exit" & """"`  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox path & parameter`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim myresults As GlobalConfig.results`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`myresults = `[`ShellRunResult`](ShellRunResult.md)`(path & parameter, "C:\BAK\cmd.log", False, False, currentRow)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(myresults.rowNum, 20) = "'" & myresults.resultStr`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' ShellRun path & parameter, False`  
`ErrorHandler:`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox.md)` Err.Number & " " & Err.Description, 30`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`'    If Hold Then`  
`'        Dim exeName As String: exeName = ExtractEXE(path)`  
`'        MsgBox "exeName: " & exeName`  
`'`  
`'        While True = IsExeRunning(exeName)`  
`'            Sleep 5000`  
`'        Wend`  
`'    End If`  
`End Sub`  


# BeCaller
- ScpDlParam{S}(6)->[[Fold]]{S}
- ScpDlParam{S}(8)->[[GetAppDrive]]{F}
- ScpDlParam{S}(42)->[[EndsWith]]{F}
- ScpDlParam{S}(53)->[[ReadPropertyInAppFiles]]{F}
- ScpDlParam{S}(74)->[[ShellRunResult]]{F}
- ScpDlParam{S}(78)->[[MyMsgBox]]{S}

