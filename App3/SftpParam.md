&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub SftpParam(Hold As Boolean)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrorHandler`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`Fold`](Fold)  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim path As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`path = `[`GetAppDrive`](GetAppDrive)`() & "\FlashFXP\FlashFXP.exe "`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim parameter As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fqdn As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`fqdn = Cells(currentRow, 2)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim uid As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`uid = Cells(currentRow, 3)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim ruid As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`ruid = uid`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim pass As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`pass = Cells(currentRow, 4)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fileOrFolder As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`fileOrFolder = Cells(currentRow, 5)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim localFolder As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`localFolder = Cells(currentRow, 9)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim port As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`port = Cells(currentRow, 7)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If (port = "") Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`port = "22"`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If (Len(port) > 5) Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`port = "22"`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If Not IsNumeric(port) Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`port = "22"`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim Length As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`Length = Len(fileOrFolder)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim index As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`index = InStrRev(fileOrFolder, "/")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`fileOrFolder = Left(fileOrFolder, index)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If Trim(pass) = "" Or Trim(uid) = "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'ruid = Environ$("username")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Dim propsMap As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Set propsMap = `[`ReadPropertyInAppFiles`](ReadPropertyInAppFiles)`("identity.ini")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`pass = propsMap("AD_PASSWORD")`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`parameter = "sftp://" & ruid & ":" & pass & "@" & fqdn & ":" & port & fileOrFolder & " -local=" & """" & localFolder & """"`  
&nbsp;&nbsp;&nbsp;&nbsp;`' parameter = "sftp://" & uid & ":" & pass & "@" & fqdn & fileOrFolder & " -local=" & """" & localFolder & """"`  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox path & parameter`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`' ShellRunMax path & parameter`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If Hold Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Dim pyResult As GlobalConfig.results`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`pyResult = `[`ShellRunResult`](ShellRunResult)`(path & parameter, "C:\BAK\cmd.log", False, False, currentRow)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(pyResult.rowNum, 20) = pyResult.resultStr`  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`ShellRunMax`](ShellRunMax)` path & parameter`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' ShellRunStd path & parameter`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`ErrorHandler:`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 30`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'    If Hold Then`  
`'        Dim exeName As String: exeName = ExtractEXE(path)`  
`'        While True = IsExeRunning(exeName)`  
`'            Sleep 5000`  
`'        Wend`  
`'    End If`  
`End Sub`  


# BeCaller
- SftpParam{S}(6)->[[Fold]]{S}
- SftpParam{S}(8)->[[GetAppDrive]]{F}
- SftpParam{S}(42)->[[ReadPropertyInAppFiles]]{F}
- SftpParam{S}(48)->[[ShellRunResult]]{F}
- SftpParam{S}(51)->[[ShellRunMax]]{S}
- SftpParam{S}(55)->[[MyMsgBox]]{S}

