&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub SftpDlParam(hold As Boolean)`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrorHandler`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim path As String`
&nbsp;&nbsp;&nbsp;&nbsp;`path = `[`GetAppDrive`](GetAppDrive)`() & "\FlashFXP\FlashFXP.exe "`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim parameter As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fqdn As String`
&nbsp;&nbsp;&nbsp;&nbsp;`fqdn = Cells(currentRow, 2)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim uid As String`
&nbsp;&nbsp;&nbsp;&nbsp;`uid = Cells(currentRow, 3)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim pass As String`
&nbsp;&nbsp;&nbsp;&nbsp;`pass = Cells(currentRow, 4)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fileOrFolder As String`
&nbsp;&nbsp;&nbsp;&nbsp;`fileOrFolder = Cells(currentRow, 5)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim localFolder As String`
&nbsp;&nbsp;&nbsp;&nbsp;`localFolder = Cells(currentRow, 9)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim port As String`
&nbsp;&nbsp;&nbsp;&nbsp;`port = Cells(currentRow, 7)`
&nbsp;&nbsp;&nbsp;&nbsp;`If (port = "") Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`port = "22"`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If (Len(port) > 5) Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`port = "22"`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If Not IsNumeric(port) Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`port = "22"`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim Length As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`Length = Len(fileOrFolder)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim index As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`index = InStrRev(fileOrFolder, "/")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`fileOrFolder = Left(fileOrFolder, index)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If pass = "" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'uid = Environ$("username")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Dim propsMap As Variant`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Set propsMap = `[`ReadPropertyInAppFiles`](ReadPropertyInAppFiles)`("identity.ini")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`pass = propsMap("AD_PASSWORD")`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If (InStr(Cells(currentRow, 11), "*") = 0 And InStr(Cells(currentRow, 11), "#") = 0) Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`parameter = " -download sftp://" & uid & ":" & pass & "@" & fqdn & ":" & port & " -remotepath=""" & fileOrFolder & Cells(currentRow, 11) & """ -local=" & """" & localFolder & """" & " -localpath=" & """" & localFolder & Replace(Cells(currentRow, 11), "/", "\") & """"`
&nbsp;&nbsp;&nbsp;&nbsp;`ElseIf True = `[`StartsWith`](StartsWith)`(Cells(currentRow, 11), "#") Then`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`parameter = " -download sftp://" & uid & ":" & pass & "@" & fqdn & ":" & port & " -remotepath=""" & fileOrFolder & """ -local=" & """" & localFolder & """" & " -localpath=" & """" & localFolder & """" & " -ruleset=" & """" & Replace(Cells(currentRow, 11).Value, "#", "") & """"`
&nbsp;&nbsp;&nbsp;&nbsp;`Else`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`parameter = " -download sftp://" & uid & ":" & pass & "@" & fqdn & ":" & port & " -remotepath=""" & fileOrFolder & """ -local=" & """" & localFolder & """" & " -localpath=" & """" & localFolder & """" & " -ruleset=" & """" & Cells(currentRow, 11) & """"`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox path & parameter`
&nbsp;&nbsp;&nbsp;&nbsp;`'Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`ShellRunMax`](ShellRunMax)` path & parameter`
&nbsp;&nbsp;&nbsp;&nbsp;
`ErrorHandler:`
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 30`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If hold Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Dim exeName As String: exeName = `[`ExtractEXE`](ExtractEXE)`(path)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`While True = `[`IsExeRunning`](IsExeRunning)`(exeName)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Sleep 5000`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Wend`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
`End Sub`


# BeCaller
- SftpDlParam{S}(7)->[[GetAppDrive]]{F}
- SftpDlParam{S}(39)->[[ReadPropertyInAppFiles]]{F}
- SftpDlParam{S}(44)->[[StartsWith]]{F}
- SftpDlParam{S}(49)->[[ShellRunMax]]{S}
- SftpDlParam{S}(52)->[[MyMsgBox]]{S}
- SftpDlParam{S}(55)->[[ExtractEXE]]{F}
- SftpDlParam{S}(56)->[[IsExeRunning]]{F}

