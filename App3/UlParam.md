&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub UlParam(hold As Boolean)`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrorHandler`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`Fold`](Fold)
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim path As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim parameter As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'path = `[`GetAppDrive`](GetAppDrive)`() & "\ptty\pscp.exe -i C:\AppFiles\ptty\myPri.ppk"`
&nbsp;&nbsp;&nbsp;&nbsp;`path = `[`GetAppDrive`](GetAppDrive)`() & "\ptty\pscp.exe "`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim hn As String`
&nbsp;&nbsp;&nbsp;&nbsp;`hn = Cells(currentRow, 2)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim uid As String`
&nbsp;&nbsp;&nbsp;&nbsp;`uid = Cells(currentRow, 3)`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim pass As String`
&nbsp;&nbsp;&nbsp;&nbsp;`pass = Cells(currentRow, 4)`
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
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fileOrFolder As String`
&nbsp;&nbsp;&nbsp;&nbsp;`fileOrFolder = Cells(currentRow, 5)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim Length As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`Length = Len(fileOrFolder)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim index As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`index = InStrRev(fileOrFolder, "/")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`fileOrFolder = Left(fileOrFolder, index)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim specialFile As String`
&nbsp;&nbsp;&nbsp;&nbsp;`specialFile = Cells(currentRow, 11)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'fileOrFolder = fileOrFolder & specialFile`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim localFolder As String`
&nbsp;&nbsp;&nbsp;&nbsp;`localFolder = Cells(currentRow, 9)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If Not EndsWith(localFolder, "\") Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`localFolder = localFolder & "\"`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If pass = "" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'uid = Environ$("username")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Dim propsMap As Variant`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Set propsMap = `[`ReadPropertyInAppFiles`](ReadPropertyInAppFiles)`("identity.ini")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`pass = propsMap("AD_PASSWORD")`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`If pass = "" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`parameter = " -sftp -l " & uid & " -P " & port & " -p """ & localFolder & specialFile & """" & " """ & hn & ":" & fileOrFolder & """"`
&nbsp;&nbsp;&nbsp;&nbsp;`Else`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`parameter = " -sftp -l " & uid & " -pw " & pass & " -P " & port & " -p """ & localFolder & specialFile & """" & " """ & hn & ":" & fileOrFolder & """"`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
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
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Sleep 10000`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`ShellRun`](ShellRun)` path & "--close", False`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Wend`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
`End Sub`


# BeCaller
- UlParam{S}(6)->[[Fold]]{S}
- UlParam{S}(10)->[[GetAppDrive]]{F}
- UlParam{S}(45)->[[ReadPropertyInAppFiles]]{F}
- UlParam{S}(53)->[[ShellRunMax]]{S}
- UlParam{S}(56)->[[MyMsgBox]]{S}
- UlParam{S}(59)->[[ExtractEXE]]{F}
- UlParam{S}(60)->[[IsExeRunning]]{F}
- UlParam{S}(62)->[[ShellRun]]{S}

