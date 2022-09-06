&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub PttyHOST()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrorHandler`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim path As String`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim parameter As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`path = `[`GetAppDrive`](GetAppDrive)`() & "\ptty\putty.exe "`
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim uid As String`
&nbsp;&nbsp;&nbsp;&nbsp;`uid = Cells(currentRow, 3)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim pass As String`
&nbsp;&nbsp;&nbsp;&nbsp;`pass = Cells(currentRow, 4)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fqdn As String`
&nbsp;&nbsp;&nbsp;&nbsp;`fqdn = Cells(currentRow, 2)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim commandPath As String`
&nbsp;&nbsp;&nbsp;&nbsp;`commandPath = `[`GetBakDrive`](GetBakDrive)`() & "\ptty_command.txt"`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim cmd As String`
&nbsp;&nbsp;&nbsp;&nbsp;`cmd = "hostname -s" & Chr(13) & Chr(10) & "hostname -a" & Chr(13) & Chr(10) & "hostname -i" & Chr(13) & Chr(10) & "hostname -A" & Chr(13) & Chr(10) & "hostname -I"`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`WriteTxt2Tmp`](WriteTxt2Tmp)` cmd & Chr(13) & Chr(10) & "exit", commandPath`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`parameter = fqdn & " -l " & uid & " -pw " & pass & " -m """ & commandPath & """ -t"`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If pass = "" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`WriteTxt2Tmp`](WriteTxt2Tmp)` "dzdo /bin/su - " & uid & " -c '" & cmd & "'" & Chr(13) & Chr(10) & "exit", commandPath`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`uid = Environ$("username")`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Dim propsMap As Variant`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Set propsMap = `[`ReadPropertyInAppFiles`](ReadPropertyInAppFiles)`("identity.ini")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`pass = propsMap("AD_PASSWORD")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`parameter = fqdn & " -l " & uid & " -pw " & pass & " -m """ & commandPath & """ -t"`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim cntEXE As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`cntEXE = `[`CntExeRunning`](CntExeRunning)`(ExtractEXE(path))`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`ShellRunHide`](ShellRunHide)` path & parameter`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Sleep 1000`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim killFlag As Boolean`
&nbsp;&nbsp;&nbsp;&nbsp;`While `[`CntExeRunning`](CntExeRunning)`(ExtractEXE(path)) - cntEXE > 0 And killFlag = False`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Sleep 3000`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If Now - LastModDate("C:\BAK\putty.log") > 3000 / 1000 / 60 / 24 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'MyQuestionBox "Do U want to kill", "Yes", "No", 6`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'If confirmation = "Yes" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'killFlag = KillExeRunning(ExtractEXE(path))`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'End If`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`Wend`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'Sleep 1000`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'    Dim exeName As String: exeName = ExtractEXE(path)`
&nbsp;&nbsp;&nbsp;&nbsp;`'    While True = IsExeRunning(exeName)`
&nbsp;&nbsp;&nbsp;&nbsp;`'        Sleep 5000`
&nbsp;&nbsp;&nbsp;&nbsp;`'    Wend`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim pttyResult1 As String`
&nbsp;&nbsp;&nbsp;&nbsp;`pttyResult1 = `[`SearchRegxKwInFile`](SearchRegxKwInFile)`("C:\BAK\putty.log", "Using username .*\r\n(.*)\r\n", True)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'Dim pttyResult1 As String`
&nbsp;&nbsp;&nbsp;&nbsp;`'pttyResult1 = `[`SearchRegxKwInFile`](SearchRegxKwInFile)`("C:\BAK\putty.log", "(hk[^\.]*|tkcs[^\.]*|mtcs[^\.]*)$")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'Dim pttyResult2 As String`
&nbsp;&nbsp;&nbsp;&nbsp;`'pttyResult2 = searchRegxKwInFile("C:\BAK\putty.log", "(130[^ ]*)$")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'Dim pttyResult3 As String`
&nbsp;&nbsp;&nbsp;&nbsp;`'pttyResult3 = `[`SearchRegxKwInFile`](SearchRegxKwInFile)`("C:\BAK\putty.log", "(hk.*) $")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'Dim pttyResult4 As String`
&nbsp;&nbsp;&nbsp;&nbsp;`'pttyResult4 = `[`SearchRegxKwInFile`](SearchRegxKwInFile)`("C:\BAK\putty.log", "(130.*) $")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'Dim arr1`
&nbsp;&nbsp;&nbsp;&nbsp;`'Dim arr2`
&nbsp;&nbsp;&nbsp;&nbsp;`'arr1 = Split(pttyResult3, " ")`
&nbsp;&nbsp;&nbsp;&nbsp;`'arr2 = Split(pttyResult4, " ")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'    If pttyResult <> "" Then`
&nbsp;&nbsp;&nbsp;&nbsp;`'        pttyResult = "WebSphere MQ " & pttyResult`
&nbsp;&nbsp;&nbsp;&nbsp;`'    End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 1) = pttyResult1`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`' On Error Resume Next`
&nbsp;&nbsp;&nbsp;&nbsp;`' Cells(currentRow, 6) = arr2(Application.Match(fqdn, arr1, False) - 1)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'If Err.Number <> 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;`'    Cells(currentRow, 6) = arr2(0)`
&nbsp;&nbsp;&nbsp;&nbsp;`'End If`
`ErrorHandler:`
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 30`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
`End Sub`


> [!Getting information]
> Ribbon path please refer to ==**Engineer >> Putty >> PttyConfig >> PttyHOST**==


# BeCaller
- PttyHOST{S}(9)->[[GetAppDrive]]{F}
- PttyHOST{S}(18)->[[GetBakDrive]]{F}
- PttyHOST{S}(21)->[[WriteTxt2Tmp]]{S}
- PttyHOST{S}(24)->[[WriteTxt2Tmp]]{S}
- PttyHOST{S}(27)->[[ReadPropertyInAppFiles]]{F}
- PttyHOST{S}(32)->[[CntExeRunning]]{F}
- PttyHOST{S}(33)->[[ShellRunHide]]{S}
- PttyHOST{S}(42)->[[SearchRegxKwInFile]]{F}
- PttyHOST{S}(46)->[[MyMsgBox]]{S}

