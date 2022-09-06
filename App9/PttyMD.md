&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub PttyMD()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'    On Error GoTo LineHandler`
&nbsp;&nbsp;&nbsp;&nbsp;`'    Dim subName As String`
&nbsp;&nbsp;&nbsp;&nbsp;`'362:     Err.Raise 1979`
&nbsp;&nbsp;&nbsp;&nbsp;`'LineHandler:`
&nbsp;&nbsp;&nbsp;&nbsp;`'    If Err.Number = 1979 Then`
&nbsp;&nbsp;&nbsp;&nbsp;`'        subName = GetSubName("AllSpecial", Erl)`
&nbsp;&nbsp;&nbsp;&nbsp;`'        If subName = "" Then`
&nbsp;&nbsp;&nbsp;&nbsp;`'            MsgBox "Process name is not available. Please contact administrator. "`
&nbsp;&nbsp;&nbsp;&nbsp;`'            'Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`'        Else`
&nbsp;&nbsp;&nbsp;&nbsp;`'            LogToDB subName`
&nbsp;&nbsp;&nbsp;&nbsp;`'        End If`
&nbsp;&nbsp;&nbsp;&nbsp;`'    End If`
&nbsp;&nbsp;&nbsp;&nbsp;`'    Resume Next`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrorHandler`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim path As String`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim parameter As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`path = `[`GetAppDrive`](GetAppDrive)`() & "\ptty\putty.exe "`
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim uid As String`
&nbsp;&nbsp;&nbsp;&nbsp;`uid = Cells(currentRow, 3)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim pass As String`
&nbsp;&nbsp;&nbsp;&nbsp;`pass = Cells(currentRow, 4)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim commandPath As String`
&nbsp;&nbsp;&nbsp;&nbsp;`commandPath = `[`GetBakDrive`](GetBakDrive)`() & "\ptty_command.txt"`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim cmd As String`
&nbsp;&nbsp;&nbsp;&nbsp;`cmd = "mkdir -p " & Cells(currentRow, 5)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`WriteTxt2Tmp`](WriteTxt2Tmp)` cmd & Chr(13) & Chr(10) & "exit", commandPath`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`parameter = Cells(currentRow, 2) & " -l " & uid & " -pw " & pass & " -m """ & commandPath & """ -t"`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If pass = "" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`WriteTxt2Tmp`](WriteTxt2Tmp)` "dzdo /bin/su - " & uid & " -c '" & cmd & "'" & Chr(13) & Chr(10) & "exit", commandPath`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`uid = Environ$("username")`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Dim propsMap As Variant`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Set propsMap = `[`ReadPropertyInAppFiles`](ReadPropertyInAppFiles)`("identity.ini")`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`pass = propsMap("AD_PASSWORD")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`parameter = Cells(currentRow, 2) & " -l " & uid & " -pw " & pass & " -m """ & commandPath & """ -t"`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`ShellRunHide`](ShellRunHide)` path & parameter`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim exeName As String: exeName = `[`ExtractEXE`](ExtractEXE)`(path)`
&nbsp;&nbsp;&nbsp;&nbsp;`While True = `[`IsExeRunning`](IsExeRunning)`(exeName)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Sleep 5000`
&nbsp;&nbsp;&nbsp;&nbsp;`Wend`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim pttyResult As String`
&nbsp;&nbsp;&nbsp;&nbsp;`pttyResult = `[`SearchRegxKwInFile`](SearchRegxKwInFile)`("C:\BAK\putty.log", "(fail)")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If pttyResult = "" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`MsgBox "Server Folder Created Successfully"`
&nbsp;&nbsp;&nbsp;&nbsp;`Else`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`MsgBox "Server Folder Fail To Create"`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
`ErrorHandler:`
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 30`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'    If "On" = ReadRegAR() Then`
&nbsp;&nbsp;&nbsp;&nbsp;`'        Dim exer As String`
&nbsp;&nbsp;&nbsp;&nbsp;`'        exer = Cells(currentRow, 16)`
&nbsp;&nbsp;&nbsp;&nbsp;`'        If InStr(exer, subName) = 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;`'            Cells(currentRow, 16) = Trim(exer & " " & subName)`
&nbsp;&nbsp;&nbsp;&nbsp;`'        End If`
&nbsp;&nbsp;&nbsp;&nbsp;`'    End If`
`End Sub`


> [!Getting information]
> Ribbon path please refer to ==**Engineer >> Putty >> PttyOther >> PttyMD**==


# BeCaller
- PttyMD{S}(9)->[[GetAppDrive]]{F}
- PttyMD{S}(16)->[[GetBakDrive]]{F}
- PttyMD{S}(19)->[[WriteTxt2Tmp]]{S}
- PttyMD{S}(22)->[[WriteTxt2Tmp]]{S}
- PttyMD{S}(25)->[[ReadPropertyInAppFiles]]{F}
- PttyMD{S}(29)->[[ShellRunHide]]{S}
- PttyMD{S}(30)->[[ExtractEXE]]{F}
- PttyMD{S}(31)->[[IsExeRunning]]{F}
- PttyMD{S}(35)->[[SearchRegxKwInFile]]{F}
- PttyMD{S}(43)->[[MyMsgBox]]{S}

