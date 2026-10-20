&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub PttyMD()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' On Error GoTo ErrorHandler`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'    Dim n As Integer`  
`'    n = Selection.count`  
`'    If n > 1 Then`  
`'        n = Selection.SpecialCells(xlCellTypeVisible).count`  
`'    End If`  
`'    If n > 1 Then`  
`'        Dim curCell As Range`  
`'        For Each curCell In Selection`  
`'            If curCell.EntireColumn.Hidden = False And curCell.EntireRow.Hidden = False Then`  
`'                curCell.Select`  
`'                'MsgBox subName`  
`'                RobotRun "PttyMD"`  
`'            End If`  
`'        Next curCell`  
`'        Exit Sub`  
`'    End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim path As String`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim parameter As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`path = `[`GetAppDrive`](GetAppDrive.md)`() & "\ptty\putty.exe "`  
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fqdn As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`fqdn = Cells(currentRow, 2)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim uid As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`uid = Cells(currentRow, 3)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim pass As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`pass = Cells(currentRow, 4)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim commandPath As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`commandPath = `[`GetBakDrive`](GetBakDrive.md)`() & "\ptty_command.txt"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim cmd As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`cmd = "mkdir -p " & Cells(currentRow, 5)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim port As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`port = Cells(currentRow, 7)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If port = "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`port = "22"`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;[`WriteTxt2Code`](WriteTxt2Code.md)` cmd & Chr(13) & Chr(10) & "exit", commandPath`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`parameter = fqdn & " -l " & uid & " -pw " & pass & " -P " & port & " -m """ & commandPath & """ -t"`  
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
&nbsp;&nbsp;&nbsp;&nbsp;`If ppkPath <> "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`parameter = fqdn & " -l " & uid & " -i """ & ppkPath & """ -P " & port & " -m """ & commandPath & """ -t"`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If pass = "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' WriteTxt2Code "dzdo /bin/su - " & uid & " -c '" & cmd & "'" & Chr(13) & Chr(10) & "exit", commandPath`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' uid = Environ$("username")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Dim propsMap As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Set propsMap = `[`ReadPropertyInAppFiles`](ReadPropertyInAppFiles.md)`("identity.ini")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' pass = propsMap("AD_PASSWORD")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' parameter = Cells(currentRow, 2) & " -l " & uid & " -pw " & pass & " -m """ & commandPath & """ -t"`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'    Dim cntEXE As Integer`  
`'    cntEXE = CntExeRunning(ExtractEXE(path))`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' ShellRunHide path & parameter`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim pyResult As GlobalConfig.results`  
&nbsp;&nbsp;&nbsp;&nbsp;`pyResult = `[`ShellRunResult`](ShellRunResult.md)`(path & parameter, "C:\BAK\cmd.log", True, False, currentRow)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim pttyResult As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`pttyResult = `[`SearchRegxKwInFile`](SearchRegxKwInFile.md)`("C:\BAK\putty.log", "(fail)")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If pttyResult = "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox.md)` "Server Folder Created Successfully", 15, False, True`  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox.md)` "Server Folder Fail To Create", 15, False, True`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'    Dim killFlag As Boolean`  
`'    While CntExeRunning(ExtractEXE(path)) - cntEXE > 0 And killFlag = False`  
`'        Sleep 3000`  
`'        If Now - LastModDate("C:\BAK\putty.log") > 3000 / 1000 / 60 / 24 Then`  
`'            'MyQuestionBox "Do U want to kill", "Yes", "No", "", 6`  
`'            'If confirmation = "Yes" Then`  
`'            'killFlag = KillExeRunning(ExtractEXE(path))`  
`'            'End If`  
`'        End If`  
`'    Wend`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'ErrorHandler:`  
`'    If Err.Number <> 0 Then`  
`'        MyMsgBox Err.Number & " " & Err.Description, 30`  
`'    End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Engineer >> Putty >> PttyOther >> PttyMD**==


# BeCaller
- PttyMD{S}(8)->[[GetAppDrive]]{F}
- PttyMD{S}(17)->[[GetBakDrive]]{F}
- PttyMD{S}(25)->[[WriteTxt2Code]]{S}
- PttyMD{S}(30)->[[EndsWith]]{F}
- PttyMD{S}(40)->[[ReadPropertyInAppFiles]]{F}
- PttyMD{S}(43)->[[ShellRunResult]]{F}
- PttyMD{S}(45)->[[SearchRegxKwInFile]]{F}
- PttyMD{S}(47)->[[MyMsgBox]]{S}
- PttyMD{S}(49)->[[MyMsgBox]]{S}

