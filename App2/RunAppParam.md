&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub RunAppParam(isHold As Boolean, isTest As Boolean, isKeep As Boolean)`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim parameter As String`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`parameter = Cells(currentRow, 10)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim arr As Variant`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`arr = Split(parameter, Chr(10))`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim path As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim i As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`For i = 0 To UBound(arr)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If Not (StartsWith(Trim(arr(i)), "::") Or StartsWith(UCase(Trim(arr(i))), "REM")) Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`path = path & arr(i) & "&"`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`Next i`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`While Right(path, 1) = "&"`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`path = Left(path, Len(path) - 1)`
&nbsp;&nbsp;&nbsp;&nbsp;`Wend`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If Not Cells(currentRow, 9).HasFormula Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If Dir(Cells(currentRow, 9), vbDirectory) <> vbNullString Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Dim cdPath As String`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`cdPath = Cells(currentRow, 9)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`path = "cd " & cdPath & "&" & path`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If isTest Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Dim propsMap As Variant`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Set propsMap = `[`ReadPropertyInAppFiles`](ReadPropertyInAppFiles)`("identity.ini")`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`path = "PsExec -u GUANGZHOUTEST\" & Environ$("username") & " -p " & propsMap("AD_PASSWORD_UAT") & " " & path`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox path`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`ShellRunHide`](ShellRunHide)` path`
&nbsp;&nbsp;&nbsp;&nbsp;`Else`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`ShellRun`](ShellRun)` path, isKeep`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If isHold Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Dim exeName As String: exeName = `[`ExtractEXE`](ExtractEXE)`(path)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`While True = `[`IsExeRunning`](IsExeRunning)`(exeName)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Sleep 5000`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Wend`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
`End Sub`


# BeCaller
- RunAppParam{S}(30)->[[ReadPropertyInAppFiles]]{F}
- RunAppParam{S}(33)->[[ShellRunHide]]{S}
- RunAppParam{S}(35)->[[ShellRun]]{S}
- RunAppParam{S}(38)->[[ExtractEXE]]{F}
- RunAppParam{S}(39)->[[IsExeRunning]]{F}

