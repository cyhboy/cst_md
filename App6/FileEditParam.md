&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub FileEditParam(hold As Boolean, isFilter As Boolean, path As String)`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim parameter As String`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'path = """" & GetAppDrive() & "\EditPlus\editplus.exe" & """ -e"`
&nbsp;&nbsp;&nbsp;&nbsp;`'path = """" & "C:\Program Files\IDM Computer Solutions\UltraEdit\Uedit32.exe" & """ "`
&nbsp;&nbsp;&nbsp;&nbsp;`'path = "C:\AppFiles\Microsoft VS Code\Code.exe"`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'path = "C:\AppFiles\SublimeText\sublime_text.exe"`
&nbsp;&nbsp;&nbsp;&nbsp;`'path = "C:\AppFiles\Notepad++\notepad++.exe"`
&nbsp;&nbsp;&nbsp;&nbsp;`'path = "C:\Program Files\Microsoft VS Code\bin\code.cmd"`
&nbsp;&nbsp;&nbsp;&nbsp;`'path = "code.cmd"`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`
&nbsp;&nbsp;&nbsp;&nbsp;`'parameter = " " & """" & Replace(cell.Value, Chr(10), """ """) & """"`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fileStr As String`
&nbsp;&nbsp;&nbsp;&nbsp;`If Not isFilter Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`fileStr = Cells(currentRow, 11)`
&nbsp;&nbsp;&nbsp;&nbsp;`Else`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`fileStr = "*"`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fileArr As Variant`
&nbsp;&nbsp;&nbsp;&nbsp;`fileArr = Split(fileStr, Chr(10))`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim i As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`For i = 0 To UBound(fileArr)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`parameter = " " & """" & Cells(currentRow, 9) & fileArr(i) & """"`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox path & parameter`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' ShellRunHide path & parameter`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`ShellRun`](ShellRun)` path & parameter, False`
&nbsp;&nbsp;&nbsp;&nbsp;`Next i`
&nbsp;&nbsp;&nbsp;&nbsp;`If hold Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Dim exeName As String: exeName = `[`ExtractEXE`](ExtractEXE)`(path)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`While True = `[`IsExeRunning`](IsExeRunning)`(exeName)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Sleep 5000`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Wend`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
`End Sub`


# BeCaller
- FileEditParam{S}(19)->[[ShellRun]]{S}
- FileEditParam{S}(22)->[[ExtractEXE]]{F}
- FileEditParam{S}(23)->[[IsExeRunning]]{F}

