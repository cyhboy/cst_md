&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub TagProc()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrorHandler`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim n As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`n = Selection.count`
&nbsp;&nbsp;&nbsp;&nbsp;`If n > 1 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`n = Selection.SpecialCells(xlCellTypeVisible).count`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`If n > 1 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Dim curCell As Range`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`For Each curCell In Selection`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If curCell.EntireColumn.Hidden = False And curCell.EntireRow.Hidden = False Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`curCell.Select`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`RobotRunByParam`](RobotRunByParam)` "TagProc"`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Next curCell`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim module As String`
&nbsp;&nbsp;&nbsp;&nbsp;`module = Cells(currentRow, 1)`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim subb As String`
&nbsp;&nbsp;&nbsp;&nbsp;`subb = Cells(currentRow, 2)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim folder As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim resultStr As String`
&nbsp;&nbsp;&nbsp;&nbsp;`folder = "C:\SANDBOX\VB_SPACE\GIT_CST_PROJECT\" & Format(Now, "yyyyMMdd") & "\" & module & "\"`
&nbsp;&nbsp;&nbsp;&nbsp;`resultStr = `[`ReadLineByFile`](ReadLineByFile)`(folder & subb & ".bas")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim funcStr As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim subStr As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim otherStr As String`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If `[`MatchRegx`](MatchRegx)`(resultStr, "^Public Sub ") Or `[`MatchRegx`](MatchRegx)`(resultStr, "^Public Function ") Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If `[`MatchRegx`](MatchRegx)`(resultStr, "^ *If testing Then") Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 17) = "TESTING"`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Else`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 17) = "TESTER"`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`Else`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 17) = "EXEMPT"`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If `[`MatchRegx`](MatchRegx)`(resultStr, "^ *Shell ") Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 18) = "Shell"`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If `[`MatchRegx`](MatchRegx)`(resultStr, "^ *Set objshell = CreateObject\(""Wscript.Shell""\)") Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 18) = Cells(currentRow, 18) & "Wscript.Shell"`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If `[`MatchRegx`](MatchRegx)`(resultStr, "^P.* Function [^\(]+\(") Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`funcStr = `[`SearchRegxKwInStr`](SearchRegxKwInStr)`(resultStr, "^(P[^ ]+ Function [^\(]+\(.*\).*)", True)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 19) = funcStr`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(funcStr, "()") > 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 20) = 0`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Else`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 20) = `[`CntSubstring`](CntSubstring)`(funcStr, ", ") + 1`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If `[`MatchRegx`](MatchRegx)`(resultStr, "^P.* Sub [^\(]+\(") Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`subStr = `[`SearchRegxKwInStr`](SearchRegxKwInStr)`(resultStr, "^(P[^ ]+ Sub [^\(]+\(.*\).*)", True)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 19) = subStr`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(subStr, "()") > 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 20) = 0`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Else`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 20) = `[`CntSubstring`](CntSubstring)`(subStr, ", ") + 1`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If `[`MatchRegx`](MatchRegx)`(resultStr, "^P.* Property Get [^\(]+\(") Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`otherStr = `[`SearchRegxKwInStr`](SearchRegxKwInStr)`(resultStr, "^(P[^ ]+ Property Get [^\(]+\(.*\).*)", True)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 19) = otherStr`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(otherStr, "()") > 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 20) = 0`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Else`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 20) = `[`CntSubstring`](CntSubstring)`(otherStr, ", ") + 1`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
`'    If `[`MatchRegx`](MatchRegx)`(resultStr, "^ *On Error GoTo ErrorHandler") Then`
`'        Cells(currentRow, 21) = "ErrorHandler"`
`'    ElseIf `[`MatchRegx`](MatchRegx)`(resultStr, "^ *On Error Resume Next") Then`
`'        Cells(currentRow, 21) = "ErrorResume"`
`'    Else`
`'        Cells(currentRow, 21) = "ErrorUncapture"`
`'    End If`
&nbsp;&nbsp;&nbsp;&nbsp;
`'    If `[`MatchRegx`](MatchRegx)`(resultStr, "^ *On Error Resume Next") Then`
`'        Cells(currentRow, 22) = "ErrorResume"`
`'    Else`
`'        Cells(currentRow, 22) = "ErrorThrow"`
`'    End If`
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 22) = ""`
&nbsp;&nbsp;&nbsp;&nbsp;
`'    If `[`MatchRegx`](MatchRegx)`(resultStr, "^ *On Error GoTo LineHandler") Then`
`'        Cells(currentRow, 23) = "SoftCode"`
`'    Else`
`'        Cells(currentRow, 23) = "HardCode"`
`'    End If`
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 23) = ""`
&nbsp;&nbsp;&nbsp;&nbsp;
`'    If `[`MatchRegx`](MatchRegx)`(resultStr, "^ *MsgBox ") Then`
`'        Cells(currentRow, 28) = "MsgBox Alert"`
`'    End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 28) = ""`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`TagProcRun`](TagProcRun)` resultStr, "^ *(On Error .*)", True, True, 21`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`TagProcRun`](TagProcRun)` resultStr, "^ *n = Selection.SpecialCells\(xlCellTypeVisible\)\.count", True, False, 24`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`TagProcRun`](TagProcRun)` resultStr, "^ *(Set .* = CreateObject\(""Scripting.FileSystemObject""\))", True, True, 25`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`TagProcRun`](TagProcRun)` resultStr, "^ *(Set objWMI = GetObject.*)", True, True, 26`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`TagProcRun`](TagProcRun)` resultStr, "^ *(cn.Open .*)", True, True, 27`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`TagProcRun`](TagProcRun)` resultStr, "^ *[^ ]+ = MyQuestionBox\([^,\r]+\)", True, True, 29`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`TagProcRun`](TagProcRun)` resultStr, "^ *Set fso = Nothing", True, False, 30`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`TagProcRun`](TagProcRun)` resultStr, "^ *MsgBox ""Please setup repository database. """, True, False, 31`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`TagProcRun`](TagProcRun)` resultStr, "[ \(]ActiveWorkbook.FullName", True, False, 32`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`TagProcRun`](TagProcRun)` resultStr, "[\.]Application.Cells.Find", True, False, 33`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`TagProcRun`](TagProcRun)` resultStr, "(SearchRegxKwInStrMultToList\([^,\r]+, [^,\r]+, [^,\r]+, [^,\r\)]+\))", True, True, 36`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`TagProcRun`](TagProcRun)` resultStr, "(SearchRegxKwInStr\([^,\r]+, [^,\r]+\))", True, True, 37`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`TagProcRun`](TagProcRun)` resultStr, "(SearchRegxKwInFileMultToList\([^,\r]+, [^,\r]+, [^,\r\)]+\))", True, True, 38`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`TagProcRun`](TagProcRun)` resultStr, "(SearchRegxKwInStrToList\([^,\r]+, [^,\r]+\))", True, True, 39`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`TagProcRun`](TagProcRun)` resultStr, "(SearchRegxKwInFile\([^,\r]+, [^,\r]+\))", True, True, 40`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`TagProcRun`](TagProcRun)` resultStr, "\\([^ ""\.\\]+\.vbs)", True, True, 41`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`TagProcRun`](TagProcRun)` resultStr, "\\([^ ""\.\\]+\.jar)", True, True, 42`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`TagProcRun`](TagProcRun)` resultStr, "\\([^ ""\.\\]+\.exe)", True, True, 43`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`TagProcRun`](TagProcRun)` resultStr, "\\([^ ""\.\\]+\.ps1)", True, True, 44`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`TagProcRun`](TagProcRun)` resultStr, "^ *(Set .* = CreateObject\(""Shell.Application""\).*)", True, False, 45`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`TagProcRun`](TagProcRun)` resultStr, "^ *(Set .* = CreateObject\(""InternetExplorer.Application""\).*)", True, False, 46`
&nbsp;&nbsp;&nbsp;&nbsp;
`ErrorHandler:`
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 47) = Err.Description`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
`End Sub`


# BeCaller
- TagProc{S}(16)->[[RobotRunByParam]]{S}
- TagProc{S}(30)->[[ReadLineByFile]]{F}
- TagProc{S}(46)->[[MatchRegx]]{F}
- TagProc{S}(50)->[[SearchRegxKwInStr]]{F}
- TagProc{S}(55)->[[CntSubstring]]{F}
- TagProc{S}(59)->[[SearchRegxKwInStr]]{F}
- TagProc{S}(64)->[[CntSubstring]]{F}
- TagProc{S}(68)->[[SearchRegxKwInStr]]{F}
- TagProc{S}(73)->[[CntSubstring]]{F}
- TagProc{S}(79)->[[TagProcRun]]{S}
- TagProc{S}(83)->[[TagProcRun]]{S}
- TagProc{S}(86)->[[TagProcRun]]{S}
- TagProc{S}(87)->[[TagProcRun]]{S}
- TagProc{S}(88)->[[TagProcRun]]{S}
- TagProc{S}(94)->[[TagProcRun]]{S}
- TagProc{S}(95)->[[TagProcRun]]{S}
- TagProc{S}(96)->[[TagProcRun]]{S}
- TagProc{S}(97)->[[TagProcRun]]{S}

