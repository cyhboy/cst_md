&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub Proc2Fil()`
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
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox subName`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`RobotRunByParam`](RobotRunByParam)` "Proc2Fil"`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Next curCell`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim module As String`
&nbsp;&nbsp;&nbsp;&nbsp;`module = Cells(currentRow, 1)`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim subb As String`
&nbsp;&nbsp;&nbsp;&nbsp;`subb = Cells(currentRow, 2)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim i As Long`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim proc As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim VBComp As VBComponent`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim objProject As VBIDE.VBProject`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim objCode As VBIDE.CodeModule`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim codeOfLine As String`
&nbsp;&nbsp;&nbsp;&nbsp;`'Dim startOfProc As Long`
&nbsp;&nbsp;&nbsp;&nbsp;`'Dim lengthOfProc As Long`
&nbsp;&nbsp;&nbsp;&nbsp;`'startOfProc = objCode.ProcStartLine(proc, vbext_pk_Proc)`
&nbsp;&nbsp;&nbsp;&nbsp;`'lengthOfProc = objCode.ProcCountLines(proc, vbext_pk_Proc)`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim resultStr As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Set objProject = ThisWorkbook.VBProject`
&nbsp;&nbsp;&nbsp;&nbsp;`For Each VBComp In objProject.VBComponents`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If module = VBComp.Name Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' Find the code module for the project.`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Set objCode = VBComp.CodeModule`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`For i = 1 To objCode.CountOfLines`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`codeOfLine = objCode.Lines(i, 1)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'If Trim(codeOfLine) <> "" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`proc = objCode.ProcOfLine(i, vbext_pk_Proc)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If subb = proc Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`resultStr = resultStr & codeOfLine & Chr(13) & Chr(10)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'End If`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Next i`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`Next VBComp`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Dim folder As String`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'folder = "C:\SANDBOX\VB_SPACE\CST_PROJECT\" & Format(Now, "yyyyMMddHHmm") & "\" & module & "\"`
&nbsp;&nbsp;&nbsp;&nbsp;`'folder = "C:\SANDBOX\VB_SPACE\GIT_CST_PROJECT\" & Format(Now, "yyyyMMddHHmm") & "\" & module & "\"`
&nbsp;&nbsp;&nbsp;&nbsp;`folder = "C:\SANDBOX\VB_SPACE\GIT_CST_PROJECT\" & Format(Now, "yyyyMMdd") & "\" & module & "\"`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`CreateFolder`](CreateFolder)` folder`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`WriteTxt2Tmp`](WriteTxt2Tmp)` resultStr, folder & subb & ".vb"`
`ErrorHandler:`
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 5`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 47) = "###"`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
`End Sub`


> [!Getting information]
> Ribbon path please refer to ==**Engineer >> Project >> Proc2Fil**==


# BeCaller
- Proc2Fil{S}(16)->[[RobotRunByParam]]{S}
- Proc2Fil{S}(51)->[[SearchRegxKwInStr]]{F}
- Proc2Fil{S}(56)->[[CntSubstring]]{F}
- Proc2Fil{S}(60)->[[SearchRegxKwInStr]]{F}
- Proc2Fil{S}(65)->[[CntSubstring]]{F}
- Proc2Fil{S}(69)->[[SearchRegxKwInStr]]{F}
- Proc2Fil{S}(74)->[[CntSubstring]]{F}
- Proc2Fil{S}(92)->[[MatchRegx]]{F}
- Proc2Fil{S}(95)->[[MatchRegx]]{F}
- Proc2Fil{S}(96)->[[SearchRegxKwInStr]]{F}
- Proc2Fil{S}(98)->[[MatchRegx]]{F}
- Proc2Fil{S}(101)->[[MatchRegx]]{F}
- Proc2Fil{S}(102)->[[SearchRegxKwInStr]]{F}
- Proc2Fil{S}(107)->[[MatchRegx]]{F}
- Proc2Fil{S}(161)->[[MatchRegx]]{F}
- Proc2Fil{S}(162)->[[SearchRegxKwInStr]]{F}
- Proc2Fil{S}(164)->[[MatchRegx]]{F}
- Proc2Fil{S}(165)->[[SearchRegxKwInStr]]{F}
- Proc2Fil{S}(167)->[[MatchRegx]]{F}
- Proc2Fil{S}(174)->[[CreateFolder]]{S}
- Proc2Fil{S}(175)->[[WriteTxt2Tmp]]{S}
- Proc2Fil{S}(178)->[[MyMsgBox]]{S}

