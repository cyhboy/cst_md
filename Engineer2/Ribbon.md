&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub Ribbon()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`UnHF`](UnHF)
&nbsp;&nbsp;&nbsp;&nbsp;`'Call Frz`
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`CleanRgn`](CleanRgn)
&nbsp;&nbsp;&nbsp;&nbsp;`Application.ScreenUpdating = False`
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrorHandler`
&nbsp;&nbsp;&nbsp;&nbsp;`'Dim VBAEditor As VBIDE.VBE`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim objProject As VBIDE.VBProject`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim objComponent As VBIDE.VBComponent`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim objCode As VBIDE.CodeModule`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`' Declare other miscellaneous variables.`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim iLine As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim sProcName As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim pk As VBIDE.vbext_ProcKind`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim i As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`i = 2`
&nbsp;&nbsp;&nbsp;&nbsp;`'Set VBAEditor = Application.VBE`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim codeOfLine As String`
&nbsp;&nbsp;&nbsp;&nbsp;`'Get the project details in the workbook.`
&nbsp;&nbsp;&nbsp;&nbsp;`Set objProject = ThisWorkbook.VBProject`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'Iterate through each component in the project.`
&nbsp;&nbsp;&nbsp;&nbsp;`For Each objComponent In objProject.VBComponents`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'Find the code module for the project.`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Set objCode = objComponent.CodeModule`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'Scan through the code module, looking for procedures.`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`iLine = 1`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Do While iLine < objCode.CountOfLines`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`codeOfLine = objCode.Lines(iLine, 1)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If Trim(codeOfLine) <> "" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`sProcName = objCode.ProcOfLine(iLine, pk)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If sProcName <> "" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox objComponent.Name & ": " & sProcName`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ActiveSheet.Cells(i, 1).Value = objComponent.Name`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ActiveSheet.Cells(i, 2).Value = sProcName`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ActiveSheet.Cells(i, 3).Value = iLine`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`iLine = iLine + objCode.ProcCountLines(sProcName, pk) - 2`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ActiveSheet.Cells(i, 11).Value = iLine`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'ActiveSheet.Cells(i, 13).FormulaR1C1 = "=RC[-2] - RC[-10] + 1"`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ActiveSheet.Cells(i, 12).FormulaR1C1 = "=TODAY() + 30 * 2"`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ActiveSheet.Cells(i, 14).FormulaR1C1 = "=RIGHT(CELL(""filename"", R1C1),LEN(CELL(""filename"", R1C1))-FIND(""]"",CELL(""filename"", R1C1)))"`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ActiveSheet.Cells(i, 15) = "'" & `[`LPad`](LPad)`((i - 1) & "", 4, "0")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ActiveSheet.Cells(i, 16) = "#Proc2Fil TagProc #Proc2Md"`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ActiveSheet.Cells(i, 17) = "ONGOING"`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`i = i + 1`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`iLine = iLine + 1`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Loop`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Set objCode = Nothing`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Set objComponent = Nothing`
&nbsp;&nbsp;&nbsp;&nbsp;`Next objComponent`
&nbsp;&nbsp;&nbsp;&nbsp;`'Clean up and exit.`
&nbsp;&nbsp;&nbsp;&nbsp;`Set objProject = Nothing`
&nbsp;&nbsp;&nbsp;&nbsp;
`ErrorHandler:`
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 30`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Application.ScreenUpdating = True`
`End Sub`


> [!Getting information]
> Ribbon path please refer to ==**Engineer >> Project >> Ribbon**==


# BeCaller
- Ribbon{S}(5)->[[UnHF]]{S}
- Ribbon{S}(6)->[[CleanRgn]]{S}
- Ribbon{S}(34)->[[LPad]]{F}
- Ribbon{S}(48)->[[MyMsgBox]]{S}

