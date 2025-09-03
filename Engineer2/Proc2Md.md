&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub Proc2Md()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' On Error GoTo ErrorHandler`  
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
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`RobotRun`](RobotRun)` "Proc2Md"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Next curCell`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim module As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`module = Cells(currentRow, 1)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim subb As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`subb = Cells(currentRow, 2)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim call1 As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`call1 = Cells(currentRow, 8)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim call2 As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`call2 = Cells(currentRow, 9)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim call3 As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`call3 = Cells(currentRow, 10)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim beCall As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`beCall = Cells(currentRow, 13)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim menu1 As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`menu1 = Cells(currentRow, 4)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim menu2 As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`menu2 = Cells(currentRow, 5)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim menu3 As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`menu3 = Cells(currentRow, 6)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim menu4 As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`menu4 = Cells(currentRow, 7)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim menu As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`If menu1 = "N/A" Or menu1 = "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`menu = ""`  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If menu4 = "N/A" Or menu4 = "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`menu = menu1 & " >> " & menu2 & " >> " & menu3`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`menu = menu1 & " >> " & menu2 & " >> " & menu3 & " >> " & menu4`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim highLight As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`If menu <> "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`highLight = highLight & "> [!Getting information]" & Chr(13) & Chr(10)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`highLight = highLight & "> Ribbon path please refer to ==**" & menu & "**==" & Chr(13) & Chr(10)`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim i As Long`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim proc As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim VBComp As VBComponent`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim objProject As VBIDE.VBProject`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim objCode As VBIDE.CodeModule`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim codeOfLine As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim codeOfLineMd As String`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim resultStrBas As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim resultStrMd As String`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Set objProject = ThisWorkbook.VBProject`  
&nbsp;&nbsp;&nbsp;&nbsp;`For Each VBComp In objProject.VBComponents`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If module = VBComp.Name Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' Find the code module for the project.`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Set objCode = VBComp.CodeModule`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`For i = 1 To objCode.CountOfLines`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`codeOfLine = objCode.Lines(i, 1)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' If Trim(codeOfLine) <> "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`proc = objCode.ProcOfLine(i, vbext_pk_Proc)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If subb = proc Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`resultStrBas = resultStrBas & codeOfLine & Chr(13) & Chr(10)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`codeOfLine = RTrim(codeOfLine)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If Len(codeOfLine) > 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`codeOfLineMd = `[`LPad`](LPad)`(LTrim(codeOfLine), Len(codeOfLine), "&nbsp;")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`codeOfLine = LTrim(codeOfLine)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`codeOfLineMd = Replace(codeOfLineMd, codeOfLine, "`" & codeOfLine & "`  ")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`codeOfLineMd = Replace(Space(4), " ", "&nbsp;  ")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`resultStrMd = resultStrMd & codeOfLineMd & Chr(13) & Chr(10)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' End If`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Next i`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Next VBComp`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim beCallArr As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`beCallArr = Split(beCall, Chr(13) & Chr(10))`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim j As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim beCallProc As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox UBound(beCallArr)`  
&nbsp;&nbsp;&nbsp;&nbsp;`For j = 0 To UBound(beCallArr) - 1`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If True = `[`EndsWith`](EndsWith)`(CStr(beCallArr(j)), "{S}") Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`beCallProc = `[`CutStringByStartAndEnd`](CutStringByStartAndEnd)`(CStr(beCallArr(j)), "->", "{")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(resultStrMd, "Call " & beCallProc) > 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`resultStrMd = Replace(resultStrMd, "`Call " & beCallProc & "`", "`Call `" & "[`" & beCallProc & "`](" & beCallProc & ")")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`resultStrMd = Replace(resultStrMd, "`" & beCallProc & " ", "[`" & beCallProc & "`](" & beCallProc & ")` ")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`GoTo ContinueLoop`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If True = `[`EndsWith`](EndsWith)`(CStr(beCallArr(j)), "{F}") Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`beCallProc = `[`CutStringByStartAndEnd`](CutStringByStartAndEnd)`(CStr(beCallArr(j)), "->", "{")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`resultStrMd = Replace(resultStrMd, " " & beCallProc & "(", " `[`" & beCallProc & "`](" & beCallProc & ")`(")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`GoTo ContinueLoop`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`ContinueLoop:`  
&nbsp;&nbsp;&nbsp;&nbsp;`Next j`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If highLight <> "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`resultStrMd = resultStrMd & Chr(13) & Chr(10) & Chr(13) & Chr(10) & highLight`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'    If Trim(call1) <> "" Then`  
`'        call1 = Replace(call1, "{S}(", "]]{S}(")`  
`'        call1 = Replace(call1, "{F}(", "]]{F}(")`  
`'        call1 = Replace(call1, "<-", "<-[[")`  
`'        call1 = "# " & "Caller1" & Chr(13) & Chr(10) & call1`  
`'        call1 = Replace(call1, Chr(13) & Chr(10), Chr(13) & Chr(10) & "- ")`  
`'        resultStrMd = resultStrMd & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Left(call1, Len(call1) - 2)`  
`'    End If`  
`'`  
`'    If Trim(call2) <> "" Then`  
`'        call2 = Replace(call2, "{S}(", "]]{S}(")`  
`'        call2 = Replace(call2, "{F}(", "]]{F}(")`  
`'        call2 = Replace(call2, "<-", "<-[[")`  
`'        call2 = "# " & "Caller2" & Chr(13) & Chr(10) & call2`  
`'        call2 = Replace(call2, Chr(13) & Chr(10), Chr(13) & Chr(10) & "- ")`  
`'        resultStrMd = resultStrMd & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Left(call2, Len(call2) - 2)`  
`'    End If`  
`'`  
`'    If Trim(call3) <> "" Then`  
`'        call3 = Replace(call3, "{S}(", "]]{S}(")`  
`'        call3 = Replace(call3, "{F}(", "]]{F}(")`  
`'        call3 = Replace(call3, "<-", "<-[[")`  
`'        call3 = "# " & "Caller3" & Chr(13) & Chr(10) & call3`  
`'        call3 = Replace(call3, Chr(13) & Chr(10), Chr(13) & Chr(10) & "- ")`  
`'        resultStrMd = resultStrMd & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Left(call3, Len(call3) - 2)`  
`'    End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If Trim(beCall) <> "N/A" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`beCall = Replace(beCall, "{S}" & Chr(13) & Chr(10), "]]{S}" & Chr(13) & Chr(10))`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`beCall = Replace(beCall, "{F}", "]]{F}")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`beCall = Replace(beCall, "->", "->[[")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`beCall = "# " & "BeCaller" & Chr(13) & Chr(10) & beCall`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`beCall = Replace(beCall, Chr(13) & Chr(10), Chr(13) & Chr(10) & "- ")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`resultStrMd = resultStrMd & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Left(beCall, Len(beCall) - 2)`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim folderSrc As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim folderMd As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`folderSrc = "C:\SANDBOX\VB_SPACE\GIT_CST_PROJECT\" & Format(Now, "yyyyMMdd") & "\" & module & "\"`  
&nbsp;&nbsp;&nbsp;&nbsp;`'folderSrc = "C:\SANDBOX\VB_SPACE\GIT_CST_MD\" & Format(Now, "yyyyMMddHHmm") & "\" & module & "\"`  
&nbsp;&nbsp;&nbsp;&nbsp;`'folderMd = "C:\MD_SPACE\" & module & "\"`  
&nbsp;&nbsp;&nbsp;&nbsp;`folderMd = "C:\SANDBOX\VB_SPACE\GIT_CST_MD_PROJECT\" & Format(Now, "yyyyMMdd") & "\" & module & "\"`  
&nbsp;&nbsp;&nbsp;&nbsp;[`CreateFolder`](CreateFolder)` folderSrc`  
&nbsp;&nbsp;&nbsp;&nbsp;[`CreateFolder`](CreateFolder)` folderMd`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;[`WriteTxt2Code`](WriteTxt2Code)` resultStrBas, folderSrc & subb & ".bas"`  
&nbsp;&nbsp;&nbsp;&nbsp;[`WriteTxt2Code`](WriteTxt2Code)` resultStrMd, folderMd & subb & ".md"`  
`ErrorHandler:`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 5`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 47) = "###"`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`End Sub`  


# BeCaller
- Proc2Md{S}(15)->[[RobotRun]]{S}
- Proc2Md{S}(77)->[[LPad]]{F}
- Proc2Md{S}(93)->[[EndsWith]]{F}
- Proc2Md{S}(94)->[[CutStringByStartAndEnd]]{F}
- Proc2Md{S}(102)->[[EndsWith]]{F}
- Proc2Md{S}(103)->[[CutStringByStartAndEnd]]{F}
- Proc2Md{S}(124)->[[CreateFolder]]{S}
- Proc2Md{S}(125)->[[CreateFolder]]{S}
- Proc2Md{S}(126)->[[WriteTxt2Code]]{S}
- Proc2Md{S}(127)->[[WriteTxt2Code]]{S}
- Proc2Md{S}(130)->[[MyMsgBox]]{S}

