&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub FmtSht()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`'On Error Resume Next`  
`'    If Application.WindowState <> xlMinimized Then`  
`'        Application.WindowState = xlMinimized`  
`'    End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If ActiveWorkbook.ActiveSheet.Name = "Info" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyQuestionBox`](MyQuestionBox.md)` "format info sheet or not? ", "No", "Yes", "", 5`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If confirmation = "No" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ActiveWorkbook.Sheets("Sheet1").Activate`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Call PDPH`  
&nbsp;&nbsp;&nbsp;&nbsp;`If ActiveWorkbook.Name = "cst_template.xlsm" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`RstDt`](RstDt.md)  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`'Exit Sub`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`CreTitl`](CreTitl.md)  
&nbsp;&nbsp;&nbsp;&nbsp;`Range("A2").Select`  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`FitScr`](FitScr.md)  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`Sample`](Sample.md)  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`DrawTbl`](DrawTbl.md)  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`Frz`](Frz.md)  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`RstCf`](RstCf.md)  
&nbsp;&nbsp;&nbsp;&nbsp;`' Call SavAll`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Auto >> FmtSht**==


# BeCaller
- FmtSht{S}(6)->[[MyQuestionBox]]{S}
- FmtSht{S}(12)->[[RstDt]]{S}
- FmtSht{S}(14)->[[CreTitl]]{S}
- FmtSht{S}(16)->[[FitScr]]{S}
- FmtSht{S}(17)->[[Sample]]{S}
- FmtSht{S}(18)->[[DrawTbl]]{S}
- FmtSht{S}(19)->[[Frz]]{S}
- FmtSht{S}(20)->[[RstCf]]{S}

