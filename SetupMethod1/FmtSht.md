&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub FmtSht()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`On Error Resume Next`  
`'    If Application.WindowState <> xlMinimized Then`  
`'        Application.WindowState = xlMinimized`  
`'    End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Call PDPH`  
&nbsp;&nbsp;&nbsp;&nbsp;`If ActiveWorkbook.Name = "cst_template.xlsm" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`RstDt`](RstDt)  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`'Exit Sub`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`CreTitl`](CreTitl)  
&nbsp;&nbsp;&nbsp;&nbsp;`Range("A2").Select`  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`FitScr`](FitScr)  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`Sample`](Sample)  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`DrawTbl`](DrawTbl)  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`Frz`](Frz)  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`RstCf`](RstCf)  
&nbsp;&nbsp;&nbsp;&nbsp;`' Call SavAll`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Auto >> FmtSht**==


# BeCaller
- FmtSht{S}(7)->[[RstDt]]{S}
- FmtSht{S}(9)->[[CreTitl]]{S}
- FmtSht{S}(11)->[[FitScr]]{S}
- FmtSht{S}(12)->[[Sample]]{S}
- FmtSht{S}(13)->[[DrawTbl]]{S}
- FmtSht{S}(14)->[[Frz]]{S}
- FmtSht{S}(15)->[[RstCf]]{S}

