&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub FmtSht()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`CreTitl`](CreTitl)
&nbsp;&nbsp;&nbsp;&nbsp;`Range("A2").Select`
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`FitScr`](FitScr)
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`Sample`](Sample)
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`DrawTbl`](DrawTbl)
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`Frz`](Frz)
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`RstCf`](RstCf)
`End Sub`


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Auto >> FmtSht**==


# BeCaller
- FmtSht{S}(5)->[[CreTitl]]{S}
- FmtSht{S}(7)->[[FitScr]]{S}
- FmtSht{S}(8)->[[Sample]]{S}
- FmtSht{S}(9)->[[DrawTbl]]{S}
- FmtSht{S}(10)->[[Frz]]{S}
- FmtSht{S}(11)->[[RstCf]]{S}

