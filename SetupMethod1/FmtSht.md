&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub FmtSht()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`PDPH`](PDPH)
&nbsp;&nbsp;&nbsp;&nbsp;
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
- FmtSht{S}(5)->[[PDPH]]{S}
- FmtSht{S}(6)->[[CreTitl]]{S}
- FmtSht{S}(8)->[[FitScr]]{S}
- FmtSht{S}(9)->[[Sample]]{S}
- FmtSht{S}(10)->[[DrawTbl]]{S}
- FmtSht{S}(11)->[[Frz]]{S}
- FmtSht{S}(12)->[[RstCf]]{S}

