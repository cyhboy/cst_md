&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub Cntif(colFun, colCnt)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim colOffset As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`colOffset = colCnt - colFun`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Cells(2, colFun).Select`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(2, colFun).FormulaR1C1 = "=COUNTIF(R2C" & colCnt & ":RC[" & colOffset & "],RC[" & colOffset & "])"`  
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(2, colFun).Select`  
&nbsp;&nbsp;&nbsp;&nbsp;`If ActiveCell.CurrentRegion.Rows.count > 2 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Selection.AutoFill Destination:=Range(Cells(2, colFun), Cells(ActiveCell.CurrentRegion.Rows.count, colFun))`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`ActiveWindow.LargeScroll ToRight:=-1`  
`End Sub`  

