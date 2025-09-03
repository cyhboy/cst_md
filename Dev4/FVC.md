&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub FVC(Optional colNum As Integer = 9)`  
&nbsp;&nbsp;&nbsp;&nbsp;`' FirstVisibleCell`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`With ActiveSheet.AutoFilter.Range`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'.Range("I" & .offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(.offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row, colNum).Select`  
&nbsp;&nbsp;&nbsp;&nbsp;`End With`  
`End Sub`  

