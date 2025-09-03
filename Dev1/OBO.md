&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub OBO(Optional colNum As Integer = 15, Optional desc As Boolean = True)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Range("A1").Select`  
&nbsp;&nbsp;&nbsp;&nbsp;`Range(Selection, Selection.End(xlToRight)).Select`  
&nbsp;&nbsp;&nbsp;&nbsp;`Range(Selection, Selection.End(xlDown)).Select`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear`  
&nbsp;&nbsp;&nbsp;&nbsp;`If desc Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ActiveWorkbook.ActiveSheet.AutoFilter.Sort.SortFields.Add2 Key:=Selection.Columns(colNum), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal`  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ActiveWorkbook.ActiveSheet.AutoFilter.Sort.SortFields.Add2 Key:=Selection.Columns(colNum), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`With ActiveWorkbook.ActiveSheet.AutoFilter.Sort`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.Header = xlYes`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.MatchCase = False`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.Orientation = xlTopToBottom`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.SortMethod = xlPinYin`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.Apply`  
&nbsp;&nbsp;&nbsp;&nbsp;`End With`  
`End Sub`  

