&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub RstCf(Optional control As IRibbonControl)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Cells.FormatConditions.Delete`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If 1 = 2 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyQuestionBox`](MyQuestionBox)` "Clean Conditional Format Only? ", "No", "Yes", "", 3`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If confirmation = "Yes" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Cells.Select`  
&nbsp;&nbsp;&nbsp;&nbsp;`Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=CELL(""row"")=ROW()"`  
&nbsp;&nbsp;&nbsp;&nbsp;`Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority`  
&nbsp;&nbsp;&nbsp;&nbsp;`With Selection.FormatConditions(1).Interior`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.PatternColorIndex = xlAutomatic`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.ThemeColor = xlThemeColorAccent2`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.TintAndShade = 0.599963377788629`  
&nbsp;&nbsp;&nbsp;&nbsp;`End With`  
&nbsp;&nbsp;&nbsp;&nbsp;`Selection.FormatConditions(1).StopIfTrue = False`  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**NLB >> BkShtCll >> RstCf**==


# BeCaller
- RstCf{S}(7)->[[MyQuestionBox]]{S}

