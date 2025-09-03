&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub FitScr(Optional control As IRibbonControl)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Application.WindowState = xlMaximized`  
&nbsp;&nbsp;&nbsp;&nbsp;`ActiveWindow.WindowState = xlMaximized`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim zoom As Double`  
&nbsp;&nbsp;&nbsp;&nbsp;`zoom = ActiveWindow.zoom`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim ww As Double`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim w As Double`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim cw As Double`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim xx As Double`  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox ActiveWindow.Width`  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox ActiveWindow.UsableWidth`  
&nbsp;&nbsp;&nbsp;&nbsp;`ww = ActiveWindow.Width`  
&nbsp;&nbsp;&nbsp;&nbsp;`' ww = ActiveWindow.UsableWidth`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim sumDbl As Double`  
&nbsp;&nbsp;&nbsp;&nbsp;`sumDbl = 0`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim ratioAry As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`ratioAry = Array(0.03, 0.04, 0.04, 0.03, 0.06, 0.065, 0.025, 0.075, 0.07, 0.12, 0.04, 0.12, 0.04, 0.04, 0.025, 0.03, 0.035, 0.025, 0.025, 0.025, 0.025, 0.025, 0.055)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim i As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`For i = 1 To 23`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`sumDbl = sumDbl + ratioAry(i - 1)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`With Columns(i)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`w = .Width`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`cw = .ColumnWidth`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`xx = ww * cw * 100 * ratioAry(i - 1) / w / zoom`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If xx < 255 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.ColumnWidth = xx`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.ColumnWidth = 255`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End With`  
&nbsp;&nbsp;&nbsp;&nbsp;`Next i`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If 1 = 2 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`MyMsgBox "Total fill windows rate --> " & sumDbl, 3`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**NLB >> Email >> FitScr**==

