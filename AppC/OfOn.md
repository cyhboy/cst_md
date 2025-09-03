&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub OfOn(Optional control As IRibbonControl)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrorHandler`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim flag As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`flag = ActiveWorkbook.Sheets("Info").Range("A1")`  
&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(flag, "#") > 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ActiveWorkbook.Sheets("Info").Range("A1") = Trim(Replace(flag, "#", ""))`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` "This file have been onlined", 10`  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ActiveWorkbook.Sheets("Info").Range("A1") = Trim(flag) & " #"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` "This file have been offlined", 10`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`ErrorHandler:`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 30`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`End Sub`  


# BeCaller
- OfOn{S}(10)->[[MyMsgBox]]{S}
- OfOn{S}(13)->[[MyMsgBox]]{S}
- OfOn{S}(17)->[[MyMsgBox]]{S}

