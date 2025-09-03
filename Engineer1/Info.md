&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub Info()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim mes As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`mes = "Thanks for choosing Common Support Toolkits! " & vbCrLf`  
&nbsp;&nbsp;&nbsp;&nbsp;`mes = mes & "Current Workbook is " & ActiveWorkbook.Name & ". " & vbCrLf`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`MsgBox mes, vbInformation, "Information"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 30`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Auto >> Info**==


# BeCaller
- Info{S}(10)->[[MyMsgBox]]{S}

