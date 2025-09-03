&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`' Public Sub MyMsgBoxHide(Optional bkName As String = "", Optional shName As String = "")`  
`Public Sub MyMsgBoxHide()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' On Error Resume Next`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If uf1.Visible Then ' Check if the form is visible`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' uf1.Hide`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`Unload`](Unload)` uf1 ' Or MyModalForm.Hide`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If gactwb <> "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' Workbooks(bkName).Activate`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Workbooks(gactwb).Activate`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If gactws <> "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' Worksheets(shName).Activate`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Worksheets(gactws).Activate`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Application.ScreenUpdating = True`  
&nbsp;&nbsp;&nbsp;&nbsp;`' uf1.Hide`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Set uf1 = Nothing`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Application.WindowState = xlMaximized`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Application.Visible = True`  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox "This is a scheduler"`  
`End Sub`  


# BeCaller
- MyMsgBoxHide{S}(6)->[[Unload]]{S}

