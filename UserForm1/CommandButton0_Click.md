&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Private Sub CommandButton0_Click()`  
&nbsp;&nbsp;&nbsp;&nbsp;`' On Error Resume Next`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Application.OnTime nexttime, "MyMsgBoxHide """ & gactwb & """, """ & gactws & """", , False`  
&nbsp;&nbsp;&nbsp;&nbsp;`Application.OnTime nexttime, "MyMsgBoxHide", , False`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' UserForm1.Hide`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Unload UserForm1`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If uf1.Visible Then ' Check if the form is visible`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' uf1.Hide`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`Unload`](Unload)` uf1 ' Or MyModalForm.Hide`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' uf1.Hide`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Set uf1 = Nothing`  
&nbsp;&nbsp;&nbsp;&nbsp;`Workbooks(gactwb).Activate`  
&nbsp;&nbsp;&nbsp;&nbsp;`Worksheets(gactws).Activate`  
`End Sub`  


# BeCaller
- CommandButton0_Click{S}(4)->[[Unload]]{S}

