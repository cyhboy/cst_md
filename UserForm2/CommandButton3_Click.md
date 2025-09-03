&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Private Sub CommandButton3_Click()`  
&nbsp;&nbsp;&nbsp;&nbsp;`Application.OnTime nexttime, "MyQuestionBoxHide", , False`  
&nbsp;&nbsp;&nbsp;&nbsp;`'UserForm2.Hide`  
&nbsp;&nbsp;&nbsp;&nbsp;`'confirmation = UserForm2.CommandButton2.Caption`  
&nbsp;&nbsp;&nbsp;&nbsp;`uf2.Hide`  
&nbsp;&nbsp;&nbsp;&nbsp;`confirmation = uf2.CommandButton3.Caption`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set uf2 = Nothing`  
`End Sub`  

