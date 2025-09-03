&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Private Sub CommandButton1_Click()`  
&nbsp;&nbsp;&nbsp;&nbsp;`Application.OnTime nexttime, "MyQuestionBoxHide", , False`  
&nbsp;&nbsp;&nbsp;&nbsp;`'UserForm2.Hide`  
&nbsp;&nbsp;&nbsp;&nbsp;`'confirmation = UserForm2.CommandButton1.Caption`  
&nbsp;&nbsp;&nbsp;&nbsp;`uf2.Hide`  
&nbsp;&nbsp;&nbsp;&nbsp;`confirmation = uf2.CommandButton1.Caption`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set uf2 = Nothing`  
`End Sub`  

