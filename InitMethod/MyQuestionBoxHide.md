&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub MyQuestionBoxHide()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'confirmation = UserForm2.CommandButton1.Caption`
&nbsp;&nbsp;&nbsp;&nbsp;`'UserForm2.Hide`
&nbsp;&nbsp;&nbsp;&nbsp;`confirmation = uf2.CommandButton1.Caption`
&nbsp;&nbsp;&nbsp;&nbsp;`uf2.Hide`
&nbsp;&nbsp;&nbsp;&nbsp;`Set uf2 = Nothing`
&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox "This is a scheduler"`
`End Sub`

