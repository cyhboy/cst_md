&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub MyQuestionBox(detail As String, answer1 As String, answer2 As String, duration As Long)`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`nexttime = Now() + TimeSerial(0, 0, duration)`
&nbsp;&nbsp;&nbsp;&nbsp;`Application.OnTime nexttime, "MyQuestionBoxHide"`
&nbsp;&nbsp;&nbsp;&nbsp;`confirmation = ""`
&nbsp;&nbsp;&nbsp;&nbsp;`'UserForm2.CommandButton1.Caption = answer1`
&nbsp;&nbsp;&nbsp;&nbsp;`'UserForm2.CommandButton2.Caption = answer2`
&nbsp;&nbsp;&nbsp;&nbsp;`'UserForm2.TextBox1.text = detail`
&nbsp;&nbsp;&nbsp;&nbsp;`'UserForm2.TextBox1.SetFocus`
&nbsp;&nbsp;&nbsp;&nbsp;`'UserForm2.Show`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Set uf2 = New UserForm2`
&nbsp;&nbsp;&nbsp;&nbsp;`uf2.CommandButton1.Caption = answer1`
&nbsp;&nbsp;&nbsp;&nbsp;`uf2.CommandButton2.Caption = answer2`
&nbsp;&nbsp;&nbsp;&nbsp;`uf2.TextBox1.text = detail`
&nbsp;&nbsp;&nbsp;&nbsp;`uf2.TextBox1.SetFocus`
&nbsp;&nbsp;&nbsp;&nbsp;`uf2.Show`
`End Sub`

