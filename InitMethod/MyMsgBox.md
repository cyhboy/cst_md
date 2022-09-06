&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub MyMsgBox(detail As String, duration As Long)`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`nexttime = Now() + TimeSerial(0, 0, duration)`
&nbsp;&nbsp;&nbsp;&nbsp;`Application.OnTime nexttime, "MyMsgBoxHide"`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'UserForm1.TextBox1.text = detail`
&nbsp;&nbsp;&nbsp;&nbsp;`'UserForm1.TextBox1.SetFocus`
&nbsp;&nbsp;&nbsp;&nbsp;`'UserForm1.Show`
&nbsp;&nbsp;&nbsp;&nbsp;`Set uf1 = New UserForm1`
&nbsp;&nbsp;&nbsp;&nbsp;`uf1.TextBox1.text = detail`
&nbsp;&nbsp;&nbsp;&nbsp;`uf1.TextBox1.SetFocus`
&nbsp;&nbsp;&nbsp;&nbsp;`uf1.Show`
`End Sub`

