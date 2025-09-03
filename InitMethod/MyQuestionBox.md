&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub MyQuestionBox(detail As String, answer1 As String, answer2 As String, Optional answer3 As String = "", Optional duration As Long = 120)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If gdetail = detail Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If confirmation = ganswer1 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ElseIf confirmation = ganswer2 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ganswer1 = answer2`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ganswer2 = answer1`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ElseIf confirmation = ganswer3 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ganswer1 = answer3`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ganswer3 = answer1`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`gdetail = detail`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ganswer1 = answer1`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ganswer2 = answer2`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ganswer3 = answer3`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`nexttime = Now() + TimeSerial(0, 0, duration)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Application.OnTime nexttime, "MyQuestionBoxHide"`  
&nbsp;&nbsp;&nbsp;&nbsp;`confirmation = "DNN"`  
&nbsp;&nbsp;&nbsp;&nbsp;`'UserForm2.CommandButton1.Caption = ganswer1`  
&nbsp;&nbsp;&nbsp;&nbsp;`'UserForm2.CommandButton2.Caption = ganswer2`  
&nbsp;&nbsp;&nbsp;&nbsp;`'UserForm2.TextBox1.text = gdetail`  
&nbsp;&nbsp;&nbsp;&nbsp;`'UserForm2.TextBox1.SetFocus`  
&nbsp;&nbsp;&nbsp;&nbsp;`'UserForm2.Show`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Set uf2 = New UserForm2`  
&nbsp;&nbsp;&nbsp;&nbsp;`uf2.CommandButton1.Caption = ganswer1`  
&nbsp;&nbsp;&nbsp;&nbsp;`uf2.CommandButton2.Caption = ganswer2`  
&nbsp;&nbsp;&nbsp;&nbsp;`If ganswer3 <> "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`uf2.CommandButton3.Visible = True`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`uf2.CommandButton3.Caption = ganswer3`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`uf2.CommandButton2.Left = (uf2.CommandButton1.Left + uf2.CommandButton3.Left) / 2`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`uf2.TextBox1.text = detail`  
&nbsp;&nbsp;&nbsp;&nbsp;`uf2.TextBox1.SetFocus`  
&nbsp;&nbsp;&nbsp;&nbsp;`uf2.Show`  
`End Sub`  

