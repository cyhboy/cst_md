&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub RobotRunByParams(comm As String, ParamArray params() As Variant)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim paramSize As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`paramSize = UBound(params) - LBound(params) + 1`  
&nbsp;&nbsp;&nbsp;&nbsp;`If paramSize = 1 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Application.Run comm, params(0)`  
&nbsp;&nbsp;&nbsp;&nbsp;`ElseIf paramSize = 2 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Application.Run comm, params(0), params(1)`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`End Sub`  

