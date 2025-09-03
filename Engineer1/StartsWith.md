&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Function StartsWith(str As String, start As String) As Boolean`  
`'    If testing Then`  
`'        Exit Function`  
`'    End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`str = CStr(str)`  
&nbsp;&nbsp;&nbsp;&nbsp;`start = CStr(start)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim startLen As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`startLen = Len(start)`  
&nbsp;&nbsp;&nbsp;&nbsp;`StartsWith = (Left(UCase(str), startLen) = UCase(start))`  
`End Function`  

