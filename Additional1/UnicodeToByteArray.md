&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Function UnicodeToByteArray(str As String) As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Len(str) = 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim bytes() As Byte`  
&nbsp;&nbsp;&nbsp;&nbsp;`bytes = str`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim l As Long`  
&nbsp;&nbsp;&nbsp;&nbsp;`For l = 0 To UBound(bytes) - 1`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`UnicodeToByteArray = UnicodeToByteArray & "&H" & `[`Hex`](Hex)`(bytes(l)) & ","`  
&nbsp;&nbsp;&nbsp;&nbsp;`Next`  
&nbsp;&nbsp;&nbsp;&nbsp;`UnicodeToByteArray = UnicodeToByteArray & "&H" & `[`Hex`](Hex)`(bytes(UBound(bytes)))`  
`End Function`  


# BeCaller
- UnicodeToByteArray]]{F}(12)->[[Hex]]{F}
- UnicodeToByteArray]]{F}(14)->[[Hex]]{F}

