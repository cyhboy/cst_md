&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Function GenerateRandomUppercase() As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim i As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim randomChar As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim result As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim characters As String`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`characters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" ' Uppercase letters`  
&nbsp;&nbsp;&nbsp;&nbsp;`Randomize`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Generate a random string of 3 uppercase characters`  
&nbsp;&nbsp;&nbsp;&nbsp;`For i = 1 To 3`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`randomChar = Mid(characters, Int((Len(characters) * Rnd) + 1), 1)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`result = result & randomChar`  
&nbsp;&nbsp;&nbsp;&nbsp;`Next i`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`GenerateRandomUppercase = result`  
`End Function`  

