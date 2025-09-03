&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Function ReadLineByFileUTF8(fileName As String)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim objStream As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set objStream = CreateObject("ADODB.Stream")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`objStream.Charset = "UTF-8"`  
&nbsp;&nbsp;&nbsp;&nbsp;`objStream.Type = 2`  
&nbsp;&nbsp;&nbsp;&nbsp;`objStream.Mode = 3`  
&nbsp;&nbsp;&nbsp;&nbsp;`objStream.Open`  
&nbsp;&nbsp;&nbsp;&nbsp;`'objStream.Position = 2`  
&nbsp;&nbsp;&nbsp;&nbsp;`objStream.LoadFromFile fileName`  
&nbsp;&nbsp;&nbsp;&nbsp;`ReadLineByFileUTF8 = Trim(objStream.ReadText())`  
&nbsp;&nbsp;&nbsp;&nbsp;`objStream.Close`  
`End Function`  

