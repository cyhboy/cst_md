&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub UnhideFile(filePath As String)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Check if the file exists`  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox filePath`  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox Dir(filePath, vbHidden)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Dir(filePath, vbHidden) <> "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' Change the file attribute to normal`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`SetAttr`](SetAttr)` filePath, vbNormal`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox "done"`  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`End Sub`  


# BeCaller
- UnhideFile{S}(6)->[[SetAttr]]{S}

