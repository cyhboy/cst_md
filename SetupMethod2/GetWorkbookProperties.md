&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Function GetWorkbookProperties(ByVal filePath As String, ByVal propName As String)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim RetValue As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim appOffice As New Application`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim richFile As Workbook`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set richFile = appOffice.Workbooks.Open(filePath)`  
&nbsp;&nbsp;&nbsp;&nbsp;`RetValue = richFile.BuiltinDocumentProperties(propName)`  
&nbsp;&nbsp;&nbsp;&nbsp;`richFile.Saved = True`  
&nbsp;&nbsp;&nbsp;&nbsp;`'richFile.Close`  
&nbsp;&nbsp;&nbsp;&nbsp;`appOffice.Workbooks.Close`  
&nbsp;&nbsp;&nbsp;&nbsp;`appOffice.Quit`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set appOffice = Nothing`  
&nbsp;&nbsp;&nbsp;&nbsp;`GetWorkbookProperties = RetValue`  
`End Function`  

