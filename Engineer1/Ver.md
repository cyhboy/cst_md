&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub Ver()`  
&nbsp;&nbsp;&nbsp;&nbsp;`testing = True`  
&nbsp;&nbsp;&nbsp;&nbsp;`'On Error GoTo ErrorHandler`  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`CntOfficeUI`](CntOfficeUI)  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fso, fileObject As Object`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim addInsPath As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`'    addInsPath = "C:\Program Files\Microsoft Office\Office14\Library\cst.xlam"`  
&nbsp;&nbsp;&nbsp;&nbsp;`'    If Is64bit Then`  
&nbsp;&nbsp;&nbsp;&nbsp;`'        addInsPath = "C:\Program Files (x86)\Microsoft Office\Office14\Library\cst.xlam"`  
&nbsp;&nbsp;&nbsp;&nbsp;`'    End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`addInsPath = "C:\AppFiles\cst.xlam"`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set fso = CreateObject("Scripting.FileSystemObject")`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set fileObject = fso.GetFile(addInsPath)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim mes As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`mes = "Thanks for choosing Common Support Toolkits! " & vbCrLf`  
&nbsp;&nbsp;&nbsp;&nbsp;`mes = mes & "The release timestamp of your current copy is " & fileObject.DateLastModified & ". " & vbCrLf`  
&nbsp;&nbsp;&nbsp;&nbsp;`mes = mes & "Current total number of tab definition is " & tabNum & ". " & vbCrLf`  
&nbsp;&nbsp;&nbsp;&nbsp;`mes = mes & "Current total number of group definition is " & groupNum & ". " & vbCrLf`  
&nbsp;&nbsp;&nbsp;&nbsp;`mes = mes & "Current total number of button definition is " & buttonNum & ". " & vbCrLf`  
&nbsp;&nbsp;&nbsp;&nbsp;`mes = mes & "Current total number of menu definition is " & menuNum & ". " & vbCrLf`  
&nbsp;&nbsp;&nbsp;&nbsp;`mes = mes & "Current Workbook is " & ActiveWorkbook.Name & ". " & vbCrLf`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`MsgBox mes, vbInformation, "Version"`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set fileObject = Nothing`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set fso = Nothing`  
&nbsp;&nbsp;&nbsp;&nbsp;`'ErrorHandler:`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 30`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`testing = False`  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Engineer >> Project >> Ver**==


# BeCaller
- Ver{S}(3)->[[CntOfficeUI]]{S}
- Ver{S}(21)->[[MyMsgBox]]{S}

