&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub ExpTab()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrorHandler`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim destFile As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim Suffix As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Suffix = Format(Now, "yyyyMMddhhmmss")`  
&nbsp;&nbsp;&nbsp;&nbsp;`destFile = `[`GetBakDrive`](GetBakDrive)`() & "\" & ActiveSheet.Name & "_" & Suffix & ".txt"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If Range("A2") = "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`WriteTxt2Code`](WriteTxt2Code)` "", destFile`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Range("A1").Select`  
&nbsp;&nbsp;&nbsp;&nbsp;`Range(Selection, Selection.End(xlToRight)).Select`  
&nbsp;&nbsp;&nbsp;&nbsp;`Range(Selection, Selection.End(xlDown)).Select`  
&nbsp;&nbsp;&nbsp;&nbsp;`'Range("A1").CurrentRegion.Select`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;[`QuoteTabExpByFileName`](QuoteTabExpByFileName)` destFile, 1`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` "Done", 3`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`ErrorHandler:`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 30`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Auto >> ExpTab**==


# BeCaller
- ExpTab{S}(9)->[[GetBakDrive]]{F}
- ExpTab{S}(11)->[[WriteTxt2Code]]{S}
- ExpTab{S}(17)->[[QuoteTabExpByFileName]]{S}
- ExpTab{S}(18)->[[MyMsgBox]]{S}
- ExpTab{S}(21)->[[MyMsgBox]]{S}

