&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub ExpQut()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`'    On Error GoTo ErrorHandler`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Range("A1") = "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim destFile As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim Suffix As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Suffix = Format(Now, "yyyyMMddhhmm")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`destFile = `[`GetBakDrive`](GetBakDrive)`() & "\" & ActiveSheet.Name & "_" & Suffix & ".txt"`  
&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox destFile`  
&nbsp;&nbsp;&nbsp;&nbsp;`Range("A1").Select`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If Trim(Range("B1")) <> "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Range(Selection, Selection.End(xlToRight)).Select`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If Trim(Range("A2")) <> "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Range(Selection, Selection.End(xlDown)).Select`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`'Range("A1").CurrentRegion.Select`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;[`QuoteCommaExpByFileName`](QuoteCommaExpByFileName)` destFile, 1, "'"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` "Done", 10`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`'ErrorHandler:`  
&nbsp;&nbsp;&nbsp;&nbsp;`'    If Err.Number <> 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;`'        MyMsgBox Err.Number & " " & Err.Description, 30`  
&nbsp;&nbsp;&nbsp;&nbsp;`'    End If`  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Auto >> ExpQut**==


# BeCaller
- ExpQut{S}(11)->[[GetBakDrive]]{F}
- ExpQut{S}(19)->[[QuoteCommaExpByFileName]]{S}
- ExpQut{S}(20)->[[MyMsgBox]]{S}

