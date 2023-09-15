&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub ExpTblWithoutPrompt()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`'Dim lastrow As Long`
&nbsp;&nbsp;&nbsp;&nbsp;`'lastrow = Range("B2").End(xlDown).Row`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim destFile As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim Suffix As String`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Suffix = Format(Now, "yyyyMMddhhmm")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`destFile = `[`GetBakDrive`](GetBakDrive)`() & "\" & ActiveSheet.Name & "_" & Suffix & ".txt"`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Range("A1").Select`
&nbsp;&nbsp;&nbsp;&nbsp;`Range(Selection, Selection.End(xlToRight)).Select`
&nbsp;&nbsp;&nbsp;&nbsp;`Range(Selection, Selection.End(xlDown)).Select`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;[`QuoteCommaExpByFileName`](QuoteCommaExpByFileName)` destFile, 2, """"`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox "DONE!"`
`End Sub`


# BeCaller
- ExpTblWithoutPrompt{S}(8)->[[GetBakDrive]]{F}
- ExpTblWithoutPrompt{S}(12)->[[QuoteCommaExpByFileName]]{S}

