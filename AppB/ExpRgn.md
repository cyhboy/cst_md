&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub ExpRgn()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrorHandler`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim destFile As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim Suffix As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Suffix = Format(Now, "yyyyMMddhhmm")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`destFile = `[`GetBakDrive`](GetBakDrive)`() & "\" & ActiveSheet.Name & "_" & Suffix & ".txt"`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Range("A1").Select`
&nbsp;&nbsp;&nbsp;&nbsp;`'Range(Selection, Selection.End(xlToRight)).Select`
&nbsp;&nbsp;&nbsp;&nbsp;`'Range(Selection, Selection.End(xlDown)).Select`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`ActiveCell.CurrentRegion.Select`
&nbsp;&nbsp;&nbsp;&nbsp;[`QuoteCommaExpByFileName`](QuoteCommaExpByFileName)` destFile, 1, """"`
&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` "DONE!", 2`
&nbsp;&nbsp;&nbsp;&nbsp;
`ErrorHandler:`
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 30`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
`End Sub`


> [!Getting information]
> Ribbon path please refer to ==**Extra >> Common Extra >> ExpRgn**==


# BeCaller
- ExpRgn{S}(9)->[[GetBakDrive]]{F}
- ExpRgn{S}(12)->[[QuoteCommaExpByFileName]]{S}
- ExpRgn{S}(13)->[[MyMsgBox]]{S}
- ExpRgn{S}(16)->[[MyMsgBox]]{S}

