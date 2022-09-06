&nbsp;&nbsp;&nbsp;&nbsp;
`Private Sub json_SkipSpaces(json_String As String, ByRef json_Index As Long)`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`' Increment index to skip over spaces`
&nbsp;&nbsp;&nbsp;&nbsp;`Do While json_Index > 0 And json_Index <= VBA.Len(json_String) And VBA.Mid$(json_String, json_Index, 1) = " "`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`json_Index = json_Index + 1`
&nbsp;&nbsp;&nbsp;&nbsp;`Loop`
`End Sub`

