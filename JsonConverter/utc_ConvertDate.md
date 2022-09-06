&nbsp;&nbsp;&nbsp;&nbsp;
`' ============================================= '`
`' Private Functions`
`' ============================================= '`
&nbsp;&nbsp;&nbsp;&nbsp;
`#If Mac Then`
&nbsp;&nbsp;&nbsp;&nbsp;
`Private Function utc_ConvertDate(utc_Value As Date, Optional utc_ConvertToUtc As Boolean = False) As Date`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then Exit Function`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim utc_ShellCommand As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim utc_Result As utc_ShellResult`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim utc_Parts() As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim utc_DateParts() As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim utc_TimeParts() As String`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If utc_ConvertToUtc Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`utc_ShellCommand = "date -ur `date -jf '%Y-%m-%d %H:%M:%S' " & _`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`"'" & VBA.Format$(utc_Value, "yyyy-mm-dd HH:mm:ss") & "' " & _`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`" +'%s'` +'%Y-%m-%d %H:%M:%S'"`
&nbsp;&nbsp;&nbsp;&nbsp;`Else`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`utc_ShellCommand = "date -jf '%Y-%m-%d %H:%M:%S %z' " & _`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`"'" & VBA.Format$(utc_Value, "yyyy-mm-dd HH:mm:ss") & " +0000' " & _`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`"+'%Y-%m-%d %H:%M:%S'"`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`utc_Result = utc_ExecuteInShell(utc_ShellCommand)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If utc_Result.utc_Output = "" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Err.Raise 10015, "UtcConverter.utc_ConvertDate", "'date' command failed"`
&nbsp;&nbsp;&nbsp;&nbsp;`Else`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`utc_Parts = Split(utc_Result.utc_Output, " ")`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`utc_DateParts = Split(utc_Parts(0), "-")`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`utc_TimeParts = Split(utc_Parts(1), ":")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`utc_ConvertDate = `[`DateSerial`](DateSerial)`(utc_DateParts(0), utc_DateParts(1), utc_DateParts(2)) + _`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`TimeSerial(utc_TimeParts(0), utc_TimeParts(1), utc_TimeParts(2))`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
`End Function`


# BeCaller
- utc_ConvertDate{}(25)->[[DateSerial]]{F}

