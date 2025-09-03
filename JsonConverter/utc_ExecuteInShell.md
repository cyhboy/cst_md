&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Private Function utc_ExecuteInShell(utc_ShellCommand As String) As utc_ShellResult`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`#If VBA7 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim utc_File As LongPtr`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim utc_Read As LongPtr`  
`#Else`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim utc_File As Long`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim utc_Read As Long`  
`#End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim utc_Chunk As String`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo utc_ErrorHandling`  
&nbsp;&nbsp;&nbsp;&nbsp;`utc_File = utc_popen(utc_ShellCommand, "r")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If utc_File = 0 Then: Exit Function`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Do While utc_feof(utc_File) = 0`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`utc_Chunk = VBA.Space$(50)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`utc_Read = `[`CLng`](CLng)`(utc_fread(utc_Chunk, 1, Len(utc_Chunk) - 1, utc_File))`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If utc_Read > 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`utc_Chunk = VBA.Left$(utc_Chunk, `[`CLng`](CLng)`(utc_Read))`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`utc_ExecuteInShell.utc_Output = utc_ExecuteInShell.utc_Output & utc_Chunk`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Loop`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`utc_ErrorHandling:`  
&nbsp;&nbsp;&nbsp;&nbsp;`utc_ExecuteInShell.utc_ExitCode = `[`CLng`](CLng)`(utc_pclose(utc_File))`  
`End Function`  


# BeCaller
- utc_ExecuteInShell]]{F}(18)->[[CLng]]{F}
- utc_ExecuteInShell]]{F}(20)->[[CLng]]{F}
- utc_ExecuteInShell]]{F}(25)->[[CLng]]{F}

