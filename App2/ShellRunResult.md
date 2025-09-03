&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Function ShellRunResult(cmd As String, Optional logFile As String = "C:\BAK\cmd.log", Optional hideFlag As Boolean = True, Optional APPEND As Boolean = False, Optional inputRow As Integer = 1) As GlobalConfig.results`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim cmdLogFile As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`cmdLogFile = Replace(logFile, ".log", "_" & LPad(Trim(str(inputRow)), 6, "0") & ".log")`  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox cmdLogFile`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;[`TchFil`](TchFil)` cmdLogFile, Not APPEND`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim path As String`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If APPEND Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`path = "cmd.exe /C " & cmd & " >> " & cmdLogFile`  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`path = "cmd.exe /C " & cmd & " > " & cmdLogFile`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`path = Replace(path, "2 >", "2>")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'    Dim cntEXE As Integer`  
`'    cntEXE = CntExeRunning(ExtractEXE(path))`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'    If InStr(path, " npm ") > 0 Then`  
`'        cntEXE = cntEXE + CntExeRunning("node.exe")`  
`'    End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox path`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim wsh As Object`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set wsh = VBA.CreateObject("WScript.Shell")`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim waitOnReturn As Boolean: waitOnReturn = True`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim windowStyle As Integer: windowStyle = vbNormal`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If hideFlag Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`windowStyle = vbHide`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`wsh.Run path, windowStyle, waitOnReturn`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`VBA.DoEvents`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Shell path, vbNormalFocus`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Shell path, vbHide`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'    If Hold Then`  
`'        While Hold`  
`'            Dim cntEXE2 As Integer`  
`'            cntEXE2 = CntExeRunning(ExtractEXE(path))`  
`'            If InStr(path, " npm ") > 0 Then`  
`'                cntEXE2 = cntEXE2 + CntExeRunning("node.exe")`  
`'            End If`  
`'            If cntEXE2 - cntEXE > 0 Then`  
`'                Sleep 3000`  
`'            Else`  
`'                Hold = False`  
`'            End If`  
`'        Wend`  
`'    Else`  
`'        While DateDiff("s", LastModDate(cmdLogFile), Now()) < 3`  
`'            Sleep 3000`  
`'        Wend`  
`'    End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim results As GlobalConfig.results`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`With results`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.resultStr = `[`ReadLineByFileUTF8`](ReadLineByFileUTF8)`(cmdLogFile)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.rowNum = inputRow`  
&nbsp;&nbsp;&nbsp;&nbsp;`End With`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`ShellRunResult = results`  
`End Function`  


# BeCaller
- ShellRunResult]]{F}(7)->[[TchFil]]{S}
- ShellRunResult]]{F}(26)->[[ReadLineByFileUTF8]]{F}

