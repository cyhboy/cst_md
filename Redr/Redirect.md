&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Function Redirect(szBinaryPath As String, szCommandLn As String) As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim tSA_CreatePipe              As SECURITY_ATTRIBUTES`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim tSA_CreateProcessPrc        As SECURITY_ATTRIBUTES`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim tSA_CreateProcessThrd       As SECURITY_ATTRIBUTES`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim tSA_CreateProcessPrcInfo    As PROCESS_INFORMATION`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim tStartupInfo                As STARTUPINFO`  
`#If VBA7 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim hRead                       As LongPtr`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim hWrite                      As LongPtr`  
`#Else`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim hRead                       As Long`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim hWrite                      As Long`  
`#End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim bRead                       As Long`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim abytBuff()                  As Byte`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim lngResult                   As Long`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim szFullCommand               As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim lngExitCode                 As Long`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim lngSizeOf                   As Long`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`tSA_CreatePipe.nLength = Len(tSA_CreatePipe)`  
&nbsp;&nbsp;&nbsp;&nbsp;`tSA_CreatePipe.lpSecurityDescriptor = 0&`  
&nbsp;&nbsp;&nbsp;&nbsp;`tSA_CreatePipe.bInheritHandle = True`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`tSA_CreateProcessPrc.nLength = Len(tSA_CreateProcessPrc)`  
&nbsp;&nbsp;&nbsp;&nbsp;`tSA_CreateProcessThrd.nLength = Len(tSA_CreateProcessThrd)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If (CreatePipe(hRead, hWrite, tSA_CreatePipe, 0&) <> 0&) Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`tStartupInfo.cb = Len(tStartupInfo)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`GetStartupInfo`](GetStartupInfo)` tStartupInfo`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`With tStartupInfo`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.hStdOutput = hWrite`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.hStdError = hWrite`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.wShowWindow = SW_HIDE`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End With`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`szFullCommand = """" & szBinaryPath & """" & " " & szCommandLn`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`lngResult = `[`CreateProcess`](CreateProcess)`(0&, szFullCommand, tSA_CreateProcessPrc, tSA_CreateProcessThrd, True, 0&, 0&, vbNullString, tStartupInfo, tSA_CreateProcessPrcInfo)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If (lngResult <> 0&) Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`lngResult = `[`WaitForSingleObject`](WaitForSingleObject)`(tSA_CreateProcessPrcInfo.hProcess, WAIT_INFINITE)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`lngSizeOf = `[`GetFileSize`](GetFileSize)`(hRead, 0&)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If (lngSizeOf > 0) Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`ReDim abytBuff(lngSizeOf - 1)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If ReadFile(hRead, abytBuff(0), UBound(abytBuff) + 1, bRead, ByVal 0&) Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Redirect = `[`StrConv`](StrConv)`(abytBuff, vbUnicode)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`Call`](Call)` GetExitCodeProcess(tSA_CreateProcessPrcInfo.hProcess, lngExitCode)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`CloseHandle`](CloseHandle)` tSA_CreateProcessPrcInfo.hThread`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`CloseHandle`](CloseHandle)` tSA_CreateProcessPrcInfo.hProcess`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If (lngExitCode <> 0&) Then Err.Raise vbObject + 1235&, "GetExitCodeProcess", "Non-zero Application exist code"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`CloseHandle`](CloseHandle)` hWrite`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`CloseHandle`](CloseHandle)` hRead`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Err.Raise vbObject + 1236&, "CreateProcess", "CreateProcess Failed, Code: " & Err.LastDllError`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`End Function`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  


# BeCaller
- Redirect]]{F}(30)->[[GetStartupInfo]]{S}
- Redirect]]{F}(38)->[[CreateProcess]]{F}
- Redirect]]{F}(40)->[[WaitForSingleObject]]{F}
- Redirect]]{F}(41)->[[GetFileSize]]{F}
- Redirect]]{F}(45)->[[StrConv]]{F}
- Redirect]]{F}(48)->[[Call]]{S}
- Redirect]]{F}(49)->[[CloseHandle]]{S}
- Redirect]]{F}(50)->[[CloseHandle]]{S}
- Redirect]]{F}(52)->[[CloseHandle]]{S}
- Redirect]]{F}(53)->[[CloseHandle]]{S}

