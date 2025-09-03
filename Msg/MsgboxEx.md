&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`' Purpose   :    Displays a msgbox at a specified location on the screen`  
`' Inputs    :    As per a standard MsgBox +`  
`'               Position                An enumerated type which controls the screen position of the MsgBox`  
`' Outputs   :    As per a standard Msgbox`  
`' Notes     :`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`' Purpose   :    Displays a msgbox at a specified location on the screen`  
`' Inputs    :    As per a standard MsgBox +`  
`'               Position                An enumerated type which controls the screen position of the MsgBox`  
`' Outputs   :    As per a standard Msgbox`  
`' Notes     :    VB only, doesn't work in VBA`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Function MsgboxEx(Prompt As String, Optional Buttons As VbMsgBoxStyle, Optional Title, Optional HelpFile, Optional Context, Optional Position As ePosMsgBox = eCentreScreen) As VbMsgBoxResult`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim lhInst As Long`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim lThread As Long`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`'Set up the CBT hook`  
&nbsp;&nbsp;&nbsp;&nbsp;`lhInst = `[`GetWindowLong`](GetWindowLong)`(GetForegroundWindow, GWL_HINSTANCE)`  
&nbsp;&nbsp;&nbsp;&nbsp;`lThread = `[`GetCurrentThreadId`](GetCurrentThreadId)`()`  
&nbsp;&nbsp;&nbsp;&nbsp;`zlhHook = `[`SetWindowsHookEx`](SetWindowsHookEx)`(WH_CBT, AddressOf zWindowProc, lhInst, lThread)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`zePosition = Position`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`'Display the message box`  
&nbsp;&nbsp;&nbsp;&nbsp;`MsgboxEx = MsgBox(Prompt, Buttons, Title, HelpFile, Context)`  
`End Function`  


# BeCaller
- MsgboxEx]]{F}(7)->[[GetWindowLong]]{F}
- MsgboxEx]]{F}(8)->[[GetCurrentThreadId]]{F}
- MsgboxEx]]{F}(9)->[[SetWindowsHookEx]]{F}

