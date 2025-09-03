&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub AutoCloseMsgBox(Optional sText As String = "Default Information Text", Optional sTitle As String = "Information", Optional WaitTime As Integer = 5)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Message box will be displayed for a set amount of time or until user action is taken`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Adapted from http://www.mrexcel.com/forum/showthread.php?t=20789`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Const btnOK As Integer = 0`  
&nbsp;&nbsp;&nbsp;&nbsp;`Const btnOKCancel As Integer = 1`  
&nbsp;&nbsp;&nbsp;&nbsp;`Const btnAbReIg As Integer = 2`  
&nbsp;&nbsp;&nbsp;&nbsp;`Const btnYesNoCancel As Integer = 3`  
&nbsp;&nbsp;&nbsp;&nbsp;`Const btnYesNo As Integer = 4`  
&nbsp;&nbsp;&nbsp;&nbsp;`Const btnReCancel As Integer = 5`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Const iconStop As Integer = 16 'Show "Stop Mark" icon.`  
&nbsp;&nbsp;&nbsp;&nbsp;`Const iconQ As Integer = 32    'Show "Question Mark" icon.`  
&nbsp;&nbsp;&nbsp;&nbsp;`Const iconExc As Integer = 48  'Show "Exclamation Mark" icon.`  
&nbsp;&nbsp;&nbsp;&nbsp;`Const iconInfo As Integer = 64 'Show "Information Mark" icon.`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim WshShell, RetValue`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set WshShell = CreateObject("WScript.Shell")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' RetValue = WshShell.Popup(sText, WaitTime, sTitle, btnYesNo + iconInfo)`  
&nbsp;&nbsp;&nbsp;&nbsp;`RetValue = WshShell.Popup(sText, WaitTime, sTitle, 4096)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' DoEvents`  
&nbsp;&nbsp;&nbsp;&nbsp;`' AppActivate sTitle`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Valid Return values`  
&nbsp;&nbsp;&nbsp;&nbsp;`' 1 = OK Button, 2 = Cancel Button, 3 = Abort Button, 4 = Retry Button`  
&nbsp;&nbsp;&nbsp;&nbsp;`' 5 = Ignore Button, 6 = Yes Button, 7 = No Button`  
`'    AppActivate ""Information"`  
`'    DoEvents`  
`'    AppActivate Application.Caption`  
`'    DoEvents`  
`'    Select Case RetValue`  
`'       Case 6   'Yes`  
`'            ' MsgBox "Yes"`  
`'       Case 7   'No`  
`'            ' MsgBox "No"`  
`'       Case -1  'No Selection`  
`'            ' MsgBox "No Selection"`  
`'    End Select`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`End Sub`  

