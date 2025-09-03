&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub GP_legacy()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim videoPath As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim videoFileName As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;&nbsp;&nbsp;&nbsp;`videoPath = Cells(currentRow, 9)`  
&nbsp;&nbsp;&nbsp;&nbsp;`videoFileName = Cells(currentRow, 11)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim path As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`path = "'" & `[`GetAppDrive`](GetAppDrive)`() & "\GP.ps1'"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim parameter As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`parameter = "'" & videoPath & "'" & " " & "'" & videoFileName & "'"`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fullCommand As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`fullCommand = path & " " & parameter`  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox fullCommand`  
&nbsp;&nbsp;&nbsp;&nbsp;[`PowerShellRun`](PowerShellRun)` fullCommand, True`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`ErrorHandler:`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 20`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`End Sub`  


# BeCaller
- GP_legacy{S}(10)->[[GetAppDrive]]{F}
- GP_legacy{S}(15)->[[PowerShellRun]]{S}
- GP_legacy{S}(18)->[[MyMsgBox]]{S}

