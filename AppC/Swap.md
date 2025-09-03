&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub Swap()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrorHandler`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Const ForReading = 1`  
&nbsp;&nbsp;&nbsp;&nbsp;`Const ForWriting = 2`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim filePath As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim strSkipNo As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim strSkipYes As String`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim strText As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim strNewText As String`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim iRet As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim resultMsg As String`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`resultMsg = "Nothing was changed"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`filePath = `[`GetAppDrive`](GetAppDrive)`() & "\FlashFXP\FlashFXP.ini"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`strSkipNo = "Options=1001000000000010011010011111111000000101111101111000"`  
&nbsp;&nbsp;&nbsp;&nbsp;`strSkipYes = "Options=1001000000000010011010011111111000000101111101111001"`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fso As Object`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set fso = CreateObject("Scripting.FileSystemObject")`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim objFile As Object`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set objFile = fso.OpenTextFile(filePath, ForReading)`  
&nbsp;&nbsp;&nbsp;&nbsp;`strText = objFile.readall`  
&nbsp;&nbsp;&nbsp;&nbsp;`objFile.Close`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(strText, strSkipNo) > 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`iRet = MsgBox("Current mode will include sub folder in download and upload, Would U like to change? ", vbYesNo, "Question")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If iRet = vbYes Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`strNewText = Replace(strText, strSkipNo, strSkipYes)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`resultMsg = "Switch mode completed, Folder will skip. "`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Set objFile = fso.OpenTextFile(filePath, ForWriting)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`objFile.Write strNewText`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`objFile.Close`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(strText, strSkipYes) > 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`iRet = MsgBox("Current mode will skip sub folder in download and upload, Would U like to change? ", vbYesNo, "Question")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If iRet = vbYes Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`strNewText = Replace(strText, strSkipYes, strSkipNo)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`resultMsg = "Switch mode completed, Folder will include. "`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Set objFile = fso.OpenTextFile(filePath, ForWriting)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`objFile.Write strNewText`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`objFile.Close`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Set fso = Nothing`  
&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` resultMsg, 10`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`ErrorHandler:`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`resultMsg = Err.Number & " " & Err.Description`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Engineer >> Project >> Switch >> Swap**==


# BeCaller
- Swap{S}(16)->[[GetAppDrive]]{F}
- Swap{S}(46)->[[MyMsgBox]]{S}

