&nbsp;&nbsp;&nbsp;&nbsp;
`Private Sub Workbook_Open()`
&nbsp;&nbsp;&nbsp;&nbsp;`'On Error GoTo ErrorHandler`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`skipping = False`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fso As Object`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim cFileObject As Object`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim mFileObject As Object`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim obj As Object`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Set fso = CreateObject("Scripting.FileSystemObject")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim updateMacroPath As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim updateUiPath As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim updateUiPathVer As String`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim theFolder As String`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim scriptPath As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim scriptParameter As String`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim finalUiPath As String`
&nbsp;&nbsp;&nbsp;&nbsp;`finalUiPath = Environ("USERPROFILE") & "\AppData\Local\Microsoft\Office\Excel.officeUI"`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim finalMacroPath As String`
&nbsp;&nbsp;&nbsp;&nbsp;`finalMacroPath = ThisWorkbook.FullName`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`theDrive = `[`CommonGetTheDrive`](CommonGetTheDrive)`()`
&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox theDrive`
&nbsp;&nbsp;&nbsp;&nbsp;`'theUser = RespExtMail(Environ$("username"), "EXTERNAL_MAIL")`
&nbsp;&nbsp;&nbsp;&nbsp;`theUser = Environ$("username")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'theUser = extMail(Environ$("username"))`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox theUser`
&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox ReadEnv("%PROGRAMFILES%")`
&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox Environ("AppData")`
&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox Environ("USERPROFILE")`
&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox ThisWorkbook.FullName`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim copyMacroPathVer As String`
&nbsp;&nbsp;&nbsp;&nbsp;`copyMacroPathVer = "C:\AppFiles\cst.xlam"`
&nbsp;&nbsp;&nbsp;&nbsp;`Set cFileObject = fso.GetFile(copyMacroPathVer)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim mFileDate As Date`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim cFileDate As Date`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`mFileDate = DateAdd("yyyy", -5, Now)`
&nbsp;&nbsp;&nbsp;&nbsp;`cFileDate = DateAdd("yyyy", -5, Now)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`cFileDate = cFileObject.DateLastModified`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`For Each obj In fso.Drives()`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox obj.path & " " & obj.DriveType`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'If obj.DriveType = 3 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If obj.path <> "C:" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'If Dir(obj.path & "\AppFiles\SupportSetup\cst.xlam") <> "" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If fso.fileexists(obj.path & "\AppFiles\SupportSetup\cst.xlam") Then`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Set mFileObject = fso.GetFile(obj.path & "\AppFiles\SupportSetup\cst.xlam")`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If mFileObject.DateLastModified > mFileDate Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`mFileDate = mFileObject.DateLastModified`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`theFolder = mFileObject.parentFolder`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`updateMacroPath = obj.path & "\AppFiles\SupportSetup\cst.xlam"`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`updateUiPath = obj.path & "\AppFiles\SupportSetup\Excel.officeUI"`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`updateUiPathVer = obj.path & "\AppFiles\SupportSetup\" & "Excel_" & Environ$("username") & ".officeUI"`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`Next obj`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Set mFileObject = Nothing`
&nbsp;&nbsp;&nbsp;&nbsp;`Set cFileObject = Nothing`
&nbsp;&nbsp;&nbsp;&nbsp;`Set fso = Nothing`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If (mFileDate - cFileDate > 0) Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` "Dear CST Users, Thanks for choosing common support toolkits for your daily work. You now was recommended to upgrade to a new CST version, Please free 1 min to close your office suites and double click " & theFolder & "\install.bat. Thanks very much in deep. ", 10`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'CpFil2Fil updateUiPath, finalUiPath, False`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'CpFil2Fil finalUiPath, updateUiPathVer, False`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'CpFil2Fil updateMacroPath, copyMacroPathVer, False`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'scriptPath = "WScript.exe C:\AppFiles\WaitThenRunHiddenJob.vbs "`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'scriptParameter = """cmd.exe /C copy /Y %22" & updateMacroPath & "%22" & " " & "%22" & finalMacroPath & "%22""" & " " & """5000"""`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'ShellRunHide scriptPath & scriptParameter`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'Sleep 1000`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'ThisWorkbook.Saved = True`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'ThisWorkbook.Close`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;
`ErrorHandler:`
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` "Dear CST Users, When you see this message, The initialization of Common Support Toolkits may encounter some abnormal, It would not affect your daily excel operation, Be patience and try to dump this screen to CST Support, Thanks much. " & Err.Number & " " & Err.Description, 15`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
`End Sub`
&nbsp;&nbsp;&nbsp;&nbsp;


# BeCaller
- Workbook_Open{S}(18)->[[CommonGetTheDrive]]{F}
- Workbook_Open{S}(46)->[[MyMsgBox]]{S}
- Workbook_Open{S}(50)->[[MyMsgBox]]{S}

