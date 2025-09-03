&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Private Function ReadIniFileString(ByVal Sect As String, ByVal Keyname As String) As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim Worked As Long`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim RetStr As String * 128`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim StrSize As Long`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim iNoOfCharInIni As Integer: iNoOfCharInIni = 0`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim sIniString As String: sIniString = ""`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim sProfileString As String`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If Sect = "" Or Keyname = "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`MsgBox "Section Or Key To Read Not Specified !!!", vbExclamation, "INI"`  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`sProfileString = ""`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`RetStr = `[`Space`](Space)`(128)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`StrSize = Len(RetStr)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Worked = `[`GetPrivateProfileString`](GetPrivateProfileString)`(Sect, Keyname, "", RetStr, StrSize, IniFileName)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If Worked Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`iNoOfCharInIni = Worked`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`sIniString = Left$(RetStr, Worked)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`ReadIniFileString = sIniString`  
`End Function`  


# BeCaller
- ReadIniFileString]]{F}(15)->[[Space]]{F}
- ReadIniFileString]]{F}(17)->[[GetPrivateProfileString]]{F}

