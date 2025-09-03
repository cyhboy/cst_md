&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Private Function WriteIniFileString(ByVal Sect As String, ByVal Keyname As String, ByVal Wstr As String) As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim Worked As Long`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim iNoOfCharInIni As Integer: iNoOfCharInIni = 0`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim sIniString As String: sIniString = ""`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Sect = "" Or Keyname = "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`MsgBox "Section Or Key To Write Not Specified !!!", vbExclamation, "INI"`  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Worked = `[`WritePrivateProfileString`](WritePrivateProfileString)`(Sect, Keyname, Wstr, IniFileName)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If Worked Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`iNoOfCharInIni = Worked`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`sIniString = Wstr`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`WriteIniFileString = sIniString`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`End Function`  


# BeCaller
- WriteIniFileString]]{F}(11)->[[WritePrivateProfileString]]{F}

