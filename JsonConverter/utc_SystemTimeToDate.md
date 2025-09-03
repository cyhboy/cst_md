&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Private Function utc_SystemTimeToDate(utc_Value As utc_SYSTEMTIME) As Date`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`utc_SystemTimeToDate = `[`DateSerial`](DateSerial)`(utc_Value.utc_wYear, utc_Value.utc_wMonth, utc_Value.utc_wDay) + _`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`TimeSerial(utc_Value.utc_wHour, utc_Value.utc_wMinute, utc_Value.utc_wSecond)`  
`End Function`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`#End If`  


# BeCaller
- utc_SystemTimeToDate]]{F}(5)->[[DateSerial]]{F}

