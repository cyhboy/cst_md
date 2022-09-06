&nbsp;&nbsp;&nbsp;&nbsp;
`#Else`
&nbsp;&nbsp;&nbsp;&nbsp;
[`Private`](Private)` Function utc_DateToSystemTime(utc_Value As Date) As utc_SYSTEMTIME`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then Exit Function`
&nbsp;&nbsp;&nbsp;&nbsp;`utc_DateToSystemTime.utc_wYear = VBA.Year(utc_Value)`
&nbsp;&nbsp;&nbsp;&nbsp;`utc_DateToSystemTime.utc_wMonth = VBA.Month(utc_Value)`
&nbsp;&nbsp;&nbsp;&nbsp;`utc_DateToSystemTime.utc_wDay = VBA.Day(utc_Value)`
&nbsp;&nbsp;&nbsp;&nbsp;`utc_DateToSystemTime.utc_wHour = VBA.Hour(utc_Value)`
&nbsp;&nbsp;&nbsp;&nbsp;`utc_DateToSystemTime.utc_wMinute = VBA.Minute(utc_Value)`
&nbsp;&nbsp;&nbsp;&nbsp;`utc_DateToSystemTime.utc_wSecond = VBA.Second(utc_Value)`
&nbsp;&nbsp;&nbsp;&nbsp;`utc_DateToSystemTime.utc_wMilliseconds = 0`
`End Function`


# BeCaller
- utc_DateToSystemTime{}(2)->[[Private]]{S}

