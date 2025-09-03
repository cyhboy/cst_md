&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`''`  
`' Convert local date to UTC date`  
`'`  
`' @method ConvertToUrc`  
`' @param {Date} utc_LocalDate`  
`' @return {Date} UTC date`  
`' @throws 10012 - UTC conversion error`  
`''`  
`Public Function ConvertToUtc(utc_LocalDate As Date) As Date`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo utc_ErrorHandling`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`#If Mac Then`  
&nbsp;&nbsp;&nbsp;&nbsp;`ConvertToUtc = utc_ConvertDate(utc_LocalDate, utc_ConvertToUtc:=True)`  
`#Else`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim utc_UtcDate As utc_SYSTEMTIME`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`utc_GetTimeZoneInformation utc_TimeZoneInfo`  
&nbsp;&nbsp;&nbsp;&nbsp;`utc_TzSpecificLocalTimeToSystemTime utc_TimeZoneInfo, utc_DateToSystemTime(utc_LocalDate), utc_UtcDate`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`ConvertToUtc = utc_SystemTimeToDate(utc_UtcDate)`  
`#End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`utc_ErrorHandling:`  
&nbsp;&nbsp;&nbsp;&nbsp;`Err.Raise 10012, "UtcConverter.ConvertToUtc", "UTC conversion error: " & Err.Number & " - " & Err.Description`  
`End Function`  

