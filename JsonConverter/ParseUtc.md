&nbsp;&nbsp;&nbsp;&nbsp;
`''`
`' VBA-UTC v1.0.6`
`' (c) Tim Hall - https://github.com/VBA-tools/VBA-UtcConverter`
`'`
`' UTC/ISO 8601 Converter for VBA`
`'`
`' Errors:`
`' 10011 - UTC parsing error`
`' 10012 - UTC conversion error`
`' 10013 - ISO 8601 parsing error`
`' 10014 - ISO 8601 conversion error`
`'`
`' @module UtcConverter`
`' @author tim.hall.engr@gmail.com`
`' @license MIT (http://www.opensource.org/licenses/mit-license.php)`
`'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '`
&nbsp;&nbsp;&nbsp;&nbsp;
`' (Declarations moved to top)`
&nbsp;&nbsp;&nbsp;&nbsp;
`' ============================================= '`
`' Public Methods`
`' ============================================= '`
&nbsp;&nbsp;&nbsp;&nbsp;
`''`
`' Parse UTC date to local date`
`'`
`' @method ParseUtc`
`' @param {Date} UtcDate`
`' @return {Date} Local date`
`' @throws 10011 - UTC parsing error`
`''`
`Public Function ParseUtc(utc_UtcDate As Date) As Date`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then Exit Function`
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo utc_ErrorHandling`
&nbsp;&nbsp;&nbsp;&nbsp;
`#If Mac Then`
&nbsp;&nbsp;&nbsp;&nbsp;`ParseUtc = utc_ConvertDate(utc_UtcDate)`
`#Else`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim utc_LocalDate As utc_SYSTEMTIME`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`utc_GetTimeZoneInformation utc_TimeZoneInfo`
&nbsp;&nbsp;&nbsp;&nbsp;`utc_SystemTimeToTzSpecificLocalTime utc_TimeZoneInfo, utc_DateToSystemTime(utc_UtcDate), utc_LocalDate`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`ParseUtc = utc_SystemTimeToDate(utc_LocalDate)`
`#End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`
&nbsp;&nbsp;&nbsp;&nbsp;
`utc_ErrorHandling:`
&nbsp;&nbsp;&nbsp;&nbsp;`Err.Raise 10011, "UtcConverter.ParseUtc", "UTC parsing error: " & Err.Number & " - " & Err.Description`
`End Function`

