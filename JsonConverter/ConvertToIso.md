&nbsp;&nbsp;&nbsp;&nbsp;
`''`
`' Convert local date to ISO 8601 string`
`'`
`' @method ConvertToIso`
`' @param {Date} utc_LocalDate`
`' @return {Date} ISO 8601 string`
`' @throws 10014 - ISO 8601 conversion error`
`''`
`Public Function ConvertToIso(utc_LocalDate As Date) As String`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then Exit Function`
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo utc_ErrorHandling`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`ConvertToIso = VBA.Format$(ConvertToUtc(utc_LocalDate), "yyyy-mm-ddTHH:mm:ss.000Z")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`
&nbsp;&nbsp;&nbsp;&nbsp;
`utc_ErrorHandling:`
&nbsp;&nbsp;&nbsp;&nbsp;`Err.Raise 10014, "UtcConverter.ConvertToIso", "ISO 8601 conversion error: " & Err.Number & " - " & Err.Description`
`End Function`

