&nbsp;&nbsp;&nbsp;&nbsp;
`' ============================================= '`
`' Public Methods`
`' ============================================= '`
&nbsp;&nbsp;&nbsp;&nbsp;
`''`
`' Convert JSON string to object (Dictionary/Collection)`
`'`
`' @method ParseJson`
`' @param {String} json_String`
`' @return {Object} (Dictionary or Collection)`
`' @throws 10001 - JSON parse error`
`''`
`Public Function ParseJson(ByVal JsonString As String) As Object`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then Exit Function`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim json_Index As Long`
&nbsp;&nbsp;&nbsp;&nbsp;`json_Index = 1`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`' Remove vbCr, vbLf, and vbTab from json_String`
&nbsp;&nbsp;&nbsp;&nbsp;`JsonString = VBA.Replace(VBA.Replace(VBA.Replace(JsonString, VBA.vbCr, ""), VBA.vbLf, ""), VBA.vbTab, "")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`json_SkipSpaces JsonString, json_Index`
&nbsp;&nbsp;&nbsp;&nbsp;`Select Case VBA.Mid$(JsonString, json_Index, 1)`
&nbsp;&nbsp;&nbsp;&nbsp;`Case "{"`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Set ParseJson = json_ParseObject(JsonString, json_Index)`
&nbsp;&nbsp;&nbsp;&nbsp;`Case "["`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Set ParseJson = json_ParseArray(JsonString, json_Index)`
&nbsp;&nbsp;&nbsp;&nbsp;`Case Else`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' Error: Invalid JSON string`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(JsonString, json_Index, "Expecting '{' or '['")`
&nbsp;&nbsp;&nbsp;&nbsp;`End Select`
`End Function`

