&nbsp;&nbsp;&nbsp;&nbsp;
`Private Sub json_BufferAppend(ByRef json_Buffer As String, _`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`ByRef`](ByRef)` json_Append As Variant, _`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`ByRef`](ByRef)` json_BufferPosition As Long, _`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`ByRef`](ByRef)` json_BufferLength As Long)`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`' VBA can be slow to append strings due to allocating a new string for each append`
&nbsp;&nbsp;&nbsp;&nbsp;`' Instead of using the traditional append, allocate a large empty string and then copy string at append position`
&nbsp;&nbsp;&nbsp;&nbsp;`'`
&nbsp;&nbsp;&nbsp;&nbsp;`' Example:`
&nbsp;&nbsp;&nbsp;&nbsp;`' Buffer: "abc  "`
&nbsp;&nbsp;&nbsp;&nbsp;`' Append: "def"`
&nbsp;&nbsp;&nbsp;&nbsp;`' Buffer Position: 3`
&nbsp;&nbsp;&nbsp;&nbsp;`' Buffer Length: 5`
&nbsp;&nbsp;&nbsp;&nbsp;`'`
&nbsp;&nbsp;&nbsp;&nbsp;`' Buffer position + Append length > Buffer length -> Append chunk of blank space to buffer`
&nbsp;&nbsp;&nbsp;&nbsp;`' Buffer: "abc       "`
&nbsp;&nbsp;&nbsp;&nbsp;`' Buffer Length: 10`
&nbsp;&nbsp;&nbsp;&nbsp;`'`
&nbsp;&nbsp;&nbsp;&nbsp;`' Put "def" into buffer at position 3 (0-based)`
&nbsp;&nbsp;&nbsp;&nbsp;`' Buffer: "abcdef    "`
&nbsp;&nbsp;&nbsp;&nbsp;`'`
&nbsp;&nbsp;&nbsp;&nbsp;`' Approach based on cStringBuilder from vbAccelerator`
&nbsp;&nbsp;&nbsp;&nbsp;`' http://www.vbaccelerator.com/home/VB/Code/Techniques/RunTime_Debug_Tracing/VB6_Tracer_Utility_zip_cStringBuilder_cls.asp`
&nbsp;&nbsp;&nbsp;&nbsp;`'`
&nbsp;&nbsp;&nbsp;&nbsp;`' and clsStringAppend from Philip Swannell`
&nbsp;&nbsp;&nbsp;&nbsp;`' https://github.com/VBA-tools/VBA-JSON/pull/82`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim json_AppendLength As Long`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim json_LengthPlusPosition As Long`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`json_AppendLength = VBA.Len(json_Append)`
&nbsp;&nbsp;&nbsp;&nbsp;`json_LengthPlusPosition = json_AppendLength + json_BufferPosition`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If json_LengthPlusPosition > json_BufferLength Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' Appending would overflow buffer, add chunk`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' (double buffer length or append length, whichever is bigger)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Dim json_AddedLength As Long`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`json_AddedLength = IIf(json_AppendLength > json_BufferLength, json_AppendLength, json_BufferLength)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`json_Buffer = json_Buffer & VBA.Space$(json_AddedLength)`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`json_BufferLength = json_BufferLength + json_AddedLength`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`' Note: Namespacing with VBA.Mid$ doesn't work properly here, throwing compile error:`
&nbsp;&nbsp;&nbsp;&nbsp;`' Function call on left-hand side of assignment must return Variant or Object`
&nbsp;&nbsp;&nbsp;&nbsp;`Mid$(json_Buffer, json_BufferPosition + 1, json_AppendLength) = CStr(json_Append)`
&nbsp;&nbsp;&nbsp;&nbsp;`json_BufferPosition = json_BufferPosition + json_AppendLength`
`End Sub`


# BeCaller
- json_BufferAppend{S}(2)->[[ByRef]]{S}
- json_BufferAppend{S}(3)->[[ByRef]]{S}
- json_BufferAppend{S}(4)->[[ByRef]]{S}

