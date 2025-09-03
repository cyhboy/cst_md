&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Private Sub UserForm_Initialize()`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Const C_VBA6_USERFORM_CLASSNAME = "ThunderDFrame"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim ret As Long`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim formHWnd As LongPtr`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`'Get window handle of the userform`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`formHWnd = `[`FindWindow`](FindWindow)`(C_VBA6_USERFORM_CLASSNAME, Me.Caption)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If formHWnd = 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Debug.Print Err.LastDllError`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`'Set userform window to 'always on top'`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`ret = `[`SetWindowPos`](SetWindowPos)`(formHWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If ret = 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Debug.Print Err.LastDllError`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`End Sub`  


# BeCaller
- UserForm_Initialize{S}(5)->[[FindWindow]]{F}
- UserForm_Initialize{S}(9)->[[SetWindowPos]]{F}

