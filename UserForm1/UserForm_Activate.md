&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Private Sub UserForm_Activate()`  
&nbsp;&nbsp;&nbsp;&nbsp;`On Error Resume Next`  
&nbsp;&nbsp;&nbsp;&nbsp;`Me.Repaint`  
&nbsp;&nbsp;&nbsp;&nbsp;[`AppActivate`](AppActivate)` Application.Caption`  
&nbsp;&nbsp;&nbsp;&nbsp;[`AppActivate`](AppActivate)` Me.Caption`  
`End Sub`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'Private Sub UserForm_Initialize()`  
`'`  
`'    Const C_VBA6_USERFORM_CLASSNAME = "ThunderDFrame"`  
`'`  
`'    Dim ret As Long`  
`'    Dim formHWnd As LongPtr`  
`'`  
`'    'Get window handle of the userform`  
`'`  
`'    formHWnd = FindWindow(C_VBA6_USERFORM_CLASSNAME, Me.Caption)`  
`'    If formHWnd = 0 Then`  
`'        Debug.Print Err.LastDllError`  
`'    End If`  
`'`  
`'    'Set userform window to 'always on top'`  
`'`  
`'    ret = SetWindowPos(formHWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)`  
`'    If ret = 0 Then`  
`'        Debug.Print Err.LastDllError`  
`'    End If`  
`'`  
`'End Sub`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  


# BeCaller
- UserForm_Activate{S}(4)->[[AppActivate]]{S}
- UserForm_Activate{S}(5)->[[AppActivate]]{S}

