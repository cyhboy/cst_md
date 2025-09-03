&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`' Purpose   :    Returns the screen dimensions, not including the tastbar`  
`' Inputs    :    N/A`  
`' Outputs   :    A type which defines the extent of the screen work area.`  
`' Notes     :`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Function ScreenWorkArea() As RECT`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim tScreen As RECT`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim lRet As Long`  
&nbsp;&nbsp;&nbsp;&nbsp;`Const SPI_GETWORKAREA = 48`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`lRet = `[`SystemParametersInfo`](SystemParametersInfo)`(SPI_GETWORKAREA, vbNull, tScreen, 0)`  
&nbsp;&nbsp;&nbsp;&nbsp;`ScreenWorkArea = tScreen`  
`End Function`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  


# BeCaller
- ScreenWorkArea]]{F}(8)->[[SystemParametersInfo]]{F}

