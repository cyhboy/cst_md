&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub GP()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim videoPath As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim videoFileName As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;&nbsp;&nbsp;&nbsp;`videoPath = Cells(currentRow, 9)`  
&nbsp;&nbsp;&nbsp;&nbsp;`videoFileName = Cells(currentRow, 11)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim FullPath As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`FullPath = videoPath & videoFileName`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim exeName As String: exeName = `[`ExtractEXE`](ExtractEXE)`("dllhost.exe")`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim cntEXE As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`cntEXE = `[`CntExeRunning`](CntExeRunning)`(exeName)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`With CreateObject("Shell.Application").Namespace(0).ParseName(FullPath)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`.Invokeverb "Properties"`  
&nbsp;&nbsp;&nbsp;&nbsp;`End With`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`While `[`CntExeRunning`](CntExeRunning)`(exeName) = cntEXE + 1`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Sleep 3000`  
&nbsp;&nbsp;&nbsp;&nbsp;`Wend`  
`End Sub`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  


# BeCaller
- GP{S}(13)->[[ExtractEXE]]{F}
- GP{S}(15)->[[CntExeRunning]]{F}

