&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub ChkVF()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim n As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`n = Selection.count`  
&nbsp;&nbsp;&nbsp;&nbsp;`If n > 1 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`n = Selection.SpecialCells(xlCellTypeVisible).count`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`If n > 1 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Dim curCell As Range`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`For Each curCell In Selection`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If curCell.EntireColumn.Hidden = False And curCell.EntireRow.Hidden = False Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`curCell.Select`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox subName`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`RobotRun`](RobotRun)` "ChkVF"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Next curCell`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim localFolder As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fileName As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim videoFileName As String`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;&nbsp;&nbsp;&nbsp;`localFolder = Cells(currentRow, 9)`  
&nbsp;&nbsp;&nbsp;&nbsp;`fileName = Cells(currentRow, 11)`  
&nbsp;&nbsp;&nbsp;&nbsp;`videoFileName = localFolder & fileName`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim cmdStr1 As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim cmdStr2 As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`cmdStr1 = "ffprobe -v error -select_streams v:0 -show_entries stream=codec_name,width,height -of csv=s=_:p=0 """ & videoFileName & """"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`cmdStr2 = "ffprobe -v error -select_streams a:0 -show_entries stream=codec_name,bit_rate -of csv=s=_:p=0 """ & videoFileName & """"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox cmdStr1`  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox cmdStr2`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim ffResult1 As GlobalConfig.results`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim ffResult2 As GlobalConfig.results`  
&nbsp;&nbsp;&nbsp;&nbsp;`'ffResultStr = `[`ShellRunResult`](ShellRunResult)`(cmdStr1 & " && " & cmdStr2, "C:\BAK\cmd.log", True, True)`  
&nbsp;&nbsp;&nbsp;&nbsp;`ffResult1 = `[`ShellRunResult`](ShellRunResult)`(cmdStr1, "C:\BAK\cmd.log", False, False, currentRow)`  
&nbsp;&nbsp;&nbsp;&nbsp;`ffResult2 = `[`ShellRunResult`](ShellRunResult)`(cmdStr2, "C:\BAK\cmd.log", False, True, currentRow)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(ffResult2.rowNum, 20) = "'" & ffResult2.resultStr`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Document >> ChkVF**==


# BeCaller
- ChkVF{S}(15)->[[RobotRun]]{S}
- ChkVF{S}(34)->[[ShellRunResult]]{F}
- ChkVF{S}(35)->[[ShellRunResult]]{F}

