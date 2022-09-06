&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub ConVA()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
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
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`RobotRunByParam`](RobotRunByParam)` "ConVA"`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Next curCell`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim localFolder As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fileName As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim orgFileName As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim videoFileName, audioFileName As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`
&nbsp;&nbsp;&nbsp;&nbsp;`localFolder = Cells(currentRow, 9)`
&nbsp;&nbsp;&nbsp;&nbsp;`fileName = Cells(currentRow, 13)`
&nbsp;&nbsp;&nbsp;&nbsp;`orgFileName = Left(fileName, InStrRev(fileName, ".") - 1)`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fileList As Variant`
&nbsp;&nbsp;&nbsp;&nbsp;`fileList = `[`GetFileList`](GetFileList)`(localFolder & orgFileName & ".f*")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`If UBound(fileList) < 2 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fileSize As Double`
&nbsp;&nbsp;&nbsp;&nbsp;`fileSize = 1.79769313486231E+308`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fso As Object`
&nbsp;&nbsp;&nbsp;&nbsp;`Set fso = CreateObject("Scripting.FileSystemObject")`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim myFileObj As Object`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim myFile As Variant`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`For Each myFile In fileList`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Set myFileObj = fso.GetFile(localFolder & CStr(myFile))`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If FileLen(localFolder & CStr(myFile)) < fileSize Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`videoFileName = audioFileName`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`audioFileName = myFileObj.Name`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`fileSize = FileLen(localFolder & CStr(myFile))`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`Next myFile`
&nbsp;&nbsp;&nbsp;&nbsp;`Set fso = Nothing`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim cmdStr As String`
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox GetVideoDuration(localFolder & videoFileName)`
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox IsNumeric(GetVideoDuration(localFolder & videoFileName))`
&nbsp;&nbsp;&nbsp;&nbsp;`If IsNumeric(GetVideoDuration(localFolder & videoFileName)) And 1 <> 1 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' Never go here as I found copy merge also can fail on a normal case`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`cmdStr = "ffmpeg -y -i """ & localFolder & videoFileName & """ -i """ & localFolder & audioFileName & """ -c copy """ & localFolder & fileName & """"`
&nbsp;&nbsp;&nbsp;&nbsp;`Else`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' cmdStr = "ffmpeg -y -i """ & localFolder & videoFileName & """ -i """ & localFolder & audioFileName & """ -c copy -map 0:v -map 1:a -shortest -af apad """ & localFolder & fileName & """"`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`cmdStr = "ffmpeg -y -i """ & localFolder & videoFileName & """ -i """ & localFolder & audioFileName & """ -af apad -shortest """ & localFolder & fileName & """"`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox cmdStr`
&nbsp;&nbsp;&nbsp;&nbsp;[`ShellRun`](ShellRun)` cmdStr, True`
&nbsp;&nbsp;&nbsp;&nbsp;
`End Sub`


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Auto >> ConVA**==


# BeCaller
- ConVA{S}(15)->[[RobotRunByParam]]{S}
- ConVA{S}(30)->[[GetFileList]]{F}
- ConVA{S}(55)->[[ShellRun]]{S}

