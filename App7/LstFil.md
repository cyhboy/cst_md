&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub LstFil()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrorHandler`
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
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`RobotRunByParam`](RobotRunByParam)` "LstFil"`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Next curCell`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim localFolder As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim keyword As String`
&nbsp;&nbsp;&nbsp;&nbsp;`localFolder = Cells(currentRow, 9)`
&nbsp;&nbsp;&nbsp;&nbsp;`keyword = Cells(currentRow, 13)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim act As String`
&nbsp;&nbsp;&nbsp;&nbsp;`act = Cells(currentRow, 16)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`'    If EndsWith(localFolder, "\") Then`
&nbsp;&nbsp;&nbsp;&nbsp;`'        localFolder = Left(localFolder, Len(localFolder) - 1)`
&nbsp;&nbsp;&nbsp;&nbsp;`'    End If`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox localFolder & keyword`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currFilename As String`
&nbsp;&nbsp;&nbsp;&nbsp;`currFilename = Cells(currentRow, 11)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim lstFilename As String`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fileList As Variant`
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox keyword`
&nbsp;&nbsp;&nbsp;&nbsp;`' fileList = `[`GetFileList`](GetFileList)`(localFolder & keyword & "*")`
&nbsp;&nbsp;&nbsp;&nbsp;`fileList = `[`GetFileList`](GetFileList)`(localFolder & "*" & keyword & "*")`
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox TypeName(fileList)`
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox VarType(fileList)`
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox Len(fileList)`
&nbsp;&nbsp;&nbsp;&nbsp;`' Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim date1 As Date`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`date1 = DateAdd("yyyy", -20, Now)`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fso As Object`
&nbsp;&nbsp;&nbsp;&nbsp;`Set fso = CreateObject("Scripting.FileSystemObject")`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim myFileObj As Object`
&nbsp;&nbsp;&nbsp;&nbsp;`Dim myFile As Variant`
&nbsp;&nbsp;&nbsp;&nbsp;`For Each myFile In fileList`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'If CStr(myFile) <> currFilename Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Set myFileObj = fso.GetFile(localFolder & CStr(myFile))`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If myFileObj.DateLastModified > date1 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`date1 = myFileObj.DateLastModified`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'If rtnType = 1 Then filename1 = myFile.path`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'If rtnType = 2 Then filename1 = myFile.Name`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`lstFilename = myFileObj.Name`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'End If`
&nbsp;&nbsp;&nbsp;&nbsp;`Next myFile`
&nbsp;&nbsp;&nbsp;&nbsp;`Set fso = Nothing`
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 11) = "'" & lstFilename`
&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(act, "EmRe") = 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 12) = date1`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
`ErrorHandler:`
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' MyMsgBox Err.Number & " " & Err.Description, 5`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If Cells(currentRow, 11) = "" Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 11) = Err.Description`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
`End Sub`


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Document >> LstFil**==


# BeCaller
- LstFil{S}(16)->[[RobotRunByParam]]{S}
- LstFil{S}(33)->[[GetFileList]]{F}

