&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub MERG()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim targetFolder As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`targetFolder = Cells(currentRow, 9)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim sourceFolder As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`sourceFolder = `[`SelectFolder`](SelectFolder)`(targetFolder)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If sourceFolder = "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If False = `[`EndsWith`](EndsWith)`(sourceFolder, "\") Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`sourceFolder = sourceFolder & "\"`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;[`MvFilesAndSubfolders`](MvFilesAndSubfolders)` sourceFolder, targetFolder, "*.mp4"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fileList As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`fileList = `[`GetFileList`](GetFileList)`(sourceFolder & "*.cmd")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim sourceBookPath As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim sourceBook As Workbook`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim copyRange As Range`  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`UnHF`](UnHF)  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`sourceBookPath = `[`ReadLineByFileUTF8`](ReadLineByFileUTF8)`(sourceFolder & CStr(fileList(1)))`  
&nbsp;&nbsp;&nbsp;&nbsp;`sourceBookPath = Replace(sourceBookPath, Chr(13) & Chr(10), "")`  
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 18) = sourceBookPath`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim awb As Workbook`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim aws As Worksheet`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set awb = ActiveWorkbook`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set aws = ActiveSheet`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`MsgBox sourceBookPath`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Dir(sourceBookPath) <> "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Set sourceBook = Workbooks.Open(sourceBookPath)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Set copyRange = sourceBook.Worksheets("Sheet1").Range("A1").CurrentRegion`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`copyRange.Resize(copyRange.Rows.count - 1).offset(1, 0).Copy`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`aws.Range("A1048576").End(xlUp).offset(1, 0).PasteSpecial`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Application.CutCopyMode = False`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`awb.Save`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`sourceBook.Close`  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`MsgBox "index sheet not found, media file copy only. "`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Auto >> PlyVA >> MERG**==


# BeCaller
- MERG{S}(10)->[[SelectFolder]]{F}
- MERG{S}(14)->[[EndsWith]]{F}
- MERG{S}(17)->[[MvFilesAndSubfolders]]{S}
- MERG{S}(19)->[[GetFileList]]{F}
- MERG{S}(23)->[[UnHF]]{S}
- MERG{S}(24)->[[ReadLineByFileUTF8]]{F}

