&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub BCP()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim procNames As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`procNames = Cells(currentRow, 2)`  
&nbsp;&nbsp;&nbsp;&nbsp;`' procNames = Cells(currentRow, 16)`  
&nbsp;&nbsp;&nbsp;&nbsp;`procNames = Replace(procNames, "#", "")`  
&nbsp;&nbsp;&nbsp;&nbsp;`procNames = Replace(procNames, "_", ",")`  
&nbsp;&nbsp;&nbsp;&nbsp;`procNames = Replace(procNames, " ", ",")`  
&nbsp;&nbsp;&nbsp;&nbsp;`procNames = procNames & "," & "Robot,Robp,Robn,Ver,TestVBA,MyMsgBoxHide,Workbook_Open,Worksheet_SelectionChange"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim procNamesArr As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`procNamesArr = Split(procNames, ",")`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim i As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim becallerListAll As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`becallerListAll = ""`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim becallerList As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`For i = 0 To UBound(procNamesArr)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox procNamesArr(i)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`becallerList = `[`BeCallerP`](BeCallerP)`(procNamesArr(i))`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If becallerList <> "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`becallerListAll = becallerListAll & Chr(13) & Chr(10) & becallerList`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Next i`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim regx As Object`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set regx = New RegExp`  
&nbsp;&nbsp;&nbsp;&nbsp;`regx.Pattern = "\([\d]+\)"`  
&nbsp;&nbsp;&nbsp;&nbsp;`regx.Global = True`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`becallerListAll = regx.Replace(becallerListAll, "")`  
&nbsp;&nbsp;&nbsp;&nbsp;`becallerListAll = Replace(becallerListAll, "{S}", "")`  
&nbsp;&nbsp;&nbsp;&nbsp;`becallerListAll = Replace(becallerListAll, "{F}", "")`  
&nbsp;&nbsp;&nbsp;&nbsp;`becallerListAll = Replace(becallerListAll, Chr(13) + Chr(10), "")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`becallerListAll = Left(becallerListAll, Len(becallerListAll) - 2)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'    Cells(currentRow, 13) = becallerListAll`  
`'    Exit Sub`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim becallerListArr As Variant`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`becallerListArr = Split(becallerListAll, "->")`  
&nbsp;&nbsp;&nbsp;&nbsp;`becallerListArr = `[`DeDupeArray`](DeDupeArray)`(becallerListArr)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 13) = Join(becallerListArr, ",") & ",Test"`  
`End Sub`  


# BeCaller
- BCP{S}(20)->[[BeCallerP]]{F}
- BCP{S}(36)->[[DeDupeArray]]{F}

