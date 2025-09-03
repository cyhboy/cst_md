&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Function SoldFun(param As String)`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim paramArr As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`paramArr = Split(param, "\")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim matchPath As String`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim wbname As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`wbname = ActiveWorkbook.Name`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim srchFolder As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`If True = `[`EndsWith`](EndsWith)`(wbname, "_videos.xlsm") Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`srchFolder = paramArr(UBound(paramArr) - 1)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`matchPath = param`  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`srchFolder = paramArr(UBound(paramArr) - 2)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`matchPath = Left(param, InStr(param, srchFolder) + Len(srchFolder))`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox matchPath`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Exit Sub`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim rootFolderStr As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`rootFolderStr = "D:\"`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim FileSystem As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set FileSystem = CreateObject("Scripting.FileSystemObject")`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim rootFolder As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set rootFolder = FileSystem.GetFolder(rootFolderStr)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim subFolder As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim userFolder As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`For Each subFolder In rootFolder.SubFolders`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`For Each userFolder In subFolder.SubFolders`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If userFolder.Name = srchFolder Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`SoldFun = userFolder & "\"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Next userFolder`  
&nbsp;&nbsp;&nbsp;&nbsp;`Next subFolder`  
`End Function`  


# BeCaller
- SoldFun]]{F}(11)->[[EndsWith]]{F}

