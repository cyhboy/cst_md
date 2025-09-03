&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub BldTbl()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`'On Error GoTo ErrorHandler`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Const ForReading = 1, ForWriting = 2`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim dbname As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim tblname As String`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(ActiveSheet.Name, ".") = 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`dbname = "common_data"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`tblname = ActiveSheet.Name`  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`dbname = Split(ActiveSheet.Name, ".")(0)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`tblname = Split(ActiveSheet.Name, ".")(UBound(Split(ActiveSheet.Name, ".")))`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim sql1 As String`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`sql1 = "drop table " & tblname`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim filePath As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`filePath = `[`FnLstFil`](FnLstFil)`("C:\BAK\" & ActiveSheet.Name & "_*.txt")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim firstLine As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`firstLine = `[`FnGetFileLine`](FnGetFileLine)`(filePath, 1)`  
&nbsp;&nbsp;&nbsp;&nbsp;`firstLine = Left(firstLine, Len(firstLine) - 1)`  
&nbsp;&nbsp;&nbsp;&nbsp;`firstLine = Right(firstLine, Len(firstLine) - 1)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim firstLineAry As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`firstLineAry = Split(firstLine, "','")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim sql2 As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`sql2 = "CREATE TABLE " & tblname & " ("`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim i As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim k As Integer`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`For i = 0 To UBound(firstLineAry)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`sql2 = sql2 & firstLineAry(i) & " LONGTEXT,"`  
&nbsp;&nbsp;&nbsp;&nbsp;`Next i`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`'sql2 = sql2 & "remark LONGTEXT"`  
&nbsp;&nbsp;&nbsp;&nbsp;`sql2 = Left(sql2, Len(sql2) - 1)`  
&nbsp;&nbsp;&nbsp;&nbsp;`sql2 = sql2 & ")"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`MsgBox sql2`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim strPath As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`strPath = "M:\AppFiles\" & dbname & ".accdb"`  
&nbsp;&nbsp;&nbsp;&nbsp;`'strPath = "C:\AppFiles\" & dbname & ".accdb"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If Dir(strPath) = "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`strPath = Replace(strPath, Left(strPath, 2), "C:")`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim db As DAO.Database`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim rst As DAO.Recordset`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Set db = OpenDatabase(strPath)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`On Error Resume Next`  
&nbsp;&nbsp;&nbsp;&nbsp;`'conn.Execute sql1`  
&nbsp;&nbsp;&nbsp;&nbsp;`db.Execute sql1`  
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo 0`  
&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox sql2`  
&nbsp;&nbsp;&nbsp;&nbsp;`db.Execute sql2`  
&nbsp;&nbsp;&nbsp;&nbsp;`'conn.Execute sql2`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim dataLineAry As Variant`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim fso As Object`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim ts As Object`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set fso = CreateObject("Scripting.FileSystemObject")`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set ts = fso.OpenTextFile(filePath, ForReading)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim lineNo As Long`  
&nbsp;&nbsp;&nbsp;&nbsp;`lineNo = 0`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim dataLine As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set rst = db.OpenRecordset(tblname)`  
&nbsp;&nbsp;&nbsp;&nbsp;`Do Until ts.AtEndOfStream`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`lineNo = lineNo + 1`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`dataLine = ts.readline`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If lineNo > 1 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`dataLine = Left(dataLine, Len(dataLine) - 1)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`dataLine = Right(dataLine, Len(dataLine) - 1)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`dataLineAry = Split(dataLine, "','")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`rst.AddNew`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`For k = 0 To UBound(firstLineAry)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`rst(firstLineAry(k)) = Replace(dataLineAry(k), "\n", vbCrLf)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Next k`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`rst.Update`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Loop`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Set ts = Nothing`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set fso = Nothing`  
&nbsp;&nbsp;&nbsp;&nbsp;`rst.Close`  
&nbsp;&nbsp;&nbsp;&nbsp;`db.Close`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`'ErrorHandler:`  
&nbsp;&nbsp;&nbsp;&nbsp;`'    If Err.Number <> 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;`'        MyMsgBox Err.Number & " " & Err.Description, 30`  
&nbsp;&nbsp;&nbsp;&nbsp;`'    End If`  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Auto >> BldTbl**==


# BeCaller
- BldTbl{S}(18)->[[FnLstFil]]{F}
- BldTbl{S}(20)->[[FnGetFileLine]]{F}

