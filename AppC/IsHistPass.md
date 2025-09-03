&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Function IsHistPass(inpPass As String) As Boolean`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Function`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrorHandler`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'    Dim inpPass As String`  
`'    inpPass = Cells(currentRow, 5)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim cn As Object`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set cn = CreateObject("ADODB.Connection")`  
&nbsp;&nbsp;&nbsp;&nbsp;`cn.Open "DSN=HACHBLUT;UID=hachkua1;PWD=g4f4r4d5"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim sSQL As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`sSQL = "SELECT distinct(PASS) AS APASS FROM HAC.TASKPENDING WHERE UPPER(PASS) = '" & UCase(inpPass) & "' AND CO like 'CP_" & "%'"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim rs As Object`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set rs = CreateObject("ADODB.RecordSet")`  
&nbsp;&nbsp;&nbsp;&nbsp;`rs.Open sSQL, cn`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If Not rs.EOF Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'MsgBox "The Password " & inpPass & " was in history already, Please Regen. "`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`IsHistPass = True`  
&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`IsHistPass = False`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`rs.Close`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set rs = Nothing`  
&nbsp;&nbsp;&nbsp;&nbsp;`cn.Close`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set cn = Nothing`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`ErrorHandler:`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'MyMsgBox Err.Number & " " & Err.Description, 30`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`IsHistPass = True`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`End Function`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  

