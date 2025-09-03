&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'Public Sub VsUI()`  
`'    'to be legacy`  
`'    If testing Then Exit Sub`  
`'    'On Error GoTo ErrorHandler`  
`'    Dim n As Integer`  
`'    n = Selection.count`  
`'    If n > 1 Then`  
`'        n = Selection.SpecialCells(xlCellTypeVisible).count`  
`'    End If`  
`'    If n > 1 Then`  
`'        Dim curCell As Range`  
`'        For Each curCell In Selection`  
`'            If curCell.EntireColumn.Hidden = False And curCell.EntireRow.Hidden = False Then`  
`'                curCell.Select`  
`'                'MsgBox subName`  
`'                RobotRun "VsUI"`  
`'            End If`  
`'        Next curCell`  
`'        Exit Sub`  
`'    End If`  
`'`  
`'    Application.ScreenUpdating = False`  
`'`  
`'    Dim currentRow As Integer`  
`'    currentRow = ActiveCell.Row`  
`'`  
`'    Dim procName As String`  
`'`  
`'    procName = Cells(currentRow, 2)`  
`'`  
`'    Dim iCol As Integer`  
`'    iCol = 3`  
`'`  
`'    Dim oXML As Object`  
`'    Set oXML = New DOMDocument`  
`'    Dim strURL As String, resultStr As String`  
`'`  
`'    Dim strFilePath As String`  
`'    strFilePath = "C:\Users\" & Environ$("username") & "\AppData\Local\Microsoft\Office\Excel.officeUI"`  
`'`  
`'    oXML.Load (strFilePath)`  
`'`  
`'    Dim oXmlNodes As IXMLDOMNodeList`  
`'`  
`'    Set oXmlNodes = oXML.SelectNodes("//mso:customUI/mso:ribbon/mso:tabs")`  
`'`  
`'    Dim node As IXMLDOMNode`  
`'    Dim x As String`  
`'    buttonNum = 0`  
`'    groupNum = 0`  
`'    tabNum = 0`  
`'    menuNum = 0`  
`'    For Each node In oXmlNodes`  
`'        x = ListNodes(node, procName, iCol, True)`  
`'        'x = ListNodes(node, procName, iCol, False)`  
`'        If x = "exit" Then`  
`'            'MsgBox "exit"`  
`'            Exit Sub`  
`'        End If`  
`'    Next`  
`'    Set oXML = Nothing`  
`'    Cells(currentRow, 4) = "N/A"`  
`'    Cells(currentRow, 5) = "N/A"`  
`'    Cells(currentRow, 6) = "N/A"`  
`'    Cells(currentRow, 7) = "N/A"`  
`'    'MsgBox "done"`  
`''ErrorHandler:`  
`'    If Err.Number <> 0 Then`  
`'        MyMsgBox Err.Number & " " & Err.Description, 30`  
`'    End If`  
`'    Application.ScreenUpdating = True`  
`'End Sub`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub GetFileProperties_legacy()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim videoPath As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim videoFileName As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim videoFullFilename As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`videoPath = Cells(currentRow, 9)`  
&nbsp;&nbsp;&nbsp;&nbsp;`videoFileName = Cells(currentRow, 11)`  
&nbsp;&nbsp;&nbsp;&nbsp;`videoFullFilename = videoPath & videoFileName`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim resultStr As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim oFolder, ofPName As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim i As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`With CreateObject("Shell.Application")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Set oFolder = .Namespace(Left(videoFullFilename, InStrRev(videoFullFilename, "\") - 1))`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Set ofPName = oFolder.ParseName(Right(videoFullFilename, Len(videoFullFilename) - InStrRev(videoFullFilename, "\")))`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`For i = 1 To 100`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`resultStr = resultStr & Chr(13) & Chr(10) & oFolder.GetDetailsOf(ofPName, i)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Next i`  
&nbsp;&nbsp;&nbsp;&nbsp;`End With`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Cells(currentRow, 18) = resultStr`  
`End Sub`  

