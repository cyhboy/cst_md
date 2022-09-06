&nbsp;&nbsp;&nbsp;&nbsp;
`Public Sub ImpCsv()`
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
&nbsp;&nbsp;&nbsp;&nbsp;`On Error GoTo ErrorHandler`
&nbsp;&nbsp;&nbsp;&nbsp;`Application.ScreenUpdating = False`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim filePath As String`
&nbsp;&nbsp;&nbsp;&nbsp;`filePath = `[`FnLstFil`](FnLstFil)`("C:\BAK\*.csv")`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Dim MyData As DataObject`
&nbsp;&nbsp;&nbsp;&nbsp;`Set MyData = New DataObject`
&nbsp;&nbsp;&nbsp;&nbsp;`MyData.SetText ReadLineByFile(filePath)`
&nbsp;&nbsp;&nbsp;&nbsp;`MyData.PutInClipboard`
&nbsp;&nbsp;&nbsp;&nbsp;`Set MyData = Nothing`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`ActiveSheet.Paste`
&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;`Application.ScreenUpdating = True`
`ErrorHandler:`
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 10`
&nbsp;&nbsp;&nbsp;&nbsp;`End If`
`End Sub`


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Document >> ImpCsv**==


# BeCaller
- ImpCsv{S}(8)->[[FnLstFil]]{F}
- ImpCsv{S}(18)->[[MyMsgBox]]{S}

