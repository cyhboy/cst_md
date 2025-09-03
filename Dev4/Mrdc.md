&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub Mrdc()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If coding Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Dim mcode As String`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`mcode = `[`Proc2FilFun`](Proc2FilFun)`("Mrdc")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`MsgBox mcode`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'    If Application.WindowState <> xlMinimized Then`  
`'        Application.WindowState = xlMinimized`  
`'    End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`UnHF`](UnHF)  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Call Y2X`  
&nbsp;&nbsp;&nbsp;&nbsp;`' OBO 3`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;[`Cntif`](Cntif)` 23, 9`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`FiltW`](FiltW)  
&nbsp;&nbsp;&nbsp;&nbsp;[`SltX`](SltX)` 9`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Call FVC`  
&nbsp;&nbsp;&nbsp;&nbsp;`silentMode = True`  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`Fold`](Fold)  
&nbsp;&nbsp;&nbsp;&nbsp;`silentMode = False`  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`UnHF`](UnHF)  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;[`SltX`](SltX)` 15`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim totalCnt As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`totalCnt = Selection.count`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim totalBatchCnt As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`totalBatchCnt = -Int(-totalCnt / totalThreadCnt)`  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox totalBatchCnt`  
&nbsp;&nbsp;&nbsp;&nbsp;[`ModBatch`](ModBatch)` 22, 23, totalBatchCnt`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Exit Sub`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim i As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim cntEXE1 As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim cntEXE2 As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim roundNum As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`roundNum = totalBatchCnt - 1`  
&nbsp;&nbsp;&nbsp;&nbsp;`For i = 0 To roundNum`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`cntEXE1 = `[`CntExeRunning`](CntExeRunning)`("cmd.exe")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`UnHF`](UnHF)  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`FiltV`](FiltV)` CStr(i)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Sleep 1000`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`SltX`](SltX)` 9`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox ActiveCell.Row`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' GoTo ContinueLoop`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' Sleep 2000`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`Robot`](Robot)  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' Call RunAppN`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Sleep 1000`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`DoEvents`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`While (CntExeRunning("cmd.exe") - cntEXE1) > 0`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Sleep 5000`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`DoEvents`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Wend`  
`ContinueLoop:`  
&nbsp;&nbsp;&nbsp;&nbsp;`Next i`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`UnHF`](UnHF)  
&nbsp;&nbsp;&nbsp;&nbsp;`' Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;[`SltX`](SltX)` 9`  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`Robp`](Robp)  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Auto >> Mrdc**==


# BeCaller
- Mrdc{S}(7)->[[Proc2FilFun]]{F}
- Mrdc{S}(11)->[[UnHF]]{S}
- Mrdc{S}(12)->[[Cntif]]{S}
- Mrdc{S}(13)->[[FiltW]]{S}
- Mrdc{S}(14)->[[SltX]]{S}
- Mrdc{S}(16)->[[Fold]]{S}
- Mrdc{S}(18)->[[UnHF]]{S}
- Mrdc{S}(19)->[[SltX]]{S}
- Mrdc{S}(24)->[[ModBatch]]{S}
- Mrdc{S}(31)->[[CntExeRunning]]{F}
- Mrdc{S}(32)->[[UnHF]]{S}
- Mrdc{S}(33)->[[FiltV]]{S}
- Mrdc{S}(35)->[[SltX]]{S}
- Mrdc{S}(36)->[[Robot]]{S}
- Mrdc{S}(45)->[[UnHF]]{S}
- Mrdc{S}(46)->[[SltX]]{S}
- Mrdc{S}(47)->[[Robp]]{S}

