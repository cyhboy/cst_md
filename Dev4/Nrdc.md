&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub Nrdc()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If coding Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Dim mcode As String`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`mcode = `[`Proc2FilFun`](Proc2FilFun)`("Nrdc")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`MsgBox mcode`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'    If Application.WindowState <> xlMinimized Then`  
`'        Application.WindowState = xlMinimized`  
`'    End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Call UnHF`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`Y2X`](Y2X)  
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
&nbsp;&nbsp;&nbsp;&nbsp;[`ModBatch`](ModBatch)` 22, 15, totalBatchCnt`  
&nbsp;&nbsp;&nbsp;&nbsp;`' Exit Sub`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim i, j As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim cntEXE1 As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim cntEXE2 As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim roundNum As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`roundNum = totalBatchCnt - 1`  
&nbsp;&nbsp;&nbsp;&nbsp;`For i = 0 To roundNum`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`cntEXE1 = `[`CntExeRunning`](CntExeRunning)`("cmd.exe")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`UnHF`](UnHF)  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`FiltV`](FiltV)` CStr(i)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`SltX`](SltX)` 9`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`Robot`](Robot)  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' Call RunAppN`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Sleep 3000`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`DoEvents`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`While (CntExeRunning("cmd.exe") - cntEXE1) > 0`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Sleep 3000`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`DoEvents`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Wend`  
&nbsp;&nbsp;&nbsp;&nbsp;`Next i`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`UnHF`](UnHF)  
&nbsp;&nbsp;&nbsp;&nbsp;`' Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;[`SltX`](SltX)` 9`  
&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`Robp`](Robp)  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;`For j = 0 To roundNum`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`cntEXE1 = `[`CntExeRunning`](CntExeRunning)`("cmd.exe")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`UnHF`](UnHF)  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`FiltV`](FiltV)` CStr(i)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`SltX`](SltX)` 9`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' Call Robot`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`RunAppN`](RunAppN)  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Sleep 3000`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`DoEvents`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`While (CntExeRunning("cmd.exe") - cntEXE1) > 0`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Sleep 3000`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`DoEvents`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Wend`  
&nbsp;&nbsp;&nbsp;&nbsp;`Next j`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`End Sub`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  


# BeCaller
- Nrdc{S}(7)->[[Proc2FilFun]]{F}
- Nrdc{S}(11)->[[Y2X]]{S}
- Nrdc{S}(12)->[[Cntif]]{S}
- Nrdc{S}(13)->[[FiltW]]{S}
- Nrdc{S}(14)->[[SltX]]{S}
- Nrdc{S}(16)->[[Fold]]{S}
- Nrdc{S}(18)->[[UnHF]]{S}
- Nrdc{S}(19)->[[SltX]]{S}
- Nrdc{S}(24)->[[ModBatch]]{S}
- Nrdc{S}(31)->[[CntExeRunning]]{F}
- Nrdc{S}(32)->[[UnHF]]{S}
- Nrdc{S}(33)->[[FiltV]]{S}
- Nrdc{S}(34)->[[SltX]]{S}
- Nrdc{S}(35)->[[Robot]]{S}
- Nrdc{S}(43)->[[UnHF]]{S}
- Nrdc{S}(44)->[[SltX]]{S}
- Nrdc{S}(45)->[[Robp]]{S}
- Nrdc{S}(47)->[[CntExeRunning]]{F}
- Nrdc{S}(48)->[[UnHF]]{S}
- Nrdc{S}(49)->[[FiltV]]{S}
- Nrdc{S}(50)->[[SltX]]{S}
- Nrdc{S}(51)->[[RunAppN]]{S}

