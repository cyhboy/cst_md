&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub Avlc()`  
&nbsp;&nbsp;&nbsp;&nbsp;`' available country`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim curCell As Range`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim arrStr As String`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim arrAll As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`arrAll = `[`EmptyStringArray`](EmptyStringArray)`()`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim arrPsiphon As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`arrPsiphon = `[`EmptyStringArray`](EmptyStringArray)`()`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim arr As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim astr As Variant`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim offsetCount As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`offsetCount = 0`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim psiphon_common_list As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`psiphon_common_list = `[`EmptyStringArray`](EmptyStringArray)`()`  
&nbsp;&nbsp;&nbsp;&nbsp;`psiphon_common_list = `[`AddToArray`](AddToArray)`(psiphon_common_list, "SG")`  
&nbsp;&nbsp;&nbsp;&nbsp;`psiphon_common_list = `[`AddToArray`](AddToArray)`(psiphon_common_list, "JP")`  
&nbsp;&nbsp;&nbsp;&nbsp;`psiphon_common_list = `[`AddToArray`](AddToArray)`(psiphon_common_list, "GB")`  
&nbsp;&nbsp;&nbsp;&nbsp;`psiphon_common_list = `[`AddToArray`](AddToArray)`(psiphon_common_list, "US")`  
&nbsp;&nbsp;&nbsp;&nbsp;`psiphon_common_list = `[`AddToArray`](AddToArray)`(psiphon_common_list, "PL")`  
&nbsp;&nbsp;&nbsp;&nbsp;`psiphon_common_list = `[`AddToArray`](AddToArray)`(psiphon_common_list, "NL")`  
&nbsp;&nbsp;&nbsp;&nbsp;`psiphon_common_list = `[`AddToArray`](AddToArray)`(psiphon_common_list, "CA")`  
&nbsp;&nbsp;&nbsp;&nbsp;`psiphon_common_list = `[`AddToArray`](AddToArray)`(psiphon_common_list, "AU")`  
&nbsp;&nbsp;&nbsp;&nbsp;`psiphon_common_list = `[`AddToArray`](AddToArray)`(psiphon_common_list, "DE")`  
&nbsp;&nbsp;&nbsp;&nbsp;`psiphon_common_list = `[`AddToArray`](AddToArray)`(psiphon_common_list, "FI")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;[`SltX`](SltX)` 18`  
&nbsp;&nbsp;&nbsp;&nbsp;`For Each curCell In Selection`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`'Set curCell = curRng(i, 0)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If curCell.EntireColumn.Hidden = False And curCell.EntireRow.Hidden = False Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`arrStr = `[`CutStrByStartEnd`](CutStrByStartEnd)`(curCell.Value, "[", "]", True, True)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`arrStr = Replace(arrStr, "[", "")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`arrStr = Replace(arrStr, "]", "")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`arrStr = Replace(arrStr, "'", "")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`arrStr = Replace(arrStr, " ", "")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If arrStr <> "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`arr = Split(arrStr, ",")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`For Each astr In arr`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`arrAll = `[`AddToArray`](AddToArray)`(arrAll, CStr(astr))`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If True = `[`IsInArray`](IsInArray)`(CStr(astr), psiphon_common_list) Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`arrPsiphon = `[`AddToArray`](AddToArray)`(arrPsiphon, CStr(astr))`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox astr`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Next astr`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`offsetCount = offsetCount + 1`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox CutStringByStartAndEnd(curCell.Value, "[", "]")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Next curCell`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox Join(arrAll, ",")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim dictAll As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim dictPsiphon As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim i As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim j As Integer`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Set dictAll = CreateObject("Scripting.Dictionary")`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set dictPsiphon = CreateObject("Scripting.Dictionary")`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`For i = LBound(arrAll) To UBound(arrAll)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If dictAll.Exists(arrAll(i)) Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`dictAll.Item(arrAll(i)) = dictAll.Item(arrAll(i)) + 1`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`dictAll.Add arrAll(i), 1`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Next i`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`For j = LBound(arrPsiphon) To UBound(arrPsiphon)`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If dictPsiphon.Exists(arrPsiphon(j)) Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`dictPsiphon.Item(arrPsiphon(j)) = dictPsiphon.Item(arrPsiphon(j)) + 1`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Else`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`dictPsiphon.Add arrPsiphon(j), 1`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Next j`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim xx As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim yy As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim maxCountAll As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`maxCountAll = 0`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim countryListAll As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`countryListAll = `[`EmptyStringArray`](EmptyStringArray)`()`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`For Each xx In dictAll.Keys`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If CInt(dictAll.Item(xx)) > maxCountAll Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`maxCountAll = CInt(dictAll.Item(xx))`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Next xx`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`For Each yy In dictAll.Keys`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If dictAll.Item(yy) = maxCountAll Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`countryListAll = `[`AddToArray`](AddToArray)`(countryListAll, CStr(yy))`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Next yy`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim maxCountPsiphon As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`maxCountPsiphon = 0`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim countryListPsiphon As Variant`  
&nbsp;&nbsp;&nbsp;&nbsp;`countryListPsiphon = `[`EmptyStringArray`](EmptyStringArray)`()`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`For Each xx In dictPsiphon.Keys`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If CInt(dictPsiphon.Item(xx)) > maxCountPsiphon Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`maxCountPsiphon = CInt(dictPsiphon.Item(xx))`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Next xx`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`For Each yy In dictPsiphon.Keys`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If dictPsiphon.Item(yy) = maxCountPsiphon Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`countryListPsiphon = `[`AddToArray`](AddToArray)`(countryListPsiphon, CStr(yy))`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`Next yy`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`countryListAll = `[`AddToArray`](AddToArray)`(countryListAll, CStr(maxCountAll + offsetCount))`  
&nbsp;&nbsp;&nbsp;&nbsp;`countryListPsiphon = `[`AddToArray`](AddToArray)`(countryListPsiphon, CStr(maxCountPsiphon + offsetCount))`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Set dictAll = Nothing`  
&nbsp;&nbsp;&nbsp;&nbsp;`Set dictPsiphon = Nothing`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`MyMsgBox "best region to download: " & Join(countryListAll, ",") & Chr(13) & Chr(10) & "best region to download for psiphon: " & Join(countryListPsiphon, ","), 5`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Auto >> Avlc**==


# BeCaller
- Avlc{S}(8)->[[EmptyStringArray]]{F}
- Avlc{S}(10)->[[EmptyStringArray]]{F}
- Avlc{S}(16)->[[EmptyStringArray]]{F}
- Avlc{S}(17)->[[AddToArray]]{F}
- Avlc{S}(18)->[[AddToArray]]{F}
- Avlc{S}(19)->[[AddToArray]]{F}
- Avlc{S}(20)->[[AddToArray]]{F}
- Avlc{S}(21)->[[AddToArray]]{F}
- Avlc{S}(22)->[[AddToArray]]{F}
- Avlc{S}(23)->[[AddToArray]]{F}
- Avlc{S}(24)->[[AddToArray]]{F}
- Avlc{S}(25)->[[AddToArray]]{F}
- Avlc{S}(26)->[[AddToArray]]{F}
- Avlc{S}(27)->[[SltX]]{S}
- Avlc{S}(30)->[[CutStrByStartEnd]]{F}
- Avlc{S}(38)->[[AddToArray]]{F}
- Avlc{S}(39)->[[IsInArray]]{F}
- Avlc{S}(40)->[[AddToArray]]{F}
- Avlc{S}(73)->[[EmptyStringArray]]{F}
- Avlc{S}(81)->[[AddToArray]]{F}
- Avlc{S}(87)->[[EmptyStringArray]]{F}
- Avlc{S}(95)->[[AddToArray]]{F}
- Avlc{S}(98)->[[AddToArray]]{F}
- Avlc{S}(99)->[[AddToArray]]{F}

