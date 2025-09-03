&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`Public Sub Fold()`  
&nbsp;&nbsp;&nbsp;&nbsp;`If testing Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Application.ScreenUpdating = False`  
&nbsp;&nbsp;&nbsp;&nbsp;`' On Error GoTo ErrorHandler`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim n As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`n = Selection.count`  
&nbsp;&nbsp;&nbsp;&nbsp;`If n > 1 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`n = Selection.SpecialCells(xlCellTypeVisible).count`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`If n > 1 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Dim curCell As Range`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`For Each curCell In Selection`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If curCell.EntireColumn.Hidden = False And curCell.EntireRow.Hidden = False Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`curCell.Select`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox subName`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`RobotRun`](RobotRun)` "Fold"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Next curCell`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Exit Sub`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim activeName As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`activeName = ActiveWorkbook.FullName`  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox activeName`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim strDirectory As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim currentRow As Integer`  
&nbsp;&nbsp;&nbsp;&nbsp;`currentRow = ActiveCell.Row`  
&nbsp;&nbsp;&nbsp;&nbsp;`strDirectory = Cells(currentRow, 9)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If False = `[`EndsWith`](EndsWith)`(strDirectory, "\") Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`strDirectory = strDirectory & "\"`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
`'    Dim strDirectoryArr As Variant`  
`'    If `[`EndsWith`](EndsWith)`(activeName, "_playlists.xlsm") Then`  
`'        strDirectoryArr = Split(strDirectory, "\")`  
`'        strDirectory = Replace(strDirectory, strDirectoryArr(3), "")`  
`'        strDirectory = Replace(strDirectory, "\\", "\")`  
`'    End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;[`CreateFolder`](CreateFolder)` strDirectory`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' Enhanced for youtube-dl`  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim command As String`  
&nbsp;&nbsp;&nbsp;&nbsp;`command = Cells(currentRow, 10)`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Dim cookie_time As Date`  
&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(command, "youtube") > 0 And InStr(command, "yt-dlp") > 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`cookie_time = `[`LastModDate`](LastModDate)`("C:\Users\" & Environ$("username") & "\Desktop\youtube.com_cookies.txt")`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If `[`DateDiff`](DateDiff)`("s", cookie_time, Now()) > 800 Or Dir("C:\Users\" & Environ$("username") & "\Desktop\youtube.com_cookies.txt") = "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`ExCookies`](ExCookies)` "https://www.youtube.com/"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`UnhideFile`](UnhideFile)` strDirectory & "youtube.com_cookies.txt"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`CpFil2Fil`](CpFil2Fil)` "C:\Users\" & Environ$("username") & "\Desktop\youtube.com_cookies.txt", strDirectory & "youtube.com_cookies.txt", True`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`CpFil2Fil`](CpFil2Fil)` "C:\Users\" & Environ$("username") & "\Desktop\youtube.com_cookies.txt", "D:\learning\" & "youtube.com_cookies.txt", True`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' CpFil2Fil "C:\Users\" & Environ$("username") & "\Downloads\www.youtube.com.txt", "D:\learning\" & "youtube.com_cookies.txt", True`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' HideFile strDirectory & "youtube.com_cookies.txt"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`UnhideFile`](UnhideFile)` strDirectory & "compareE.bat"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If Dir(strDirectory & "compareE.bat") = "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`CpFil2Fil`](CpFil2Fil)` "C:\Users\" & Environ$("username") & "\Desktop\compareE.bat", strDirectory & "compareE.bat", True`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`HideFile`](HideFile)` strDirectory & "compareE.bat"`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`Wr2Cmd`](Wr2Cmd)  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(command, "douyin") > 0 And InStr(command, "yt-dlp") > 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If `[`DateDiff`](DateDiff)`("s", `[`LastModDate`](LastModDate)`("C:\Users\" & Environ$("username") & "\Desktop\douyin.com_cookies.txt"), Now()) > 1800 Or Dir("C:\Users\" & Environ$("username") & "\Desktop\douyin.com_cookies.txt") = "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`ExCookies`](ExCookies)` "https://www.douyin.com/"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`UnhideFile`](UnhideFile)` strDirectory & "douyin.com_cookies.txt"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`CpFil2Fil`](CpFil2Fil)` "C:\Users\" & Environ$("username") & "\Desktop\douyin.com_cookies.txt", strDirectory & "douyin.com_cookies.txt", True`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' CpFil2Fil "C:\Users\" & Environ$("username") & "\Desktop\douyin.com_cookies.json", strDirectory & "douyin.com_cookies.json", True, True`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' HideFile strDirectory & "douyin.com_cookies.txt"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`UnhideFile`](UnhideFile)` strDirectory & "compareE.bat"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If Dir(strDirectory & "compareE.bat") = "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`CpFil2Fil`](CpFil2Fil)` "C:\Users\" & Environ$("username") & "\Desktop\compareE.bat", strDirectory & "compareE.bat", True`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`HideFile`](HideFile)` strDirectory & "compareE.bat"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`Wr2Cmd`](Wr2Cmd)  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`If InStr(command, "bilibili") > 0 And InStr(command, "yt-dlp") > 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If `[`DateDiff`](DateDiff)`("s", `[`LastModDate`](LastModDate)`("C:\Users\" & Environ$("username") & "\Desktop\bilibili.com_cookies.txt"), Now()) > 1800 Or Dir("C:\Users\" & Environ$("username") & "\Desktop\bilibili.com_cookies.txt") = "" Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`ExCookies`](ExCookies)` "https://www.bilibili.com/"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`UnhideFile`](UnhideFile)` strDirectory & "bilibili.com_cookies.txt"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`CpFil2Fil`](CpFil2Fil)` "C:\Users\" & Environ$("username") & "\Desktop\bilibili.com_cookies.txt", strDirectory & "bilibili.com_cookies.txt", True`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`' HideFile strDirectory & "bilibili.com_cookies.txt"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`UnhideFile`](UnhideFile)` strDirectory & "compareE.bat"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`If False = `[`FileExists`](FileExists)`(strDirectory & "compareE.bat") Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`CpFil2Fil`](CpFil2Fil)` "C:\Users\" & Environ$("username") & "\Desktop\compareE.bat", strDirectory & "compareE.bat", True`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`HideFile`](HideFile)` strDirectory & "compareE.bat"`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`Call `[`Wr2Cmd`](Wr2Cmd)  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`'    Dim cell As Object`  
&nbsp;&nbsp;&nbsp;&nbsp;`'    For Each cell In Selection.Cells`  
&nbsp;&nbsp;&nbsp;&nbsp;`'        If cell.EntireColumn.Hidden = False And cell.EntireRow.Hidden = False Then`  
&nbsp;&nbsp;&nbsp;&nbsp;`'            currentRow = cell.Row`  
&nbsp;&nbsp;&nbsp;&nbsp;`'            strDirectory = Cells(currentRow, 9)`  
&nbsp;&nbsp;&nbsp;&nbsp;`'            CreateFolder strDirectory`  
&nbsp;&nbsp;&nbsp;&nbsp;`'        End If`  
&nbsp;&nbsp;&nbsp;&nbsp;`'    Next cell`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox "Create standard folder " & strDirectory & " successfully"`  
&nbsp;&nbsp;&nbsp;&nbsp;`' MsgBox "Create standard folder successfully"`  
`ErrorHandler:`  
&nbsp;&nbsp;&nbsp;&nbsp;`If Err.Number <> 0 Then`  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[`MyMsgBox`](MyMsgBox)` Err.Number & " " & Err.Description, 30`  
&nbsp;&nbsp;&nbsp;&nbsp;`End If`  
&nbsp;  &nbsp;  &nbsp;  &nbsp;  
&nbsp;&nbsp;&nbsp;&nbsp;`Application.ScreenUpdating = True`  
`End Sub`  


> [!Getting information]
> Ribbon path please refer to ==**Customize >> Auto >> Fold**==


# BeCaller
- Fold{S}(16)->[[RobotRun]]{S}
- Fold{S}(27)->[[EndsWith]]{F}
- Fold{S}(30)->[[CreateFolder]]{S}
- Fold{S}(35)->[[LastModDate]]{F}
- Fold{S}(36)->[[DateDiff]]{F}
- Fold{S}(37)->[[ExCookies]]{S}
- Fold{S}(39)->[[UnhideFile]]{S}
- Fold{S}(40)->[[CpFil2Fil]]{S}
- Fold{S}(41)->[[CpFil2Fil]]{S}
- Fold{S}(42)->[[UnhideFile]]{S}
- Fold{S}(44)->[[CpFil2Fil]]{S}
- Fold{S}(46)->[[HideFile]]{S}
- Fold{S}(47)->[[Wr2Cmd]]{S}
- Fold{S}(50)->[[DateDiff]]{F}
- Fold{S}(51)->[[ExCookies]]{S}
- Fold{S}(53)->[[UnhideFile]]{S}
- Fold{S}(54)->[[CpFil2Fil]]{S}
- Fold{S}(55)->[[UnhideFile]]{S}
- Fold{S}(57)->[[CpFil2Fil]]{S}
- Fold{S}(59)->[[HideFile]]{S}
- Fold{S}(60)->[[Wr2Cmd]]{S}
- Fold{S}(63)->[[DateDiff]]{F}
- Fold{S}(64)->[[ExCookies]]{S}
- Fold{S}(66)->[[UnhideFile]]{S}
- Fold{S}(67)->[[CpFil2Fil]]{S}
- Fold{S}(68)->[[UnhideFile]]{S}
- Fold{S}(69)->[[FileExists]]{F}
- Fold{S}(70)->[[CpFil2Fil]]{S}
- Fold{S}(72)->[[HideFile]]{S}
- Fold{S}(73)->[[Wr2Cmd]]{S}
- Fold{S}(77)->[[MyMsgBox]]{S}

