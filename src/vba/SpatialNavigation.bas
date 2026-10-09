' Appended to the FormulaMap module; mFMSource/mFMNodes belong to that module.
' Navigation stays in this Excel instance and never saves/calculates the source.
Private Function FM_NavState() As Object
    Static state As Object
    If state Is Nothing Then Set state = CreateObject("Scripting.Dictionary")
    Set FM_NavState = state
End Function

Public Function StartFormulaMapNavigation(ByVal htmlPath As String, Optional ByVal openBrowser As Boolean = True) As String
    Dim state As Object, fso As Object, allowed As Object, d As Object, item As Variant
    Dim folder As String, scriptPath As String, json As String, nodes() As String, n As Long
    Dim port As Long, command As String, shell As Object, deadline As Date, url As String
    On Error GoTo Failed
    FM_StopNavigation
    If mFMSource Is Nothing Then Err.Raise 5, , "Build the source map first."
    If Len(mFMSource.Path) = 0 Then Err.Raise 5, , "Save the source workbook before connecting."
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(htmlPath) Then Err.Raise 53, , "Map HTML was not found."
    scriptPath = ThisWorkbook.Path & Application.PathSeparator & "navigation_bridge.ps1"
    If Not fso.FileExists(scriptPath) Then Err.Raise 53, , "Place navigation_bridge.ps1 beside this tool."
    folder = Environ$("TEMP") & Application.PathSeparator & "ExcelSpatialMap"
    If Not fso.FolderExists(folder) Then fso.CreateFolder folder
    folder = folder & Application.PathSeparator & fso.GetTempName
    fso.CreateFolder folder
    Set state = FM_NavState(): state.RemoveAll
    Set allowed = CreateObject("Scripting.Dictionary"): allowed.CompareMode = vbBinaryCompare
    ReDim nodes(0 To mFMNodes.Count)
    For Each d In mFMNodes
        If StrComp(CStr(d("book")), mFMSource.Name, vbTextCompare) = 0 And d.Exists("r1") Then
            If d("kind") <> "external" Then
                item = CStr(d("sheet")) & Chr$(30) & UCase$(Replace$(CStr(d("address")), "$", ""))
                If Not allowed.Exists(item) Then
                    allowed.Add item, True
                    nodes(n) = "{""sheet"":" & FM_JSON(CStr(d("sheet"))) & ",""address"":" & FM_JSON(CStr(d("address"))) & "}": n = n + 1
                End If
            End If
        End If
    Next d
    If n = 0 Then Err.Raise 5, , "No source cells are available for navigation."
    ReDim Preserve nodes(0 To n - 1)
    Randomize: port = 40000 + Int(Rnd() * 20000)
    state.Add "source", mFMSource
    state.Add "fullPath", mFMSource.FullName
    state.Add "allowed", allowed
    state.Add "folder", folder
    state.Add "active", True
    state.Add "expires", DateAdd("h", 2, Now)
    state.Add "scheduled", False
    state.Add "macro", "'" & Replace$(ThisWorkbook.Name, "'", "''") & "'!FM_PollNavigation"
    json = "{""Port"":" & CStr(port) & ",""SessionDirectory"":" & FM_JSON(folder) & _
        ",""HtmlPath"":" & FM_JSON(fso.GetAbsolutePathName(htmlPath)) & ",""SourceName"":" & FM_JSON(mFMSource.Name) & _
        ",""SourceFullPath"":" & FM_JSON(mFMSource.FullName) & ",""LifetimeSeconds"":7200,""Allowed"":" & "[" & Join(nodes, ",") & "]}"
    FM_NavWrite folder & "\config.json", json
    FM_PollNavigation
    Set shell = CreateObject("WScript.Shell")
    command = """" & Environ$("SystemRoot") & "\System32\WindowsPowerShell\v1.0\powershell.exe"" -NoLogo -NoProfile -NonInteractive -File """ & scriptPath & """ -ConfigPath """ & folder & "\config.json"""
    shell.Run command, 0, False
    deadline = DateAdd("s", 8, Now)
    Do While Now < deadline
        If fso.FileExists(folder & "\ready.txt") Then
            url = Trim$(FM_NavRead(folder & "\ready.txt"))
            If Left$(url, 17) <> "http://127.0.0.1:" Then Err.Raise 5, , "Invalid navigation address."
            state.Add "url", url
            StartFormulaMapNavigation = url
            If openBrowser Then shell.Run """" & url & """", 1, False
            Exit Function
        End If
        If fso.FileExists(folder & "\startup-error.txt") Then Err.Raise 5, , FM_NavRead(folder & "\startup-error.txt")
        DoEvents
    Loop
    Err.Raise 5, , "Navigation could not start; check PowerShell policy or port availability."
Failed:
    command = Err.Description
    FM_StopNavigation
    Err.Raise 5, "FormulaMap navigation", command
End Function

Public Sub FM_PollNavigation()
    Dim state As Object, fso As Object, folder As String, requestPath As String
    Dim command As String, parts As Variant, id As String, sheet As String, address As String
    Dim reason As String, okay As Boolean, source As Workbook, ws As Worksheet, target As Range
    On Error GoTo PollFailed
    Set state = FM_NavState()
    If Not state.Exists("active") Then Exit Sub
    If Not state("active") Then Exit Sub
    If state("scheduled") Then
        On Error Resume Next
        Application.OnTime EarliestTime:=CDate(state("nextTime")), Procedure:=CStr(state("macro")), Schedule:=False
        Err.Clear
        On Error GoTo PollFailed
    End If
    state("scheduled") = False
    folder = CStr(state("folder"))
    If Now >= CDate(state("expires")) Then FM_StopNavigation: Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folder) Or fso.FileExists(folder & "\stop") Then FM_StopNavigation: Exit Sub
    FM_NavWrite folder & "\heartbeat", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    reason = FM_NavSourceReason(state)
    FM_NavWrite folder & "\state.json", "{""alive"":" & LCase$(CStr(Len(reason) = 0)) & ",""reason"":" & FM_JSON(reason) & "}"
    requestPath = folder & "\request.txt"
    If fso.FileExists(requestPath) Then
        command = FM_NavRead(requestPath)
        fso.DeleteFile requestPath, True
        parts = Split(command, "|")
        If UBound(parts) = 2 Then
            id = parts(0): sheet = FM_Base64ToText(CStr(parts(1))): address = parts(2)
            If Len(id) = 32 And Len(sheet) <= 31 And Len(address) <= 80 Then
                If Len(reason) = 0 Then
                    If Not state("allowed").Exists(sheet & Chr$(30) & UCase$(Replace$(address, "$", ""))) Then
                        reason = "not_in_exported_map"
                    Else
                        Set source = state("source")
                        Set ws = source.Worksheets(sheet)
                        If ws.Visible <> xlSheetVisible Then
                            reason = "hidden_sheet"
                        Else
                            Set target = ws.Range(address)
                            If StrComp(target.Address(False, False), Replace$(address, "$", ""), vbTextCompare) <> 0 Then
                                reason = "invalid_address"
                            Else
                                Application.Goto Reference:=target, Scroll:=True
                                okay = (ActiveWorkbook Is source)
                                If okay Then okay = (ActiveSheet Is ws)
                                If okay Then okay = (Selection.Address(False, False) = target.Address(False, False))
                                If Not okay Then reason = "selection_not_confirmed"
                            End If
                        End If
                    End If
                End If
                FM_NavAck folder, id, okay, sheet, address, reason
            End If
        End If
    End If
ScheduleNext:
    If state("active") Then
        state("nextTime") = DateAdd("s", 1, Now)
        Application.OnTime EarliestTime:=CDate(state("nextTime")), Procedure:=CStr(state("macro")), Schedule:=True
        state("scheduled") = True
    End If
    Exit Sub
PollFailed:
    On Error Resume Next
    If Len(id) > 0 Then FM_NavAck folder, id, False, sheet, address, "excel_navigation_failed"
    If Not state Is Nothing Then
        If state.Exists("active") Then
            If state("active") Then
                Err.Clear
                state("nextTime") = DateAdd("s", 1, Now)
                Application.OnTime EarliestTime:=CDate(state("nextTime")), Procedure:=CStr(state("macro")), Schedule:=True
                state("scheduled") = True
            End If
        End If
    End If
End Sub

Public Sub FM_StopNavigation()
    Dim state As Object
    On Error Resume Next
    Set state = FM_NavState()
    If state.Exists("active") Then
        state("active") = False
        If state.Exists("scheduled") Then
            If state("scheduled") Then Application.OnTime EarliestTime:=CDate(state("nextTime")), Procedure:=CStr(state("macro")), Schedule:=False
            state("scheduled") = False
        End If
        If state.Exists("folder") Then FM_NavWrite CStr(state("folder")) & "\stop", "closed"
    End If
End Sub

Private Function FM_NavSourceReason(ByVal state As Object) As String
    Dim book As Workbook, source As Workbook, found As Boolean
    On Error GoTo ClosedSource
    Set source = state("source")
    For Each book In Application.Workbooks
        If book Is source Then found = True: Exit For
    Next book
    If Not found Then GoTo ClosedSource
    If StrComp(source.FullName, CStr(state("fullPath")), vbBinaryCompare) <> 0 Then
        FM_NavSourceReason = "source_path_changed"
    End If
    Exit Function
ClosedSource:
    FM_NavSourceReason = "source_closed"
End Function

Private Sub FM_NavAck(ByVal folder As String, ByVal id As String, ByVal okay As Boolean, ByVal sheet As String, ByVal address As String, ByVal reason As String)
    FM_NavWrite folder & "\ack.json", "{""id"":" & FM_JSON(id) & ",""ok"":" & LCase$(CStr(okay)) & _
        ",""sheet"":" & FM_JSON(sheet) & ",""address"":" & FM_JSON(address) & ",""reason"":" & FM_JSON(reason) & "}"
End Sub

Private Sub FM_NavWrite(ByVal path As String, ByVal value As String)
    Dim stream As Object, fso As Object, temp As String
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.WriteText value
    temp = path & ".writing": stream.SaveToFile temp, 2: stream.Close
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(path) Then fso.DeleteFile path, True
    fso.MoveFile temp, path
End Sub

Private Function FM_NavRead(ByVal path As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.LoadFromFile path: FM_NavRead = stream.ReadText: stream.Close
End Function
