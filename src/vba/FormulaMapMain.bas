Attribute VB_Name = "FormulaMap"
Option Explicit

Private Const FM_MAX_FORMULAS As Long = 50000
Private Const FM_MAX_NATIVE_NODES As Long = 80
Private Const FM_MAX_MEMBERSHIP_CHECKS As Long = 2000000
Private mFMNodes As Collection, mFMEdges As Collection, mFMIssues As Collection
Private mFMNodeIndex As Object, mFMEdgeIndex As Object, mFMSheets As Collection
Private mFMSource As Workbook
Private mFMFormulaCount As Long, mFMTotalFormulas As Double, mFMLimited As Boolean
Private mFMWorksheetJSON As String

Public Sub BuildFormulaMap()
    Dim src As Workbook, w As Workbook, candidates As Collection
    Dim prompt As String, answer As Variant, i As Long, folder As String, result As String
    On Error GoTo Failed
    Set candidates = New Collection
    For Each w In Application.Workbooks
        If Not w Is ThisWorkbook Then
            candidates.Add w
            prompt = prompt & candidates.Count & ". " & w.Name & vbCrLf
        End If
    Next w
    If candidates.Count = 0 Then
        MsgBox "Open the workbook to analyze first." & vbCrLf & "Open the workbook to analyze first.", vbInformation
        Exit Sub
    End If
    If Not ActiveWorkbook Is ThisWorkbook Then
        Set src = ActiveWorkbook
    Else
        answer = Application.InputBox("Select a workbook number" & vbCrLf & vbCrLf & prompt, "Formula map", 1, Type:=1)
        If VarType(answer) = vbBoolean Then Exit Sub
        If answer <> Fix(answer) Or answer < 1 Or answer > candidates.Count Then
            MsgBox "Enter a workbook number from the list.", vbInformation
            Exit Sub
        End If
        Set src = candidates(CLng(answer))
    End If
    With Application.FileDialog(4)
        .Title = "Choose output folder"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        folder = .SelectedItems(1)
    End With
    FM_StopNavigation
    result = FM_BuildFile(src.Name, folder, True)
    On Error GoTo NavigationUnavailable
    result = StartFormulaMapNavigation(result, True)
    Exit Sub
NavigationUnavailable:
    MsgBox "The map was created, but Excel navigation could not start." & vbCrLf & _
        Err.Description & vbCrLf & vbCrLf & _
        "Open the generated HTML to explore the map. The Cell position map in the report also links to source cells." & vbCrLf & result, vbInformation
    Exit Sub
Failed:
    MsgBox "The formula map could not be created." & vbCrLf & Err.Description, vbExclamation
End Sub

Public Function FM_BuildFile(ByVal sourceBookName As String, ByVal outputFolder As String, Optional ByVal silent As Boolean = True) As String
    Dim oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean, oldStatus As Variant
    Dim oldCancel As Long, ws As Worksheet, formulas As Range, cell As Range
    Dim refs As Collection, problems As Collection, ref As Object, item As Variant
    Dim target As String, source As String, report As Workbook, base As String, htmlPath As String
    Dim errorNumber As Long, errorText As String
    oldScreen = Application.ScreenUpdating: oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts: oldStatus = Application.StatusBar
    oldCancel = Application.EnableCancelKey
    On Error GoTo Failed
    Set mFMSource = Application.Workbooks(sourceBookName)
    If Len(Dir$(outputFolder, vbDirectory)) = 0 Then Err.Raise 5, , "Output folder was not found: " & outputFolder
    Application.ScreenUpdating = False: Application.EnableEvents = False
    Application.DisplayAlerts = False: Application.EnableCancelKey = xlErrorHandler
    Set mFMNodes = New Collection: Set mFMEdges = New Collection: Set mFMIssues = New Collection
    Set mFMSheets = New Collection
    Set mFMNodeIndex = CreateObject("Scripting.Dictionary"): mFMNodeIndex.CompareMode = vbTextCompare
    Set mFMEdgeIndex = CreateObject("Scripting.Dictionary"): mFMEdgeIndex.CompareMode = vbTextCompare
    mFMFormulaCount = 0: mFMTotalFormulas = 0: mFMLimited = False
    For Each ws In mFMSource.Worksheets
        mFMSheets.Add ws.Name
        Set formulas = Nothing
        On Error Resume Next
        Set formulas = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo Failed
        If Not formulas Is Nothing Then
            For Each cell In formulas.Cells
                If FM_IsFormulaAnchor(cell) Then mFMTotalFormulas = mFMTotalFormulas + 1
            Next cell
        End If
    Next ws
    For Each ws In mFMSource.Worksheets
        Set formulas = Nothing
        On Error Resume Next
        Set formulas = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo Failed
        If Not formulas Is Nothing Then
            For Each cell In formulas.Cells
                If FM_IsFormulaAnchor(cell) Then
                If mFMFormulaCount >= FM_MAX_FORMULAS Then
                    mFMLimited = True
                    Exit For
                End If
                mFMFormulaCount = mFMFormulaCount + 1
                target = FM_EnsureNode(mFMSource.Name, ws.Name, cell.Address(False, False), "formula")
                Set problems = New Collection
                Set refs = FM_ParseFormula(cell, problems)
                For Each ref In refs
                    source = FM_EnsureNode(CStr(ref("book")), CStr(ref("sheet")), CStr(ref("address")), CStr(ref("kind")))
                    FM_AddEdge source, target, CStr(ref("kind")), CStr(ref("note"))
                Next ref
                For Each item In problems
                    FM_AddIssue ws.Name, cell.Address(False, False), FM_GetFormula(cell), CStr(item)
                Next item
                If mFMFormulaCount Mod 100 = 0 Then
                    Application.StatusBar = "Analyzing formulas: " & Format$(mFMFormulaCount, "#,##0") & " / " & Format$(mFMTotalFormulas, "#,##0")
                    DoEvents
                End If
                End If
            Next cell
        End If
        If mFMLimited Then Exit For
    Next ws
    If mFMLimited Then FM_AddIssue "", "", "", "Partial analysis: formula scan limit reached."
    Application.StatusBar = "Linking formulas inside ranges"
    FM_LinkRangeMembers
    Application.StatusBar = "Capturing worksheet positions"
    mFMWorksheetJSON = FM_SnapshotJSON()
    base = outputFolder & Application.PathSeparator & "FormulaMap_" & FM_SafeName(mFMSource.Name) & "_" & Format$(Now, "yyyymmdd_hhnnss")
    base = FM_UniqueBase(base)
    htmlPath = base & ".html"
    Set report = FM_CreateReport()
    FM_AddNativeSpatialMap report
    report.SaveAs Filename:=base & ".xlsx", FileFormat:=xlOpenXMLWorkbook
    FM_WriteUTF8 htmlPath, Replace$(FM_Template(), "__FORMULA_MAP_DATA__", FM_DataJSON())
    report.Worksheets(1).Activate
    FM_BuildFile = htmlPath
Clean:
    Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts: Application.EnableCancelKey = oldCancel
    Application.StatusBar = oldStatus
    If errorNumber <> 0 Then
        If Not report Is Nothing Then
            On Error Resume Next
            report.Close False
            On Error GoTo 0
        End If
        Err.Raise errorNumber, "FormulaMap", errorText
    End If
    If Not silent Then
        MsgBox "Formula map created." & vbCrLf & vbCrLf & _
            "Formulas analyzed: " & Format$(mFMFormulaCount, "#,##0") & vbCrLf & _
            "Review notices: " & mFMIssues.Count & vbCrLf & vbCrLf & _
            htmlPath & vbCrLf & base & ".xlsx" & vbCrLf & vbCrLf & _
            "Open the HTML to search, zoom, and switch between English and Korean.", vbInformation
    End If
    Exit Function
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    Resume Clean
End Function

Private Function FM_IsFormulaAnchor(ByVal cell As Range) As Boolean
    If cell.MergeCells Then
        If cell.Address <> cell.MergeArea.Cells(1, 1).Address Then Exit Function
    End If
    FM_IsFormulaAnchor = CBool(cell.HasFormula)
End Function

Private Function FM_GetFormula(ByVal cell As Range) As String
    On Error Resume Next
    FM_GetFormula = CStr(CallByName(cell, "Formula2", VbGet))
    If Err.Number <> 0 Then Err.Clear: FM_GetFormula = CStr(cell.Formula)
    On Error GoTo 0
End Function

Private Function FM_EnsureNode(ByVal book As String, ByVal sheet As String, ByVal address As String, ByVal kind As String) As String
    Dim key As String, d As Object, r As Range, ws As Worksheet, cell As Range, displayRange As Range
    If Len(book) = 0 Then book = mFMSource.Name
    key = book & Chr$(30) & sheet & Chr$(30) & Replace$(address, "$", "")
    If mFMNodeIndex.Exists(key) Then
        FM_EnsureNode = mFMNodeIndex(key)
        Exit Function
    End If
    Set d = CreateObject("Scripting.Dictionary")
    d("id") = "n" & (mFMNodes.Count + 1): d("book") = book
    d("sheet") = sheet: d("address") = Replace$(address, "$", "")
    d("label") = "": d("value") = "": d("formula") = "": d("kind") = kind
    If StrComp(book, mFMSource.Name, vbTextCompare) = 0 And kind <> "external" Then
        On Error Resume Next
        Set ws = mFMSource.Worksheets(sheet)
        If Not ws Is Nothing Then Set r = ws.Range(address)
        On Error GoTo 0
        If Not r Is Nothing Then
            If r.CountLarge = 1 Then
                If r.HasFormula Then d("kind") = "formula": d("formula") = FM_GetFormula(r)
                d("value") = FM_CellText(r)
                d("label") = FM_NearbyLabel(r)
            Else
                d("value") = Format$(r.CountLarge, "#,##0") & " cells"
                d("kind") = "range"
            End If
            d("r1") = r.Row: d("c1") = r.Column
            d("r2") = r.Row + r.Rows.Count - 1: d("c2") = r.Column + r.Columns.Count - 1
            Set displayRange = r
            If r.CountLarge = 1 Then
                If r.MergeCells Then Set displayRange = r.MergeArea
            End If
            d("x") = CDbl(displayRange.Left) * 4 / 3: d("y") = CDbl(displayRange.Top) * 4 / 3
            d("w") = CDbl(displayRange.Width) * 4 / 3: d("h") = CDbl(displayRange.Height) * 4 / 3
        End If
    Else
        d("kind") = "external"
        d("label") = "External workbook"
    End If
    mFMNodes.Add d: mFMNodeIndex.Add key, d("id")
    FM_EnsureNode = d("id")
End Function

Private Function FM_CellText(ByVal cell As Range) As String
    On Error Resume Next
    If IsError(cell.Value2) Then
        FM_CellText = CStr(cell.Text)
    Else
        FM_CellText = CStr(cell.Value2)
    End If
    FM_CellText = Left$(FM_CellText, 400)
    On Error GoTo 0
End Function

Private Function FM_NearbyLabel(ByVal cell As Range) As String
    Dim neighbor As Range, s As String
    On Error Resume Next
    If cell.Column > 1 Then Set neighbor = cell.Offset(0, -1)
    If Not neighbor Is Nothing Then
        If Not neighbor.HasFormula And VarType(neighbor.Value2) = vbString Then
            s = Trim$(CStr(neighbor.Value2))
            If Len(s) > 0 And Len(s) <= 60 Then FM_NearbyLabel = neighbor.Address(False, False) & ": " & s: Exit Function
        End If
    End If
    Set neighbor = Nothing
    If cell.Row > 1 Then Set neighbor = cell.Offset(-1, 0)
    If Not neighbor Is Nothing Then
        If Not neighbor.HasFormula And VarType(neighbor.Value2) = vbString Then
            s = Trim$(CStr(neighbor.Value2))
            If Len(s) > 0 And Len(s) <= 60 Then FM_NearbyLabel = neighbor.Address(False, False) & ": " & s
        End If
    End If
    On Error GoTo 0
End Function

Private Sub FM_AddEdge(ByVal source As String, ByVal target As String, ByVal kind As String, ByVal note As String)
    Dim key As String, d As Object
    key = source & ">" & target & ":" & kind
    If mFMEdgeIndex.Exists(key) Then Exit Sub
    Set d = CreateObject("Scripting.Dictionary")
    d("s") = source: d("t") = target: d("kind") = kind: d("note") = note
    mFMEdges.Add d: mFMEdgeIndex.Add key, True
End Sub

Private Sub FM_AddIssue(ByVal sheet As String, ByVal address As String, ByVal formula As String, ByVal reason As String)
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("sheet") = sheet: d("address") = address: d("formula") = formula: d("reason") = reason
    mFMIssues.Add d
End Sub

Private Sub FM_LinkRangeMembers()
    Dim ranges As Collection, formulas As Collection, d As Object, f As Object
    Dim r As Range, c As Range, checks As Long, wb As String
    Set ranges = New Collection: Set formulas = New Collection
    For Each d In mFMNodes
        If d("kind") = "range" And d("book") = mFMSource.Name Then ranges.Add d
        If d("kind") = "formula" Then formulas.Add d
    Next d
    For Each d In ranges
        Set r = Nothing
        On Error Resume Next
        Set r = mFMSource.Worksheets(d("sheet")).Range(d("address"))
        On Error GoTo 0
        If Not r Is Nothing Then
            For Each f In formulas
                checks = checks + 1
                If checks > FM_MAX_MEMBERSHIP_CHECKS Then
                    FM_AddIssue "", "", "", "Range membership checks capped; direct reference list retained."
                    Exit Sub
                End If
                If f("sheet") = d("sheet") And f("book") = d("book") Then
                    Set c = Nothing
                    On Error Resume Next
                    Set c = Application.Intersect(r, mFMSource.Worksheets(f("sheet")).Range(f("address")))
                    On Error GoTo 0
                    If Not c Is Nothing Then FM_AddEdge f("id"), d("id"), "membership", "Formula cell contained in the referenced range"
                End If
            Next f
        End If
    Next d
End Sub

Private Function FM_CreateReport() As Workbook
    Dim result As Workbook, overview As Worksheet, detail As Worksheet, refs As Worksheet, issues As Worksheet
    Dim arr() As Variant, i As Long, e As Object, a As Object, b As Object, d As Object
    On Error GoTo Failed
    Set result = Workbooks.Add(xlWBATWorksheet)
    Set overview = result.Worksheets(1): overview.Name = "Worksheet relationships"
    Set detail = result.Worksheets.Add(After:=overview): detail.Name = "Cell relationships"
    Set refs = result.Worksheets.Add(After:=detail): refs.Name = "Formula references"
    Set issues = result.Worksheets.Add(After:=refs): issues.Name = "Review notices"
    FM_StyleSheet overview: FM_StyleSheet detail: FM_StyleSheet refs: FM_StyleSheet issues
    overview.Range("B2").Value2 = "Worksheet relationships"
    overview.Range("B4").Value2 = "Workbook: " & mFMSource.Name & "    Formulas analyzed: " & mFMFormulaCount & " / " & Format$(mFMTotalFormulas, "#,##0")
    overview.Range("B5").Value2 = "Source → dependent"
    overview.Range("B6").Value2 = "Connections and boxes are editable. Use the generated HTML for search and full exploration."
    If mFMLimited Then overview.Range("B7").Value2 = "Partial analysis: the 50,000-formula limit was reached."
    FM_DrawSheetMap overview
    detail.Range("B2").Value2 = "Cell relationships"
    detail.Range("B4").Value2 = "Solid: direct reference / Dashed: formula cell contained in a referenced range"
    detail.Range("B5").Value2 = "Up to " & FM_MAX_NATIVE_NODES & " cells/ranges are shown. See Formula references and the HTML for all references."
    detail.Range("B6").Value2 = "Each box shows Sheet!Address, nearby labels, and its formula. Labels are copied with their original addresses."
    FM_DrawCellMap detail
    refs.Range("B2").Value2 = "Formula references"
    refs.Range("B4").Value2 = "Captured: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "    Source: " & mFMSource.Name
    refs.Range("A6:K6").Value2 = Array("Source workbook", "Source worksheet", "Source cell/range", "Dependent worksheet", "Dependent cell", "Relationship type", "Relationship description", "English description", "Dependent formula", "Source labels (original nearby text)", "Note")
    If mFMEdges.Count > 0 Then
        ReDim arr(1 To mFMEdges.Count, 1 To 11)
        i = 0
        For Each e In mFMEdges
            i = i + 1: Set a = mFMNodes(CLng(Mid$(e("s"), 2))): Set b = mFMNodes(CLng(Mid$(e("t"), 2)))
            arr(i, 1) = a("book"): arr(i, 2) = a("sheet"): arr(i, 3) = a("address")
            arr(i, 4) = b("sheet"): arr(i, 5) = b("address"): arr(i, 6) = e("kind")
            If e("kind") = "membership" Then
                arr(i, 7) = FM_NodeText(a) & " is contained in range " & FM_NodeText(b) & "."
                arr(i, 8) = FM_NodeText(a) & " is contained in range " & FM_NodeText(b) & "."
            Else
                arr(i, 7) = FM_NodeText(b) & " references " & FM_NodeText(a) & "."
                arr(i, 8) = FM_NodeText(b) & " references " & FM_NodeText(a) & "."
            End If
            arr(i, 9) = b("formula"): arr(i, 10) = a("label"): arr(i, 11) = e("note")
        Next e
        refs.Range("A7").Resize(mFMEdges.Count, 11).NumberFormat = "@"
        refs.Range("A7").Resize(mFMEdges.Count, 11).Value2 = arr
    End If
    refs.Range("A6:K" & CStr(Application.Max(7, mFMEdges.Count + 6))).AutoFilter
    refs.Columns("A:F").ColumnWidth = 21: refs.Columns("G:I").ColumnWidth = 58: refs.Columns("J:K").ColumnWidth = 38
    refs.Rows(6).RowHeight = 32
    refs.Activate: refs.Range("D7").Select: ActiveWindow.FreezePanes = True
    issues.Range("B2").Value2 = "References to review"
    issues.Range("B4").Value2 = "This is a static analysis of formula references. Conditional formulas include references from every possible branch."
    issues.Range("B5").Value2 = "Review notices for references that depend on calculation, defined names, or table state. Hidden worksheets are also analyzed."
    issues.Range("A7:D7").Value2 = Array("Worksheet", " cells", "Formula", "Reason for review")
    If mFMIssues.Count > 0 Then
        ReDim arr(1 To mFMIssues.Count, 1 To 4): i = 0
        For Each d In mFMIssues
            i = i + 1: arr(i, 1) = d("sheet"): arr(i, 2) = d("address"): arr(i, 3) = d("formula"): arr(i, 4) = d("reason")
        Next d
        issues.Range("A8").Resize(mFMIssues.Count, 4).NumberFormat = "@"
        issues.Range("A8").Resize(mFMIssues.Count, 4).Value2 = arr
    Else
        issues.Range("A8").Value2 = "No specific review notices were detected. Formula calculation results were not validated."
    End If
    issues.Columns("A:B").ColumnWidth = 22: issues.Columns("C").ColumnWidth = 65: issues.Columns("D").ColumnWidth = 110
    issues.Range("A7:D" & CStr(Application.Max(8, mFMIssues.Count + 7))).AutoFilter
    FM_Header refs.Range("A6:K6"): FM_Header issues.Range("A7:D7")
    overview.Activate: ActiveWindow.DisplayGridlines = False: ActiveWindow.Zoom = 85
    Set FM_CreateReport = result
    Exit Function
Failed:
    Dim savedError As Long, savedDescription As String
    savedError = Err.Number: savedDescription = Err.Description
    On Error Resume Next
    If Not result Is Nothing Then result.Close False
    On Error GoTo 0
    Err.Raise savedError, "FM_CreateReport", savedDescription
End Function

Private Sub FM_StyleSheet(ByVal ws As Worksheet)
    With ws.Range("A1:Q50")
        .Font.Name = "Arial": .Font.Size = 10: .Font.Color = RGB(16, 38, 60)
        .RowHeight = 22
    End With
    ws.Columns("A").ColumnWidth = 3: ws.Columns("B:Q").ColumnWidth = 13
    ws.Range("B2").Font.Size = 16: ws.Range("B2").Font.Bold = True
    ws.Activate: ActiveWindow.DisplayGridlines = False
End Sub

Private Sub FM_Header(ByVal r As Range)
    r.Interior.Color = RGB(16, 38, 60): r.Font.Color = vbWhite: r.Font.Bold = True
    r.HorizontalAlignment = xlCenter: r.VerticalAlignment = xlCenter
End Sub

Private Function FM_NodeText(ByVal d As Object) As String
    If d("book") <> mFMSource.Name Then FM_NodeText = "[" & d("book") & "]"
    FM_NodeText = FM_NodeText & "'" & Replace$(d("sheet"), "'", "''") & "'!" & d("address")
End Function

Private Function FM_Box(ByVal ws As Worksheet, ByVal text As String, ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single, ByVal accent As Long) As Shape
    Dim s As Shape
    Set s = ws.Shapes.AddShape(5, x, y, width, height)
    s.Fill.ForeColor.RGB = RGB(245, 248, 250): s.Line.ForeColor.RGB = accent
    s.Line.Weight = 1.5
    With s.TextFrame2
        .TextRange.Text = text: .TextRange.Font.Name = "Arial": .TextRange.Font.Size = 10
        .TextRange.Font.Fill.ForeColor.RGB = RGB(16, 38, 60)
        .MarginLeft = 10: .MarginRight = 8: .MarginTop = 8: .MarginBottom = 6
        .WordWrap = -1
    End With
    Set FM_Box = s
End Function

Private Sub FM_Connect(ByVal ws As Worksheet, ByVal a As Shape, ByVal b As Shape, ByVal dashed As Boolean)
    Dim line As Shape, sourceSite As Long, targetSite As Long
    If a.Left < b.Left Then
        sourceSite = 4: targetSite = 2
    ElseIf a.Left > b.Left Then
        sourceSite = 2: targetSite = 4
    ElseIf a.Top < b.Top Then
        sourceSite = 3: targetSite = 1
    Else
        sourceSite = 1: targetSite = 3
    End If
    Set line = ws.Shapes.AddConnector(2, 0, 0, 10, 10)
    line.ConnectorFormat.BeginConnect a, sourceSite: line.ConnectorFormat.EndConnect b, targetSite
    line.Line.ForeColor.RGB = RGB(0, 127, 140): line.Line.EndArrowheadStyle = 3
    line.Line.Weight = 1.2
    If dashed Then line.Line.DashStyle = 4
    line.ZOrder 1
End Sub

Private Sub FM_DrawSheetMap(ByVal ws As Worksheet)
    Dim groups As Object, shapes As Object, counts As Object, e As Object, a As Object, b As Object
    Dim key As Variant, parts As Variant, n As Long, label As String, s As Shape
    Set groups = CreateObject("Scripting.Dictionary"): Set shapes = CreateObject("Scripting.Dictionary"): Set counts = CreateObject("Scripting.Dictionary")
    For Each key In mFMSheets: groups.Add mFMSource.Name & Chr$(30) & CStr(key), CStr(key): Next key
    For Each e In mFMEdges
        If e("kind") <> "membership" Then
            Set a = mFMNodes(CLng(Mid$(e("s"), 2))): Set b = mFMNodes(CLng(Mid$(e("t"), 2)))
            key = a("book") & Chr$(30) & a("sheet")
            If Not groups.Exists(key) Then groups.Add key, "[" & a("book") & "] " & a("sheet")
            If a("book") <> b("book") Or a("sheet") <> b("sheet") Then
                key = a("book") & Chr$(30) & a("sheet") & Chr$(29) & b("book") & Chr$(30) & b("sheet")
                If Not counts.Exists(key) Then counts.Add key, 0
                counts(key) = counts(key) + 1
            End If
        End If
    Next e
    n = 0
    For Each key In groups.Keys
        n = n + 1
        If n > 60 Then Exit For
        Set s = FM_Box(ws, groups(key), 30 + ((n - 1) Mod 3) * 310, 180 + ((n - 1) \ 3) * 160, 230, 70, RGB(0, 127, 140))
        shapes.Add key, s
    Next key
    For Each key In counts.Keys
        parts = Split(key, Chr$(29))
        If shapes.Exists(parts(0)) And shapes.Exists(parts(1)) Then FM_Connect ws, shapes(parts(0)), shapes(parts(1)), False
    Next key
    If groups.Count > 60 Then ws.Range("B8").Value2 = "Up to 60 worksheets are shown. See the HTML for the complete list."
    ws.Range("S2:U2").Value2 = Array("Source", "Dependent", "Direct references")
    ws.Columns("S:T").NumberFormat = "@"
    n = 2
    For Each key In counts.Keys
        n = n + 1: parts = Split(key, Chr$(29))
        ws.Cells(n, 19).Value2 = groups(parts(0)): ws.Cells(n, 20).Value2 = groups(parts(1)): ws.Cells(n, 21).Value2 = counts(key)
    Next key
    ws.Columns("S:T").ColumnWidth = 28: ws.Columns("U").ColumnWidth = 18
    FM_Header ws.Range("S2:U2")
End Sub

Private Sub FM_DrawCellMap(ByVal ws As Worksheet)
    Dim groups As Object, shapes As Object, rows As Object, d As Object, e As Object, s As Shape
    Dim key As String, n As Long, col As Long, r As Long, t As String, hidden As Long
    Set groups = CreateObject("Scripting.Dictionary"): Set shapes = CreateObject("Scripting.Dictionary"): Set rows = CreateObject("Scripting.Dictionary")
    For Each d In mFMNodes
        n = n + 1
        If n > FM_MAX_NATIVE_NODES Then Exit For
        key = d("book") & Chr$(30) & d("sheet")
        If Not groups.Exists(key) Then groups.Add key, groups.Count: rows.Add key, 0
        col = groups(key): r = rows(key): rows(key) = r + 1
        t = FM_NodeText(d)
        If Len(d("label")) > 0 Then t = t & vbCrLf & d("label")
        If Len(d("formula")) > 0 Then t = t & vbCrLf & Left$(d("formula"), 110)
        If d("kind") = "range" Then t = t & vbCrLf & "Referenced range"
        Set s = FM_Box(ws, t, 30 + col * 290, 170 + r * 120, 250, 96, FM_SheetColor(col))
        shapes.Add d("id"), s
    Next d
    n = 0: hidden = 0
    For Each e In mFMEdges
        If shapes.Exists(e("s")) And shapes.Exists(e("t")) And n < 200 Then
            FM_Connect ws, shapes(e("s")), shapes(e("t")), e("kind") = "membership": n = n + 1
        Else
            hidden = hidden + 1
        End If
    Next e
    ws.Range("B7").Value2 = "Shown: " & shapes.Count & " / " & mFMNodes.Count & " boxes, " & n & " / " & mFMEdges.Count & " connections. All collected data remains in Formula references and the HTML."
End Sub

Private Function FM_SheetColor(ByVal n As Long) As Long
    Select Case n Mod 5
        Case 0: FM_SheetColor = RGB(0, 127, 140)
        Case 1: FM_SheetColor = RGB(40, 93, 158)
        Case 2: FM_SheetColor = RGB(121, 77, 152)
        Case 3: FM_SheetColor = RGB(166, 99, 25)
        Case Else: FM_SheetColor = RGB(47, 126, 80)
    End Select
End Function

Private Function FM_DataJSON() As String
    Dim nodeItems() As String, edgeItems() As String, issueItems() As String, sheetItems() As String
    Dim i As Long, d As Object, nodeText As String, edgeText As String, issueText As String, sheetText As String
    If mFMNodes.Count > 0 Then
        ReDim nodeItems(0 To mFMNodes.Count - 1): i = 0
        For Each d In mFMNodes
            nodeItems(i) = FM_ObjectJSON(d, Array("id", "book", "sheet", "address", "label", "value", "formula", "kind"))
            If d.Exists("r1") Then nodeItems(i) = Left$(nodeItems(i), Len(nodeItems(i)) - 1) & "," & FM_GeometryJSON(d) & "}"
            i = i + 1
        Next d
        nodeText = Join(nodeItems, ",")
    End If
    If mFMEdges.Count > 0 Then
        ReDim edgeItems(0 To mFMEdges.Count - 1): i = 0
        For Each d In mFMEdges
            edgeItems(i) = FM_ObjectJSON(d, Array("s", "t", "kind", "note")): i = i + 1
        Next d
        edgeText = Join(edgeItems, ",")
    End If
    If mFMIssues.Count > 0 Then
        ReDim issueItems(0 To mFMIssues.Count - 1): i = 0
        For Each d In mFMIssues
            issueItems(i) = FM_ObjectJSON(d, Array("sheet", "address", "formula", "reason")): i = i + 1
        Next d
        issueText = Join(issueItems, ",")
    End If
    If mFMSheets.Count > 0 Then
        ReDim sheetItems(0 To mFMSheets.Count - 1)
        For i = 1 To mFMSheets.Count: sheetItems(i - 1) = FM_JSON(CStr(mFMSheets(i))): Next i
        sheetText = Join(sheetItems, ",")
    End If
    FM_DataJSON = "{""sourceName"":" & FM_JSON(mFMSource.Name) & ",""analyzedAt"":" & FM_JSON(Format$(Now, "yyyy-mm-dd hh:nn:ss")) & _
        ",""formulaCount"":" & mFMFormulaCount & ",""totalFormulaCount"":" & Format$(mFMTotalFormulas, "0") & _
        ",""scanLimited"":" & LCase$(CStr(mFMLimited)) & ",""nodes"": [" & nodeText & "],""edges"": [" & edgeText & _
        "],""issues"": [" & issueText & "],""sheets"": [" & sheetText & "],""worksheets"":" & mFMWorksheetJSON & "}"
End Function

Private Function FM_ObjectJSON(ByVal d As Object, ByVal keys As Variant) As String
    Dim parts() As String, i As Long
    ReDim parts(LBound(keys) To UBound(keys))
    For i = LBound(keys) To UBound(keys): parts(i) = FM_JSON(CStr(keys(i))) & ":" & FM_JSON(CStr(d(keys(i)))): Next i
    FM_ObjectJSON = "{" & Join(parts, ",") & "}"
End Function

Private Function FM_JSON(ByVal s As String) As String
    Dim i As Long, n As Long, c As String, items() As String
    If Len(s) = 0 Then FM_JSON = """""": Exit Function
    ReDim items(1 To Len(s))
    For i = 1 To Len(s)
        c = Mid$(s, i, 1): n = AscW(c)
        Select Case n
            Case 34: items(i) = "\"""
            Case 92: items(i) = "\\"
            Case 0 To 31, 60, 62, 38, &H2028, &H2029
                items(i) = "\u" & Right$("0000" & Hex$(n), 4)
            Case Else: items(i) = c
        End Select
    Next i
    FM_JSON = """" & Join(items, "") & """"
End Function

Private Sub FM_WriteUTF8(ByVal path As String, ByVal text As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.WriteText text: stream.SaveToFile path, 1: stream.Close
End Sub

Private Function FM_Base64ToText(ByVal encoded As String) As String
    Dim doc As Object, element As Object, stream As Object
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    Set element = doc.createElement("base64")
    element.DataType = "bin.base64": element.Text = encoded
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1: stream.Open: stream.Write element.nodeTypedValue
    stream.Position = 0: stream.Type = 2: stream.Charset = "utf-8"
    FM_Base64ToText = stream.ReadText: stream.Close
End Function

Private Function FM_SafeName(ByVal name As String) As String
    Dim c As Variant
    For Each c In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        name = Replace$(name, CStr(c), "_")
    Next c
    FM_SafeName = Left$(name, 80)
End Function

Private Function FM_UniqueBase(ByVal base As String) As String
    Dim suffix As Long, candidate As String
    candidate = base
    Do While Len(Dir$(candidate & ".xlsx")) > 0 Or Len(Dir$(candidate & ".html")) > 0
        suffix = suffix + 1: candidate = base & "_" & suffix
    Loop
    FM_UniqueBase = candidate
End Function
