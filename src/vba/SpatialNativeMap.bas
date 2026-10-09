' Append to FormulaMap: an offline, native Excel worksheet-position map.
' The source is read only. All shapes and hyperlinks are created in the report.
Private Sub FM_AddNativeSpatialMap(ByVal report As Workbook)
    Const MAX_SHEETS As Long = 20
    Const MAX_CONTEXT_PER_SHEET As Long = 180
    Const MAX_NODES As Long = 700
    Const MAX_EDGES As Long = 120
    Const MAX_ROWS As Long = 2000
    Const MAX_COLS As Long = 256
    Const CARD_W As Double = 620
    Const CARD_H As Double = 390
    Const MAP_W As Double = 580
    Const MAP_H As Double = 322
    Dim map As Worksheet, sourceSheet As Worksheet, used As Range, view As Range
    Dim constants As Range, formulas As Range, context As Range, cell As Range, display As Range
    Dim nodeShapes As Object, transforms As Object, seen As Object, d As Object, e As Object
    Dim node As Shape, connector As Shape, frame As Shape, tr As Variant
    Dim sheetIndex As Long, lastRow As Long, lastCol As Long, shown As Long, n As Long
    Dim omitted As Long, clipped As Long, countContext As Long, contextOmitted As Long
    Dim nodeCount As Long, edgeCount As Long, eligibleEdges As Long
    Dim cardLeft As Double, cardTop As Double, originX As Double, originY As Double
    Dim factor As Double, sourceW As Double, sourceH As Double
    Dim x As Double, y As Double, w As Double, h As Double, r As Long, c As Long
    Dim text As String, note As String, key As String, color As Long
    Dim fontSize As Double, fontColor As Long, isBold As Boolean, fontValue As Variant
    If mFMSource Is Nothing Then Exit Sub
    Set map = report.Worksheets.Add(Before:=report.Worksheets(1))
    map.Name = "Cell position map"
    map.Range("A1:AR400").Font.Name = "Arial": map.Range("A1:AR400").Font.Size = 10
    map.Range("B2").Value2 = "Formula relationships in worksheet positions"
    map.Range("B2").Font.Size = 19: map.Range("B2").Font.Bold = True
    map.Range("B4").Value2 = "Blue = formula / Orange = input / Teal = referenced range. Click a cell or range to open its source position."
    map.Range("B5").Value2 = "This miniature map preserves the A1 origin, blank space, row heights, and column widths. Click an item to read its details."
    map.Range("B6").Value2 = "Display limits: 20 worksheets, 180 context cells per sheet, 700 referenced items, and 120 direct connections. Geometry is limited to A1:IV2000; external workbook references are excluded."
    map.Range("B7").Value2 = "Source: " & mFMSource.FullName
    map.Columns("A:A").ColumnWidth = 3
    map.Columns("B:AR").ColumnWidth = 4
    Set nodeShapes = CreateObject("Scripting.Dictionary")
    Set transforms = CreateObject("Scripting.Dictionary")
    For Each sourceSheet In mFMSource.Worksheets
        If sheetIndex >= MAX_SHEETS Then Exit For
        sheetIndex = sheetIndex + 1
        cardLeft = 24 + ((sheetIndex - 1) Mod 2) * (CARD_W + 28)
        cardTop = 145 + ((sheetIndex - 1) \ 2) * (CARD_H + 26)
        Set used = sourceSheet.UsedRange
        lastRow = Application.Min(MAX_ROWS, used.Row + used.Rows.Count - 1)
        lastCol = Application.Min(MAX_COLS, used.Column + used.Columns.Count - 1)
        If used.Row + used.Rows.Count - 1 > MAX_ROWS Or used.Column + used.Columns.Count - 1 > MAX_COLS Then clipped = clipped + 1
        Set view = sourceSheet.Range(sourceSheet.Cells(1, 1), sourceSheet.Cells(lastRow, lastCol))
        sourceW = sourceSheet.Cells(1, lastCol).Left + sourceSheet.Cells(1, lastCol).Width
        sourceH = sourceSheet.Cells(lastRow, 1).Top + sourceSheet.Cells(lastRow, 1).Height
        If sourceW <= 0 Then sourceW = 1
        If sourceH <= 0 Then sourceH = 1
        factor = Application.Min(MAP_W / sourceW, MAP_H / sourceH)
        originX = cardLeft + 28: originY = cardTop + 49
        transforms.Add sourceSheet.Name, Array(originX, originY, factor, lastRow, lastCol)
        Set frame = map.Shapes.AddShape(msoShapeRectangle, cardLeft, cardTop, CARD_W, CARD_H)
        frame.Fill.Visible = msoFalse: frame.Line.ForeColor.RGB = RGB(195, 208, 215)
        frame.Line.Weight = 0.8
        Set node = FMNS_Text(map, sourceSheet.Name, cardLeft + 14, cardTop + 8, CARD_W - 28, 20, 13, True, RGB(22, 53, 70))
        FMNS_Link map, node, sourceSheet.Name, "A1", "Open source worksheet"
        note = "A1:" & sourceSheet.Cells(lastRow, lastCol).Address(False, False) & "  / Actual-position scale, " & Format$(factor * 100, "0") & "%"
        If sourceSheet.Visible <> xlSheetVisible Then note = note & " / Hidden worksheet: unhide in Excel before opening"
        If used.Row + used.Rows.Count - 1 > MAX_ROWS Or used.Column + used.Columns.Count - 1 > MAX_COLS Then note = note & "  / Partial range omitted"
        Set node = FMNS_Text(map, note, cardLeft + 14, cardTop + 30, CARD_W - 28, 14, 8, False, RGB(93, 113, 125))
        ' Sparse row/column labels anchor the miniatures to actual Excel coordinates.
        For c = 1 To lastCol
            If c = 1 Or c Mod 4 = 0 Or c = lastCol Then
                x = originX + sourceSheet.Cells(1, c).Left * factor
                If sourceSheet.Cells(1, c).Width > 0 Then Set node = FMNS_Text(map, Split(sourceSheet.Cells(1, c).Address(True, False), "$", 2)(0), x, originY - 13, 24, 12, 6, False, RGB(111, 128, 138))
            End If
        Next c
        For r = 1 To lastRow
            If r = 1 Or r Mod Application.Max(10, CLng(lastRow / 10)) = 0 Or r = lastRow Then
                y = originY + sourceSheet.Cells(r, 1).Top * factor
                If sourceSheet.Cells(r, 1).Height > 0 Then Set node = FMNS_Text(map, CStr(r), cardLeft + 5, y, 21, 10, 6, False, RGB(111, 128, 138))
            End If
        Next r
        ' Enumerate populated cells, rather than scanning a potentially huge empty grid.
        Set constants = Nothing: Set formulas = Nothing: Set context = Nothing
        On Error Resume Next
        Set constants = view.SpecialCells(xlCellTypeConstants)
        Set formulas = view.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
        If Not constants Is Nothing Then Set context = Application.Intersect(view, constants)
        If Not formulas Is Nothing Then
            If context Is Nothing Then
                Set context = Application.Intersect(view, formulas)
            Else
                Set context = Application.Union(context, Application.Intersect(view, formulas))
            End If
        End If
        Set seen = CreateObject("Scripting.Dictionary"): countContext = 0
        If Not context Is Nothing Then
            For Each cell In context.Cells
                Set display = cell
                If cell.MergeCells Then Set display = cell.MergeArea
                key = display.Address
                If Not seen.Exists(key) Then
                    seen.Add key, True
                    If countContext >= MAX_CONTEXT_PER_SHEET Then
                        contextOmitted = contextOmitted + Application.Max(0, context.CountLarge - countContext)
                        Exit For
                    ElseIf display.Left < sourceW And display.Top < sourceH And display.Width > 0 And display.Height > 0 Then
                        countContext = countContext + 1
                        text = Left$(Replace$(Replace$(FM_CellText(display.Cells(1, 1)), vbCr, " "), vbLf, " "), 45)
                        x = originX + display.Left * factor: y = originY + display.Top * factor
                        w = Application.Min(display.Width, sourceW - display.Left) * factor
                        h = Application.Min(display.Height, sourceH - display.Top) * factor
                        If w >= 1 And h >= 1 Then
                            fontSize = 10: fontColor = RGB(45, 63, 74): isBold = False
                            fontValue = cell.Font.Size
                            If Not IsNull(fontValue) Then
                                If IsNumeric(fontValue) Then fontSize = CDbl(fontValue)
                            End If
                            fontValue = cell.Font.Color
                            If Not IsNull(fontValue) Then
                                If IsNumeric(fontValue) Then
                                    If CLng(fontValue) >= 0 Then fontColor = CLng(fontValue)
                                End If
                            End If
                            fontValue = cell.Font.Bold
                            If Not IsNull(fontValue) Then isBold = CBool(fontValue)
                            Set node = FMNS_Text(map, text, x, y, w, h, Application.Min(7, Application.Max(4, fontSize * factor)), isBold, fontColor)
                            color = cell.Interior.Color
                            If color <> RGB(255, 255, 255) Then node.Fill.Visible = msoTrue: node.Fill.ForeColor.RGB = color
                            FMNS_Link map, node, sourceSheet.Name, display.Cells(1, 1).Address(False, False), text
                        End If
                    End If
                End If
            Next cell
        End If
        DoEvents
    Next sourceSheet
    ' Range boundaries are transparent; text beneath them remains recognizable.
    For Each d In mFMNodes
        If StrComp(CStr(d("book")), mFMSource.Name, vbTextCompare) = 0 And CStr(d("kind")) <> "external" Then
            If transforms.Exists(CStr(d("sheet"))) Then
                tr = transforms(CStr(d("sheet")))
                Set display = Nothing
                On Error Resume Next
                Set display = mFMSource.Worksheets(CStr(d("sheet"))).Range(CStr(d("address")))
                On Error GoTo 0
                If Not display Is Nothing Then
                    If display.CountLarge = 1 Then
                        If display.MergeCells Then Set display = display.MergeArea
                    End If
                    If display.Row <= CLng(tr(3)) And display.Column <= CLng(tr(4)) And display.Width > 0 And display.Height > 0 And nodeCount < MAX_NODES Then
                        sourceW = mFMSource.Worksheets(CStr(d("sheet"))).Cells(1, CLng(tr(4))).Left + mFMSource.Worksheets(CStr(d("sheet"))).Cells(1, CLng(tr(4))).Width
                        sourceH = mFMSource.Worksheets(CStr(d("sheet"))).Cells(CLng(tr(3)), 1).Top + mFMSource.Worksheets(CStr(d("sheet"))).Cells(CLng(tr(3)), 1).Height
                        x = CDbl(tr(0)) + display.Left * CDbl(tr(2)): y = CDbl(tr(1)) + display.Top * CDbl(tr(2))
                        w = Application.Max(2.5, Application.Min(display.Width, sourceW - display.Left) * CDbl(tr(2)))
                        h = Application.Max(2.5, Application.Min(display.Height, sourceH - display.Top) * CDbl(tr(2)))
                        color = RGB(214, 142, 29)
                        If CStr(d("kind")) = "formula" Then color = RGB(42, 112, 196)
                        If CStr(d("kind")) = "range" Then color = RGB(24, 137, 132)
                        Set node = map.Shapes.AddShape(msoShapeRectangle, x, y, w, h)
                        node.Fill.ForeColor.RGB = color: node.Fill.Transparency = 0.9
                        node.Line.ForeColor.RGB = color: node.Line.Weight = 0.8
                        node.AlternativeText = "'" & CStr(d("sheet")) & "'!" & CStr(d("address")) & " " & CStr(d("label"))
                        FMNS_Link map, node, CStr(d("sheet")), CStr(d("address")), CStr(d("label"))
                        nodeShapes.Add CStr(d("id")), node: nodeCount = nodeCount + 1
                    Else
                        omitted = omitted + 1
                    End If
                End If
            Else
                omitted = omitted + 1
            End If
        End If
    Next d
    ' Light arrows use the true locations; caps keep the static workbook usable.
    For Each e In mFMEdges
        If CStr(e("kind")) <> "membership" Then
            If nodeShapes.Exists(CStr(e("s"))) And nodeShapes.Exists(CStr(e("t"))) Then
                eligibleEdges = eligibleEdges + 1
                If edgeCount < MAX_EDGES Then
                    Set node = nodeShapes(CStr(e("s"))): Set frame = nodeShapes(CStr(e("t")))
                    Set connector = map.Shapes.AddConnector(msoConnectorStraight, node.Left + node.Width / 2, node.Top + node.Height / 2, frame.Left + frame.Width / 2, frame.Top + frame.Height / 2)
                    connector.Line.ForeColor.RGB = RGB(94, 131, 153): connector.Line.Transparency = 0.78
                    connector.Line.Weight = 0.65: connector.Line.EndArrowheadStyle = msoArrowheadTriangle
                    connector.ZOrder msoSendToBack
                    edgeCount = edgeCount + 1
                End If
            End If
        End If
    Next e
    map.Range("B9").Value2 = "Shown: " & sheetIndex & " / " & mFMSource.Worksheets.Count & " worksheets, referenced cells/ranges: " & nodeCount & ", direct connections: " & edgeCount & ". Omitted items: " & omitted & ", context cells: " & contextOmitted & " cells, connections: " & (eligibleEdges - edgeCount) & ", worksheets with clipped geometry: " & clipped & "."
    map.PageSetup.Orientation = xlLandscape
    map.PageSetup.Zoom = False: map.PageSetup.FitToPagesWide = 1: map.PageSetup.FitToPagesTall = False
    map.Activate
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False: ActiveWindow.Zoom = 75
    On Error GoTo 0
End Sub

Private Function FMNS_Text(ByVal ws As Worksheet, ByVal text As String, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal size As Double, ByVal bold As Boolean, ByVal color As Long) As Shape
    Dim shape As Shape
    Set shape = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, h)
    shape.Line.Visible = msoFalse: shape.Fill.Visible = msoFalse
    With shape.TextFrame2
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .WordWrap = msoFalse
        .TextRange.Text = text
        .TextRange.Font.Name = "Arial"
        .TextRange.Font.Size = size
        .TextRange.Font.Bold = IIf(bold, msoTrue, msoFalse)
        .TextRange.Font.Fill.ForeColor.RGB = color
    End With
    Set FMNS_Text = shape
End Function

Private Sub FMNS_Link(ByVal ws As Worksheet, ByVal shape As Shape, ByVal sheetName As String, ByVal address As String, ByVal label As String)
    Dim target As String, tip As String
    If Len(mFMSource.Path) = 0 Then Exit Sub
    target = "'" & Replace$(sheetName, "'", "''") & "'!" & address
    tip = Left$(target & "  " & label & "  (Open source cell)", 250)
    ws.Hyperlinks.Add Anchor:=shape, Address:=mFMSource.FullName, SubAddress:=target, ScreenTip:=tip
End Sub
