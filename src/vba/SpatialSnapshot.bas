' Read-only worksheet context and exact Excel point-to-CSS-pixel geometry.
Private Function FM_Num(ByVal value As Double) As String
    FM_Num = Trim$(Str$(value))
End Function

Private Function FM_GeometryJSON(ByVal d As Object) As String
    Dim fields As Variant, pieces() As String, i As Long
    fields = Array("r1", "c1", "r2", "c2", "x", "y", "w", "h")
    ReDim pieces(0 To UBound(fields))
    For i = 0 To UBound(fields): pieces(i) = FM_JSON(CStr(fields(i))) & ":" & FM_Num(CDbl(d(fields(i)))): Next i
    FM_GeometryJSON = Join(pieces, ",")
End Function

Private Function FM_SnapshotJSON() As String
    Dim ws As Worksheet, snapshots() As String, i As Long
    ReDim snapshots(0 To mFMSource.Worksheets.Count - 1)
    For Each ws In mFMSource.Worksheets
        snapshots(i) = FM_OneSnapshotJSON(ws): i = i + 1
        DoEvents
    Next ws
    FM_SnapshotJSON = "[" & Join(snapshots, ",") & "]"
End Function

Private Function FM_OneSnapshotJSON(ByVal ws As Worksheet) As String
    Const MAX_GEOMETRY_ROWS As Long = 2000
    Const MAX_GEOMETRY_COLS As Long = 256
    Const MAX_CONTEXT_CELLS As Long = 20000
    Dim used As Range, rEnd As Long, cEnd As Long, gRows As Long, gCols As Long
    Dim rows() As String, cols() As String, captures As Object, merges As Object
    Dim rr As Long, cc As Long, d As Object, key As Variant, parts As Variant
    Dim cell As Range, merged As Range, texts As Collection, t As String
    Dim r As Long, c As Long, r1 As Long, r2 As Long, c1 As Long, c2 As Long
    Dim cellTexts() As String, mergeTexts() As String, count As Long, i As Long, limited As Boolean
    Dim color As Long, fontColor As Long, bold As Boolean, f As String, value As Variant
    Dim fullWidth As Double, fullHeight As Double, width As Double, height As Double
    Set used = ws.UsedRange
    rEnd = used.Row + used.Rows.Count - 1: cEnd = used.Column + used.Columns.Count - 1
    gRows = Application.Min(MAX_GEOMETRY_ROWS, rEnd): gCols = Application.Min(MAX_GEOMETRY_COLS, cEnd)
    fullWidth = (ws.Cells(1, cEnd).Left + ws.Cells(1, cEnd).Width) * 4 / 3
    fullHeight = (ws.Cells(rEnd, 1).Top + ws.Cells(rEnd, 1).Height) * 4 / 3
    width = (ws.Cells(1, gCols).Left + ws.Cells(1, gCols).Width) * 4 / 3
    height = (ws.Cells(gRows, 1).Top + ws.Cells(gRows, 1).Height) * 4 / 3
    ReDim rows(0 To gRows - 1): ReDim cols(0 To gCols - 1)
    For r = 1 To gRows
        rows(r - 1) = "{""index"":" & r & ",""height"":" & FM_Num(ws.Rows(r).Height * 4 / 3) & ",""hidden"":" & LCase$(CStr(ws.Rows(r).Hidden)) & "}"
    Next r
    For c = 1 To gCols
        cols(c - 1) = "{""index"":" & c & ",""width"":" & FM_Num(ws.Columns(c).Width * 4 / 3) & ",""hidden"":" & LCase$(CStr(ws.Columns(c).Hidden)) & "}"
    Next c
    Set captures = CreateObject("Scripting.Dictionary"): Set merges = CreateObject("Scripting.Dictionary")
    For rr = 1 To Application.Min(gRows, 100)
        For cc = 1 To Application.Min(gCols, 40)
            captures(CStr(rr) & ":" & CStr(cc)) = True
        Next cc
    Next rr
    ' Capture context around distant referenced cells without moving them closer.
    For Each d In mFMNodes
        If d("book") = mFMSource.Name And d("sheet") = ws.Name And d.Exists("r1") Then
            r1 = Application.Max(1, CLng(d("r1")) - 3): c1 = Application.Max(1, CLng(d("c1")) - 2)
            r2 = Application.Min(gRows, CLng(d("r1")) + 3): c2 = Application.Min(gCols, CLng(d("c1")) + 3)
            For rr = r1 To r2
                For cc = c1 To c2
                    If captures.Count < MAX_CONTEXT_CELLS Then
                        captures(CStr(rr) & ":" & CStr(cc)) = True
                    ElseIf Not captures.Exists(CStr(rr) & ":" & CStr(cc)) Then
                        limited = True
                    End If
                Next cc
            Next rr
        End If
    Next d
    limited = limited Or rEnd > 100 Or cEnd > 40 Or rEnd > gRows Or cEnd > gCols
    Set texts = New Collection
    For Each key In captures.Keys
        parts = Split(CStr(key), ":"): r = CLng(parts(0)): c = CLng(parts(1))
        Set cell = ws.Cells(r, c)
        value = cell.Value2
        If IsError(value) Or cell.HasFormula Or Not IsEmpty(value) Then
            Set merged = cell
            If cell.MergeCells Then
                Set merged = cell.MergeArea
                If Not merges.Exists(merged.Address) Then
                    merges.Add merged.Address, "{""r1"":" & merged.Row & ",""c1"":" & merged.Column & _
                        ",""r2"":" & (merged.Row + merged.Rows.Count - 1) & ",""c2"":" & (merged.Column + merged.Columns.Count - 1) & _
                        ",""x"":" & FM_Num(merged.Left * 4 / 3) & ",""y"":" & FM_Num(merged.Top * 4 / 3) & _
                        ",""w"":" & FM_Num(merged.Width * 4 / 3) & ",""h"":" & FM_Num(merged.Height * 4 / 3) & "}"
                End If
            End If
            color = cell.Interior.Color: fontColor = cell.Font.Color: bold = cell.Font.Bold
            On Error Resume Next
            color = cell.DisplayFormat.Interior.Color
            fontColor = cell.DisplayFormat.Font.Color
            bold = cell.DisplayFormat.Font.Bold
            On Error GoTo 0
            f = "": If cell.HasFormula Then f = FM_GetFormula(cell)
            t = "{""r"":" & r & ",""c"":" & c & ",""address"":" & FM_JSON(cell.Address(False, False)) & _
                ",""text"":" & FM_JSON(CStr(cell.Text)) & ",""formula"":" & FM_JSON(f) & _
                ",""fill"":" & FM_JSON(FM_ColorHex(color)) & ",""fontColor"":" & FM_JSON(FM_ColorHex(fontColor)) & _
                ",""bold"":" & LCase$(CStr(bold)) & ",""format"":" & FM_JSON(CStr(cell.NumberFormat)) & _
                ",""x"":" & FM_Num(merged.Left * 4 / 3) & ",""y"":" & FM_Num(merged.Top * 4 / 3) & _
                ",""w"":" & FM_Num(merged.Width * 4 / 3) & ",""h"":" & FM_Num(merged.Height * 4 / 3) & "}"
            texts.Add t
        End If
    Next key
    If texts.Count > 0 Then
        ReDim cellTexts(0 To texts.Count - 1)
        For i = 1 To texts.Count: cellTexts(i - 1) = texts(i): Next i
        t = Join(cellTexts, ",")
    Else
        t = ""
    End If
    Dim mergeJSON As String
    If merges.Count > 0 Then
        ReDim mergeTexts(0 To merges.Count - 1): i = 0
        For Each key In merges.Keys: mergeTexts(i) = merges(key): i = i + 1: Next key
        mergeJSON = Join(mergeTexts, ",")
    End If
    FM_OneSnapshotJSON = "{""name"":" & FM_JSON(ws.Name) & ",""book"":" & FM_JSON(mFMSource.Name) & _
        ",""used"":{""r1"":" & used.Row & ",""c1"":" & used.Column & ",""r2"":" & rEnd & ",""c2"":" & cEnd & "}" & _
        ",""size"":{""width"":" & FM_Num(width) & ",""height"":" & FM_Num(height) & "}" & _
        ",""fullSize"":{""width"":" & FM_Num(fullWidth) & ",""height"":" & FM_Num(fullHeight) & "}" & _
        ",""columns"": [" & Join(cols, ",") & "],""rows"": [" & Join(rows, ",") & "],""merges"": [" & mergeJSON & _
        "],""cells"": [" & t & "],""snapshotLimited"":" & LCase$(CStr(limited)) & ",""hidden"":" & LCase$(CStr(ws.Visible <> xlSheetVisible)) & "}"
End Function

Private Function FM_ColorHex(ByVal color As Long) As String
    If color < 0 Then color = 0
    FM_ColorHex = "#" & Right$("0" & Hex$(color And 255), 2) & _
        Right$("0" & Hex$((color \ 256) And 255), 2) & Right$("0" & Hex$((color \ 65536) And 255), 2)
End Function
