Attribute VB_Name = "FormulaMapParser"
Option Explicit

' Static, read-only A1 formula reference parser. No Evaluate/Calculate/Select.
' Public contract: Collection of dictionaries with book/sheet/address/kind/note.
' Address means a syntactically referenced source, not an evaluated dependency.

Public Function FM_ParseFormula(ByVal formulaCell As Range, ByVal issues As Collection) As Collection
    Dim result As New Collection, seen As Object, warned As Object
    Dim f As String, p As Long, q As Long, tokenEnd As Long, nextP As Long
    Dim qualifier As String, token As String, addr As String, refKind As String
    Dim structText As String, cellFormula As Variant, scopedFormula As Boolean
    Set seen = CreateObject("Scripting.Dictionary")
    Set warned = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    warned.CompareMode = vbTextCompare
    On Error GoTo ParseFailure
    If formulaCell Is Nothing Then GoTo Done
    If formulaCell.Cells.CountLarge <> 1 Then
        FMP_Issue issues, warned, "Parser requires one formula cell."
        GoTo Done
    End If
    If Not formulaCell.HasFormula Then GoTo Done
    On Error Resume Next
    cellFormula = CallByName(formulaCell, "Formula2", VbGet)
    If Err.Number <> 0 Then
        Err.Clear
        cellFormula = formulaCell.Formula
    End If
    On Error GoTo ParseFailure
    f = CStr(cellFormula)
    scopedFormula = FMP_HasScopedFunction(f)
    If scopedFormula Then FMP_Issue issues, warned, "LET/LAMBDA may shadow defined names; unqualified defined names are not resolved in this formula."
    p = 1
    Do While p <= Len(f)
        If Mid$(f, p, 1) = """" Then
            p = FMP_SkipString(f, p)
        ElseIf Mid$(f, p, 1) = "@" Then
            FMP_Issue issues, warned, "Implicit intersection (@): source references are shown; the effective intersected cell is not evaluated."
            p = p + 1
        ElseIf FMP_TryQualifier(f, p, qualifier, q) Then
            If FMP_ReadQualifiedA1(f, q, qualifier, addr, refKind, nextP) Then
                FMP_AddQualified result, seen, issues, warned, formulaCell, qualifier, addr, refKind
                If refKind = "spill" Then FMP_Issue issues, warned, "Spill reference (#): only its anchor/range is recorded; the spill extent is not evaluated."
                FMP_CheckIntersection f, nextP, issues, warned
                p = nextP
            Else
                token = FMP_ReadIdentifier(f, q, tokenEnd)
                If Len(token) > 0 Then
                    If FMP_ReadStructured(f, tokenEnd, structText, nextP) Then
                        FMP_Issue issues, warned, "Sheet-qualified structured reference could not be resolved safely: " & qualifier & "!" & token & structText
                        p = nextP
                    Else
                        FMP_AddQualifiedName result, seen, issues, warned, formulaCell, qualifier, token
                        p = tokenEnd
                    End If
                Else
                    FMP_Issue issues, warned, "Unresolved qualified reference near: " & Mid$(f, p, 70)
                    p = q
                End If
            End If
        ElseIf FMP_IsTokenStart(Mid$(f, p, 1)) And FMP_IsBoundaryBefore(f, p) Then
            token = FMP_ReadIdentifier(f, p, tokenEnd)
            q = FMP_SkipSpace(f, tokenEnd)
            If Len(token) > 0 And Mid$(f, q, 1) = "(" Then
                FMP_CheckFunction token, issues, warned
                p = tokenEnd
            ElseIf Len(token) > 0 And FMP_ReadStructured(f, tokenEnd, structText, nextP) Then
                FMP_AddStructured result, seen, issues, warned, formulaCell, token, structText
                FMP_CheckIntersection f, nextP, issues, warned
                p = nextP
            ElseIf FMP_ReadA1(f, p, addr, refKind, nextP) Then
                FMP_AddSource result, seen, formulaCell.Parent.Parent.Name, formulaCell.Parent.Name, addr, refKind, ""
                If refKind = "spill" Then FMP_Issue issues, warned, "Spill reference (#): only its anchor/range is recorded; the spill extent is not evaluated."
                FMP_CheckIntersection f, nextP, issues, warned
                p = nextP
            ElseIf Len(token) > 0 Then
                If Not scopedFormula Then FMP_AddName result, seen, issues, warned, formulaCell, formulaCell.Parent, token
                p = tokenEnd
            Else
                p = p + 1
            End If
        ElseIf Mid$(f, p, 1) = "[" Then
            If FMP_ReadStructured(f, p, structText, nextP) Then
                FMP_AddStructured result, seen, issues, warned, formulaCell, "", structText
                FMP_CheckIntersection f, nextP, issues, warned
                p = nextP
            Else
                p = p + 1
            End If
        Else
            p = p + 1
        End If
    Loop
Done:
    Set FM_ParseFormula = result
    Exit Function
ParseFailure:
    FMP_Issue issues, warned, "Formula parsing stopped after an error: " & Err.Description
    Resume Done
End Function

Private Sub FMP_CheckIntersection(ByVal f As String, ByVal p As Long, ByVal issues As Collection, ByVal warned As Object)
    Dim q As Long, ch As String
    q = FMP_SkipSpace(f, p)
    If q <= p Then Exit Sub
    ch = Mid$(f, q, 1)
    If FMP_IsTokenStart(ch) Or ch = "'" Or ch = "[" Or ch = "(" Then
        FMP_Issue issues, warned, "Possible reference intersection (space operator): operand ranges are shown; the intersected result is not evaluated."
    End If
End Sub

Private Function FMP_SkipString(ByVal f As String, ByVal p As Long) As Long
    p = p + 1
    Do While p <= Len(f)
        If Mid$(f, p, 1) = """" Then
            If Mid$(f, p + 1, 1) = """" Then
                p = p + 2
            Else
                FMP_SkipString = p + 1
                Exit Function
            End If
        Else
            p = p + 1
        End If
    Loop
    FMP_SkipString = p
End Function

Private Function FMP_SkipSpace(ByVal f As String, ByVal p As Long) As Long
    Do While p <= Len(f)
        If InStr(1, " " & vbTab & vbCr & vbLf, Mid$(f, p, 1), vbBinaryCompare) = 0 Then Exit Do
        p = p + 1
    Loop
    FMP_SkipSpace = p
End Function

Private Function FMP_IsTokenStart(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    FMP_IsTokenStart = FMP_IsIdentifierChar(ch) Or ch = "$"
End Function

Private Function FMP_IsIdentifierChar(ByVal ch As String) As Boolean
    Dim n As Long
    If Len(ch) = 0 Then Exit Function
    n = AscW(ch)
    FMP_IsIdentifierChar = (ch Like "[A-Za-z0-9_]") Or ch = "." Or ch = "\" Or n < 0 Or n > 127
End Function

Private Function FMP_IsBoundaryBefore(ByVal f As String, ByVal p As Long) As Boolean
    If p <= 1 Then
        FMP_IsBoundaryBefore = True
    Else
        FMP_IsBoundaryBefore = Not (FMP_IsIdentifierChar(Mid$(f, p - 1, 1)) Or Mid$(f, p - 1, 1) = "$")
    End If
End Function

Private Function FMP_ReadIdentifier(ByVal f As String, ByVal p As Long, ByRef endP As Long) As String
    Dim startP As Long
    startP = p
    Do While p <= Len(f)
        If Not FMP_IsIdentifierChar(Mid$(f, p, 1)) Then Exit Do
        p = p + 1
    Loop
    endP = p
    FMP_ReadIdentifier = Mid$(f, startP, p - startP)
End Function

Private Function FMP_TryQualifier(ByVal f As String, ByVal p As Long, ByRef qualifier As String, ByRef afterBang As Long) As Boolean
    Dim q As Long, ch As String, s As String, depth As Long
    If Not FMP_IsBoundaryBefore(f, p) Then Exit Function
    q = p
    If Mid$(f, q, 1) = "'" Then
        q = q + 1
        Do While q <= Len(f)
            ch = Mid$(f, q, 1)
            If ch = "'" Then
                If Mid$(f, q + 1, 1) = "'" Then
                    s = s & "'"
                    q = q + 2
                Else
                    If Mid$(f, q + 1, 1) = "!" Then
                        qualifier = s
                        afterBang = q + 2
                        FMP_TryQualifier = True
                    End If
                    Exit Function
                End If
            Else
                s = s & ch
                q = q + 1
            End If
        Loop
        Exit Function
    End If
    Do While q <= Len(f)
        ch = Mid$(f, q, 1)
        If ch = "!" And depth = 0 Then
            If q > p Then
                qualifier = Mid$(f, p, q - p)
                afterBang = q + 1
                FMP_TryQualifier = True
            End If
            Exit Function
        ElseIf ch = "[" Then
            depth = depth + 1
        ElseIf ch = "]" And depth > 0 Then
            depth = depth - 1
        ElseIf depth = 0 Then
            If Not (FMP_IsIdentifierChar(ch) Or ch = ":") Then Exit Function
        End If
        q = q + 1
    Loop
End Function

Private Function FMP_ReadA1(ByVal f As String, ByVal p As Long, ByRef address As String, ByRef kind As String, ByRef endP As Long) As Boolean
    Dim firstPart As String, secondPart As String, firstKind As String, secondKind As String
    Dim q As Long, r As Long, checkP As Long, ch As String
    If Not FMP_ReadA1Part(f, p, firstPart, firstKind, q) Then Exit Function
    address = firstPart
    kind = "cell"
    checkP = FMP_SkipSpace(f, q)
    If Mid$(f, checkP, 1) = ":" Then
        If Not FMP_ReadA1Part(f, FMP_SkipSpace(f, checkP + 1), secondPart, secondKind, r) Then Exit Function
        If firstKind <> secondKind Then Exit Function
        address = firstPart & ":" & secondPart
        kind = "range"
        q = r
    ElseIf firstKind <> "cell" Then
        Exit Function
    End If
    ch = Mid$(f, q, 1)
    If FMP_IsIdentifierChar(ch) Or ch = "$" Or ch = "!" Or ch = "[" Then Exit Function
    checkP = FMP_SkipSpace(f, q)
    If Mid$(f, checkP, 1) = "(" Then Exit Function
    If ch = "#" Then
        kind = "spill"
        q = q + 1
    End If
    endP = q
    FMP_ReadA1 = True
End Function

Private Function FMP_ReadQualifiedA1(ByVal f As String, ByVal p As Long, ByVal qualifier As String, ByRef address As String, ByRef kind As String, ByRef endP As Long) As Boolean
    Dim firstPart As String, firstKind As String, secondPart As String, secondKind As String
    Dim q As Long, r As Long, s As Long, qualifier2 As String
    If FMP_ReadA1(f, p, address, kind, endP) Then
        FMP_ReadQualifiedA1 = True
        Exit Function
    End If
    If Not FMP_ReadA1Part(f, p, firstPart, firstKind, q) Then Exit Function
    q = FMP_SkipSpace(f, q)
    If Mid$(f, q, 1) <> ":" Then Exit Function
    q = FMP_SkipSpace(f, q + 1)
    If Not FMP_TryQualifier(f, q, qualifier2, r) Then Exit Function
    If StrComp(qualifier, qualifier2, vbTextCompare) <> 0 Then Exit Function
    If Not FMP_ReadA1Part(f, r, secondPart, secondKind, s) Then Exit Function
    If firstKind <> secondKind Then Exit Function
    If FMP_IsIdentifierChar(Mid$(f, s, 1)) Or Mid$(f, s, 1) = "$" Then Exit Function
    address = firstPart & ":" & secondPart
    kind = "range"
    endP = s
    FMP_ReadQualifiedA1 = True
End Function

Private Function FMP_ReadA1Part(ByVal f As String, ByVal p As Long, ByRef part As String, ByRef kind As String, ByRef endP As Long) As Boolean
    Dim q As Long, startP As Long, col As String, rowText As String, colNum As Long, rowNum As Double, ch As String
    q = p
    If Mid$(f, q, 1) = "$" Then q = q + 1
    startP = q
    Do While q <= Len(f)
        ch = UCase$(Mid$(f, q, 1))
        If ch < "A" Or ch > "Z" Then Exit Do
        q = q + 1
    Loop
    col = UCase$(Mid$(f, startP, q - startP))
    If Len(col) > 3 Then Exit Function
    If Len(col) > 0 Then
        colNum = FMP_ColumnNumber(col)
        If colNum < 1 Or colNum > 16384 Then Exit Function
        If Mid$(f, q, 1) = "$" Then q = q + 1
    End If
    startP = q
    Do While Mid$(f, q, 1) Like "[0-9]" And q <= Len(f)
        q = q + 1
    Loop
    rowText = Mid$(f, startP, q - startP)
    If Len(rowText) > 0 Then
        If Len(rowText) > 7 Then Exit Function
        If Left$(rowText, 1) = "0" Then Exit Function
        rowNum = CDbl(rowText)
        If rowNum < 1 Or rowNum > 1048576 Then Exit Function
    End If
    If Len(col) > 0 And Len(rowText) > 0 Then
        kind = "cell": part = col & CStr(CLng(rowNum))
    ElseIf Len(col) > 0 Then
        kind = "column": part = col
    ElseIf Len(rowText) > 0 Then
        kind = "row": part = CStr(CLng(rowNum))
    Else
        Exit Function
    End If
    endP = q
    FMP_ReadA1Part = True
End Function

Private Function FMP_ColumnNumber(ByVal col As String) As Long
    Dim i As Long
    For i = 1 To Len(col)
        FMP_ColumnNumber = FMP_ColumnNumber * 26 + Asc(Mid$(col, i, 1)) - 64
    Next i
End Function

Private Sub FMP_CheckFunction(ByVal token As String, ByVal issues As Collection, ByVal warned As Object)
    Dim fn As String, p As Long
    fn = UCase$(token)
    p = InStrRev(fn, ".")
    If p > 0 Then fn = Mid$(fn, p + 1)
    Select Case fn
        Case "INDIRECT", "OFFSET"
            FMP_Issue issues, warned, fn & " uses a computed reference; only explicit references in its arguments are recorded."
        Case "LET", "LAMBDA"
            FMP_Issue issues, warned, fn & " scope/variables are not evaluated; explicit cell/range/table references are recorded."
    End Select
End Sub

Private Function FMP_HasScopedFunction(ByVal f As String) As Boolean
    Dim p As Long, q As Long, n As Long, token As String
    p = 1
    Do While p <= Len(f)
        If Mid$(f, p, 1) = """" Then
            p = FMP_SkipString(f, p)
        ElseIf FMP_IsIdentifierChar(Mid$(f, p, 1)) And FMP_IsBoundaryBefore(f, p) Then
            token = FMP_ReadIdentifier(f, p, q)
            n = InStrRev(token, ".")
            If n > 0 Then token = Mid$(token, n + 1)
            If Mid$(f, FMP_SkipSpace(f, q), 1) = "(" Then
                If UCase$(token) = "LET" Or UCase$(token) = "LAMBDA" Then
                    FMP_HasScopedFunction = True
                    Exit Function
                End If
            End If
            p = q
        Else
            p = p + 1
        End If
    Loop
End Function

Private Sub FMP_AddQualified(ByVal result As Collection, ByVal seen As Object, ByVal issues As Collection, ByVal warned As Object, ByVal source As Range, ByVal qualifier As String, ByVal address As String, ByVal kind As String)
    Dim bookName As String, sheetText As String, pathText As String, isExternal As Boolean
    Dim wb As Workbook, ws As Worksheet, firstIndex As Long, lastIndex As Long, i As Long, p As Long
    Set wb = source.Parent.Parent
    FMP_SplitQualifier qualifier, bookName, sheetText, pathText, isExternal
    If isExternal Then
        If StrComp(bookName, wb.Name, vbTextCompare) = 0 And Len(pathText) = 0 Then
            isExternal = False
        Else
            If InStr(sheetText, ":") > 0 Or Len(sheetText) = 0 Then
                FMP_Issue issues, warned, "External reference is not resolved: " & qualifier & "!" & address
            Else
                FMP_AddSource result, seen, pathText & bookName, sheetText, address, "external", "External workbook; existence/value is not verified. " & pathText
                FMP_Issue issues, warned, "External workbook reference is recorded as written; source existence is not verified: " & bookName
            End If
            Exit Sub
        End If
    End If
    p = InStr(sheetText, ":")
    If p > 0 Then
        firstIndex = FMP_SheetIndex(wb, Left$(sheetText, p - 1))
        lastIndex = FMP_SheetIndex(wb, Mid$(sheetText, p + 1))
        If firstIndex = 0 Or lastIndex = 0 Then
            FMP_Issue issues, warned, "3D reference has an unresolved sheet: " & qualifier & "!" & address
            Exit Sub
        End If
        If firstIndex > lastIndex Then
            i = firstIndex: firstIndex = lastIndex: lastIndex = i
        End If
        For i = firstIndex To lastIndex
            If TypeName(wb.Sheets(i)) = "Worksheet" Then
                FMP_AddSource result, seen, wb.Name, wb.Sheets(i).Name, address, "3d", "3D source: " & qualifier & "!" & address
            Else
                FMP_Issue issues, warned, "Non-worksheet in 3D sheet interval was skipped: " & wb.Sheets(i).Name
            End If
        Next i
    Else
        Set ws = FMP_FindSheet(wb, sheetText)
        If ws Is Nothing Then
            FMP_Issue issues, warned, "Referenced worksheet was not found: " & sheetText
        Else
            FMP_AddSource result, seen, wb.Name, ws.Name, address, kind, ""
        End If
    End If
End Sub

Private Sub FMP_SplitQualifier(ByVal qualifier As String, ByRef bookName As String, ByRef sheetText As String, ByRef pathText As String, ByRef external As Boolean)
    Dim leftP As Long, rightP As Long
    bookName = "": pathText = "": sheetText = qualifier: external = False
    leftP = InStrRev(qualifier, "[")
    rightP = InStrRev(qualifier, "]")
    If leftP > 0 And rightP > leftP Then
        bookName = Mid$(qualifier, leftP + 1, rightP - leftP - 1)
        pathText = Left$(qualifier, leftP - 1)
        sheetText = Mid$(qualifier, rightP + 1)
        external = True
    End If
End Sub

Private Function FMP_FindSheet(ByVal wb As Workbook, ByVal name As String) As Worksheet
    On Error Resume Next
    Set FMP_FindSheet = wb.Worksheets(name)
    On Error GoTo 0
End Function

Private Function FMP_SheetIndex(ByVal wb As Workbook, ByVal name As String) As Long
    Dim sh As Object
    On Error Resume Next
    Set sh = wb.Sheets(name)
    If Not sh Is Nothing Then FMP_SheetIndex = sh.Index
    On Error GoTo 0
End Function

Private Sub FMP_AddQualifiedName(ByVal result As Collection, ByVal seen As Object, ByVal issues As Collection, ByVal warned As Object, ByVal source As Range, ByVal qualifier As String, ByVal token As String)
    Dim bookName As String, sheetText As String, pathText As String, external As Boolean, ws As Worksheet
    FMP_SplitQualifier qualifier, bookName, sheetText, pathText, external
    If external Or InStr(sheetText, ":") > 0 Then
        FMP_Issue issues, warned, "External/3D defined name is not resolved: " & qualifier & "!" & token
        Exit Sub
    End If
    Set ws = FMP_FindSheet(source.Parent.Parent, sheetText)
    If ws Is Nothing Then
        FMP_Issue issues, warned, "Defined-name worksheet was not found: " & qualifier
    Else
        FMP_AddName result, seen, issues, warned, source, ws, token
    End If
End Sub

Private Sub FMP_AddName(ByVal result As Collection, ByVal seen As Object, ByVal issues As Collection, ByVal warned As Object, ByVal source As Range, ByVal scope As Worksheet, ByVal token As String)
    Dim nm As Name, candidate As Name, r As Range, n As String, p As Long, rf As String
    If Len(token) = 0 Then Exit Sub
    If IsNumeric(token) Then Exit Sub
    Select Case UCase$(token)
        Case "TRUE", "FALSE", "_XLFN", "_XLWS": Exit Sub
    End Select
    On Error Resume Next
    For Each candidate In scope.Names
        n = candidate.Name
        p = InStrRev(n, "!")
        If p > 0 Then n = Mid$(n, p + 1)
        If StrComp(n, token, vbTextCompare) = 0 Then
            Set nm = candidate
            Exit For
        End If
    Next candidate
    If nm Is Nothing Then
        For Each candidate In source.Parent.Parent.Names
            n = candidate.Name
            If InStr(n, "!") = 0 And StrComp(n, token, vbTextCompare) = 0 Then
                Set nm = candidate
                Exit For
            End If
        Next candidate
    End If
    On Error GoTo 0
    If nm Is Nothing Then
        If Not FMP_FindTable(source.Parent.Parent, token) Is Nothing Then
            FMP_AddStructured result, seen, issues, warned, source, token, "[#Data]"
        End If
        Exit Sub ' Unknown tokens may be LET/LAMBDA variables, not source cells.
    End If
    On Error Resume Next
    rf = nm.RefersTo
    On Error GoTo 0
    If FMP_IsConstantName(rf) Then Exit Sub
    If FMP_NameIsRelative(nm) Then
        FMP_Issue issues, warned, "Relative defined name is not resolved against an active cell: " & nm.Name
        Exit Sub
    End If
    On Error Resume Next
    Set r = nm.RefersToRange
    On Error GoTo 0
    If r Is Nothing Then
        If Not FMP_IsConstantName(rf) Then FMP_Issue issues, warned, "Defined name has no static Range (expression, dynamic or unavailable source): " & nm.Name
    Else
        FMP_AddRange result, seen, r, "name", "Defined name: " & nm.Name
    End If
End Sub

Private Function FMP_NameIsRelative(ByVal nm As Name) As Boolean
    Dim rx As Object, matches As Object, m As Object, s As String
    On Error GoTo Conservative
    s = FMP_StripQualifierAndStrings(nm.RefersToR1C1)
    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True: rx.IgnoreCase = True
    rx.Pattern = "(^|[^A-Za-z0-9_])R(\[[-+]?\d+\]|\d*)C(\[[-+]?\d+\]|\d*)"
    Set matches = rx.Execute(s)
    For Each m In matches
        If InStr(m.Value, "[") > 0 Or Len(m.SubMatches(1)) = 0 Or Len(m.SubMatches(2)) = 0 Then
            FMP_NameIsRelative = True
            Exit Function
        End If
    Next m
    rx.Pattern = "(^|[^A-Za-z0-9_])([RC])(\[[-+]?\d+\]|\d*)(?=$|[^A-Za-z0-9_])"
    Set matches = rx.Execute(s)
    For Each m In matches
        If InStr(m.Value, "[") > 0 Or Len(m.SubMatches(2)) = 0 Then
            FMP_NameIsRelative = True
            Exit Function
        End If
    Next m
    Exit Function
Conservative:
    FMP_NameIsRelative = True
End Function

Private Function FMP_StripQualifierAndStrings(ByVal f As String) As String
    Dim p As Long, q As Long, qualifier As String
    p = 1
    Do While p <= Len(f)
        If Mid$(f, p, 1) = """" Then
            q = FMP_SkipString(f, p)
            FMP_StripQualifierAndStrings = FMP_StripQualifierAndStrings & Space$(q - p)
            p = q
        ElseIf FMP_TryQualifier(f, p, qualifier, q) Then
            FMP_StripQualifierAndStrings = FMP_StripQualifierAndStrings & Space$(q - p)
            p = q
        Else
            FMP_StripQualifierAndStrings = FMP_StripQualifierAndStrings & Mid$(f, p, 1)
            p = p + 1
        End If
    Loop
End Function

Private Function FMP_IsConstantName(ByVal f As String) As Boolean
    Dim s As String
    s = Trim$(f)
    If Left$(s, 1) = "=" Then s = Mid$(s, 2)
    If IsNumeric(s) Then FMP_IsConstantName = True: Exit Function
    If UCase$(s) = "TRUE" Or UCase$(s) = "FALSE" Then FMP_IsConstantName = True: Exit Function
    If Left$(s, 1) = """" And FMP_SkipString(s, 1) > Len(s) Then FMP_IsConstantName = True
End Function

Private Function FMP_ReadStructured(ByVal f As String, ByVal p As Long, ByRef text As String, ByRef endP As Long) As Boolean
    Dim q As Long, depth As Long, ch As String
    If Mid$(f, p, 1) <> "[" Then Exit Function
    q = p
    Do While q <= Len(f)
        ch = Mid$(f, q, 1)
        If ch = "'" And InStr(1, "[]#'@", Mid$(f, q + 1, 1), vbBinaryCompare) > 0 And q < Len(f) Then
            q = q + 2
        Else
            If ch = "[" Then depth = depth + 1
            If ch = "]" Then depth = depth - 1
            q = q + 1
            If depth = 0 Then
                text = Mid$(f, p, q - p)
                endP = q
                FMP_ReadStructured = True
                Exit Function
            End If
        End If
    Loop
End Function

Private Sub FMP_AddStructured(ByVal result As Collection, ByVal seen As Object, ByVal issues As Collection, ByVal warned As Object, ByVal source As Range, ByVal tableName As String, ByVal spec As String)
    Dim lo As ListObject, candidate As ListObject, ws As Worksheet, leaves As Collection
    Dim rowRange As Range, selected As Range, part As Range, colRange As Range, r As Range
    Dim cols As New Collection, item As Variant, raw As String, text As String, thisRow As Boolean
    Dim hasAll As Boolean, hasHeaders As Boolean, hasData As Boolean, hasTotals As Boolean
    Dim hadItem As Boolean, i As Long, minCol As Long, maxCol As Long, colIndex As Long
    On Error GoTo Failed
    If Len(tableName) > 0 Then
        Set lo = FMP_FindTable(source.Parent.Parent, tableName)
    Else
        For Each candidate In source.Parent.ListObjects
            If Not Application.Intersect(source, candidate.Range) Is Nothing Then
                Set lo = candidate
                Exit For
            End If
        Next candidate
    End If
    If lo Is Nothing Then
        FMP_Issue issues, warned, "Structured reference has no resolvable table: " & tableName & spec
        Exit Sub
    End If
    Set leaves = FMP_StructuredLeaves(spec)
    If leaves.Count = 0 Then GoTo Failed
    thisRow = FMP_StructuredHasThisRow(spec)
    hadItem = thisRow
    For Each item In leaves
        raw = CStr(item)
        text = FMP_UnescapeStructured(raw)
        If Left$(raw, 1) = "#" Then
            hadItem = True
            Select Case UCase$(Trim$(text))
                Case "#ALL": hasAll = True
                Case "#HEADERS": hasHeaders = True
                Case "#DATA": hasData = True
                Case "#TOTALS": hasTotals = True
                Case "#THIS ROW": thisRow = True
                Case Else: GoTo Failed
            End Select
        ElseIf Left$(raw, 1) = "@" Then
            thisRow = True
            hadItem = True
            text = FMP_UnescapeStructured(Mid$(raw, 2))
            If Len(text) > 0 Then cols.Add text
        Else
            cols.Add text
        End If
    Next item
    If thisRow Then
        If hasAll Or hasHeaders Or hasTotals Then GoTo Failed
        If lo.DataBodyRange Is Nothing Then GoTo Failed
        Set rowRange = Application.Intersect(lo.DataBodyRange, source.EntireRow)
        If rowRange Is Nothing Then GoTo Failed
    ElseIf hasAll Then
        Set rowRange = lo.Range
    Else
        If Not hadItem Then hasData = True
        If hasData And Not lo.DataBodyRange Is Nothing Then Set rowRange = lo.DataBodyRange
        If hasHeaders Then
            Set part = lo.HeaderRowRange
            FMP_UnionInto rowRange, part
        End If
        If hasTotals Then
            Set part = Nothing
            On Error Resume Next
            Set part = lo.TotalsRowRange
            On Error GoTo Failed
            If part Is Nothing Then
                FMP_Issue issues, warned, "Table totals reference has no visible totals row: " & lo.Name & spec
            Else
                FMP_UnionInto rowRange, part
            End If
        End If
    End If
    If rowRange Is Nothing Then
        FMP_Issue issues, warned, "Structured reference has no cells in the current table state: " & lo.Name & spec
        Exit Sub
    End If
    If cols.Count = 0 Then
        Set selected = rowRange
    ElseIf cols.Count = 2 And InStr(spec, "]:[") > 0 Then
        minCol = FMP_ListColumnIndex(lo, CStr(cols(1)))
        maxCol = FMP_ListColumnIndex(lo, CStr(cols(2)))
        If minCol = 0 Or maxCol = 0 Then GoTo Failed
        If minCol > maxCol Then i = minCol: minCol = maxCol: maxCol = i
        Set colRange = lo.Range.Columns(minCol).Resize(lo.Range.Rows.Count, maxCol - minCol + 1)
        Set selected = Application.Intersect(rowRange, colRange)
    Else
        For Each item In cols
            colIndex = FMP_ListColumnIndex(lo, CStr(item))
            If colIndex = 0 Then GoTo Failed
            Set part = Application.Intersect(rowRange, lo.ListColumns(colIndex).Range)
            FMP_UnionInto selected, part
        Next item
    End If
    If selected Is Nothing Then GoTo Failed
    FMP_AddRange result, seen, selected, "table", "Table reference: " & lo.Name & spec
    Exit Sub
Failed:
    FMP_Issue issues, warned, "Structured reference was not resolved safely: " & tableName & spec
End Sub

Private Function FMP_FindTable(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Or StrComp(lo.DisplayName, tableName, vbTextCompare) = 0 Then
                Set FMP_FindTable = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

Private Function FMP_StructuredHasThisRow(ByVal spec As String) As Boolean
    Dim p As Long
    p = 1
    Do While p <= Len(spec)
        If Mid$(spec, p, 1) = "'" And p < Len(spec) And InStr(1, "[]#'@", Mid$(spec, p + 1, 1), vbBinaryCompare) > 0 Then
            p = p + 2
        Else
            If Mid$(spec, p, 2) = "[@" Then
                FMP_StructuredHasThisRow = True
                Exit Function
            End If
            p = p + 1
        End If
    Loop
End Function

Private Function FMP_StructuredLeaves(ByVal spec As String) As Collection
    Dim result As New Collection, p As Long, ch As String, depth As Long, s As String, hadChild As Boolean
    Dim stackChild As Object
    Set stackChild = CreateObject("Scripting.Dictionary")
    p = 1
    Do While p <= Len(spec)
        ch = Mid$(spec, p, 1)
        If ch = "'" And p < Len(spec) And InStr(1, "[]#'@", Mid$(spec, p + 1, 1), vbBinaryCompare) > 0 Then
            s = s & ch & Mid$(spec, p + 1, 1)
            p = p + 2
        ElseIf ch = "[" Then
            If depth > 0 Then stackChild(CStr(depth)) = True
            depth = depth + 1
            stackChild(CStr(depth)) = False
            s = ""
            p = p + 1
        ElseIf ch = "]" Then
            If depth > 0 Then
                If Not CBool(stackChild(CStr(depth))) Then result.Add s
                stackChild.Remove CStr(depth)
                depth = depth - 1
                s = ""
            End If
            p = p + 1
        Else
            If depth > 0 Then s = s & ch
            p = p + 1
        End If
    Loop
    Set FMP_StructuredLeaves = result
End Function

Private Function FMP_UnescapeStructured(ByVal s As String) As String
    Dim p As Long, ch As String
    p = 1
    Do While p <= Len(s)
        ch = Mid$(s, p, 1)
        If ch = "'" And p < Len(s) And InStr(1, "[]#'@", Mid$(s, p + 1, 1), vbBinaryCompare) > 0 Then
            p = p + 1
            ch = Mid$(s, p, 1)
        End If
        FMP_UnescapeStructured = FMP_UnescapeStructured & ch
        p = p + 1
    Loop
End Function

Private Function FMP_ListColumnIndex(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim c As ListColumn
    For Each c In lo.ListColumns
        If StrComp(c.Name, columnName, vbTextCompare) = 0 Then
            FMP_ListColumnIndex = c.Index
            Exit Function
        End If
    Next c
End Function

Private Sub FMP_UnionInto(ByRef target As Range, ByVal part As Range)
    If part Is Nothing Then Exit Sub
    If target Is Nothing Then
        Set target = part
    Else
        Set target = Application.Union(target, part)
    End If
End Sub

Private Sub FMP_AddRange(ByVal result As Collection, ByVal seen As Object, ByVal r As Range, ByVal kind As String, ByVal note As String)
    Dim a As Range
    For Each a In r.Areas
        FMP_AddSource result, seen, a.Parent.Parent.Name, a.Parent.Name, a.Address(False, False, xlA1), kind, note
    Next a
End Sub

Private Sub FMP_AddSource(ByVal result As Collection, ByVal seen As Object, ByVal book As String, ByVal sheet As String, ByVal address As String, ByVal kind As String, ByVal note As String)
    Dim key As String, d As Object
    address = FMP_CanonicalAddress(address)
    key = book & Chr$(30) & sheet & Chr$(30) & address
    If seen.Exists(key) Then
        Set d = seen(key)
        If Len(note) > 0 And InStr(1, d("note"), note, vbTextCompare) = 0 Then
            If Len(d("note")) > 0 Then d("note") = d("note") & "; "
            d("note") = d("note") & note
        End If
        Exit Sub
    End If
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    d.Add "book", book
    d.Add "sheet", sheet
    d.Add "address", address
    d.Add "kind", kind
    d.Add "note", note
    seen.Add key, d
    result.Add d
End Sub

Private Function FMP_CanonicalAddress(ByVal address As String) As String
    Dim pieces As Variant, firstPart As String, secondPart As String, firstKind As String, secondKind As String
    Dim q As Long, c1 As Long, c2 As Long, r1 As Long, r2 As Long, t As Long, p As Long
    address = UCase$(Replace(address, "$", ""))
    FMP_CanonicalAddress = address
    pieces = Split(address, ":")
    If UBound(pieces) <> 1 Then Exit Function
    If Not FMP_ReadA1Part(CStr(pieces(0)), 1, firstPart, firstKind, q) Then Exit Function
    If Not FMP_ReadA1Part(CStr(pieces(1)), 1, secondPart, secondKind, q) Then Exit Function
    If firstKind <> secondKind Then Exit Function
    Select Case firstKind
        Case "column"
            c1 = FMP_ColumnNumber(firstPart): c2 = FMP_ColumnNumber(secondPart)
            If c1 > c2 Then t = c1: c1 = c2: c2 = t
            FMP_CanonicalAddress = FMP_ColumnLabel(c1) & ":" & FMP_ColumnLabel(c2)
        Case "row"
            r1 = CLng(firstPart): r2 = CLng(secondPart)
            If r1 > r2 Then t = r1: r1 = r2: r2 = t
            FMP_CanonicalAddress = CStr(r1) & ":" & CStr(r2)
        Case "cell"
            p = 1
            Do While Mid$(firstPart, p, 1) Like "[A-Z]": p = p + 1: Loop
            c1 = FMP_ColumnNumber(Left$(firstPart, p - 1)): r1 = CLng(Mid$(firstPart, p))
            p = 1
            Do While Mid$(secondPart, p, 1) Like "[A-Z]": p = p + 1: Loop
            c2 = FMP_ColumnNumber(Left$(secondPart, p - 1)): r2 = CLng(Mid$(secondPart, p))
            If c1 > c2 Then t = c1: c1 = c2: c2 = t
            If r1 > r2 Then t = r1: r1 = r2: r2 = t
            FMP_CanonicalAddress = FMP_ColumnLabel(c1) & CStr(r1)
            If c1 <> c2 Or r1 <> r2 Then FMP_CanonicalAddress = FMP_CanonicalAddress & ":" & FMP_ColumnLabel(c2) & CStr(r2)
    End Select
End Function

Private Function FMP_ColumnLabel(ByVal columnNumber As Long) As String
    Do While columnNumber > 0
        columnNumber = columnNumber - 1
        FMP_ColumnLabel = Chr$(65 + columnNumber Mod 26) & FMP_ColumnLabel
        columnNumber = columnNumber \ 26
    Loop
End Function

Private Sub FMP_Issue(ByVal issues As Collection, ByVal warned As Object, ByVal message As String)
    If issues Is Nothing Then Exit Sub
    If warned Is Nothing Then
        issues.Add message
    ElseIf Not warned.Exists(message) Then
        warned.Add message, True
        issues.Add message
    End If
End Sub
