"""Build-time, read-only progress instrumentation for the one-module VBA export.

Call instrument_main only after the standard single-file main/Find transforms.
Call instrument_snapshot on the original SpatialSnapshot.bas source.
Append progress_helper() once as a separate VBA module-text section.
The original v2 VBA files are deliberately not modified.
"""
import re


def _once(text: str, old: str, new: str) -> str:
    if text.count(old) != 1:
        raise AssertionError(f"Expected one progress anchor: {old[:100]!r}")
    return text.replace(old, new, 1)


def _procedure(text: str, name: str, change) -> str:
    pattern = rf"^(?:Public |Private )(?:Sub|Function) {re.escape(name)}\b.*?^End (?:Sub|Function)$"
    matches = list(re.finditer(pattern, text, re.M | re.S))
    if len(matches) != 1:
        raise AssertionError(f"Expected one procedure: {name}")
    found = matches[0]
    return text[:found.start()] + change(found.group()) + text[found.end():]


def instrument_main(main: str) -> str:
    main = main.replace("\r\n", "\n")
    main = _once(main, "Private mFMWorksheetJSON As String", """Private mFMWorksheetJSON As String
Private mFMProgressBusy As Boolean
Private mFMProgressLastTick As Single
Private mFMProgressStage As Long, mFMProgressSheetDone As Long""")

    def start(body):
        # A queued macro must stop before it opens dialogs or stops the active
        # navigation session during a DoEvents yielded by the analysis.
        return _once(body, "    On Error GoTo Failed", """    If mFMProgressBusy Then
        MsgBox "A formula map is already being created. Please wait for it to finish.", vbInformation
        Exit Sub
    End If
    On Error GoTo Failed""")

    main = _procedure(main, "ExcelMap_Start", start)

    def build(body):
        body = _once(body, "    Dim errorNumber As Long, errorText As String", """    Dim errorNumber As Long, errorText As String
    Dim fmProgressSheet As Long, oldDisplayStatusBar As Boolean
    Dim fmProgressData As String, fmProgressHTML As String
    If mFMProgressBusy Then Err.Raise 5, "FormulaMap", _
        "A formula map is already being created."
"""
        )
        body = _once(body, "    oldCancel = Application.EnableCancelKey", """    oldCancel = Application.EnableCancelKey
    oldDisplayStatusBar = Application.DisplayStatusBar""")
        body = _once(body, "    On Error GoTo Failed", """    On Error GoTo Failed
    mFMProgressBusy = True: mFMProgressLastTick = -1
    mFMProgressStage = 1: mFMProgressSheetDone = 0
    Application.DisplayStatusBar = True""")
        body = _once(body, "    Application.DisplayAlerts = False: Application.EnableCancelKey = xlErrorHandler", """    Application.DisplayAlerts = False: Application.EnableCancelKey = xlErrorHandler
    FM_ShowProgress "1/5 Finding formulas", 0, 0, "Preparing", True""")
        body = _once(body, "    For Each ws In mFMSource.Worksheets\n        mFMSheets.Add ws.Name", """    fmProgressSheet = 0
    For Each ws In mFMSource.Worksheets
        mFMProgressSheetDone = fmProgressSheet
        FM_ShowProgress "1/5 Finding formulas", fmProgressSheet, mFMSource.Worksheets.Count, ws.Name, True
        mFMSheets.Add ws.Name""")
        body = _once(body, "    Next ws\n    For Each ws In mFMSource.Worksheets", """        fmProgressSheet = fmProgressSheet + 1
        FM_ShowProgress "1/5 Finding formulas", fmProgressSheet, mFMSource.Worksheets.Count, _
            "Formulas found: " & Format$(mFMTotalFormulas, "#,##0"), True
    Next ws
    mFMProgressStage = 2
    FM_ShowProgress "2/5 Analyzing references", 0, _
        Application.Min(mFMTotalFormulas, FM_MAX_FORMULAS), "", True
    For Each ws In mFMSource.Worksheets""")
        body = _once(body, """                If mFMFormulaCount Mod 100 = 0 Then
                    Application.StatusBar = "Analyzing formulas: " & Format$(mFMFormulaCount, "#,##0") & " / " & Format$(mFMTotalFormulas, "#,##0")
                    DoEvents
                End If""", """                If mFMFormulaCount Mod 25 = 0 Then
                    FM_ShowProgress "2/5 Analyzing references", mFMFormulaCount, _
                        Application.Min(mFMTotalFormulas, FM_MAX_FORMULAS), ws.Name
                End If""")
        body = _once(body, '    Application.StatusBar = "Linking formulas inside ranges"', """    FM_ShowProgress "2/5 Analyzing references", mFMFormulaCount, _
        Application.Min(mFMTotalFormulas, FM_MAX_FORMULAS), "", True
    FM_ShowProgress "3/5 Linking ranges", 0, 0, "Preparing", True""")
        body = _once(body, '    Application.StatusBar = "Capturing worksheet positions"', """    mFMProgressStage = 4
    FM_ShowProgress "4/5 Reading sheets", 0, mFMSource.Worksheets.Count, "", True""")
        body = _once(body, '    FM_WriteUTF8 htmlPath, Replace$(FM_Template(), "__FORMULA_MAP_DATA__", FM_DataJSON())', """    mFMProgressStage = 5
    FM_ShowProgress "5/5 Creating HTML", 0, 0, "Preparing map data", True
    fmProgressData = FM_DataJSON()
    FM_ShowProgress "5/5 Creating HTML", 0, 0, "Building HTML", True
    fmProgressHTML = Replace$(FM_Template(), "__FORMULA_MAP_DATA__", fmProgressData)
    FM_ShowProgress "5/5 Creating HTML", 0, 0, "Saving the file", True
    FM_WriteUTF8 htmlPath, fmProgressHTML
    FM_ShowProgress "5/5 Creating HTML", 1, 1, "Saved", True""")
        body = _once(body, "    Application.StatusBar = oldStatus", """    Application.StatusBar = oldStatus
    Application.DisplayStatusBar = oldDisplayStatusBar
    mFMProgressBusy = False: mFMProgressStage = 0""")
        # Cleanup errors must never strand the module's reentrancy guard. All
        # settings are attempted, while the original analysis error is retained.
        body = _once(body, "Clean:\n    Application.ScreenUpdating", "Clean:\n    On Error Resume Next\n    Application.ScreenUpdating")
        body = _once(body, "    mFMProgressBusy = False: mFMProgressStage = 0", """    mFMProgressBusy = False: mFMProgressStage = 0
    On Error GoTo 0""")
        return body

    main = _procedure(main, "FM_BuildFile", build)

    def scanner(body):
        body = _once(body, "    Dim firstAddress As String", "    Dim firstAddress As String, fmProgressHits As Long")
        return _once(body, "            If hit.HasFormula Then result.Add hit", """            If hit.HasFormula Then result.Add hit
            fmProgressHits = fmProgressHits + 1
            If fmProgressHits Mod 200 = 0 Then
                If mFMProgressStage = 1 Then
                    FM_ShowProgress "1/5 Finding formulas", mFMProgressSheetDone, _
                        mFMSource.Worksheets.Count, ws.Name & " · Found: " & Format$(result.Count, "#,##0")
                ElseIf mFMProgressStage = 2 Then
                    FM_ShowProgress "2/5 Analyzing references", mFMFormulaCount, _
                        Application.Min(mFMTotalFormulas, FM_MAX_FORMULAS), ws.Name & " · Reading formula list"
                End If
            End If""")

    main = _procedure(main, "FM_CollectFormulaCells", scanner)

    def ranges(body):
        body = _once(body, "    Dim r As Range, c As Range, checks As Long, wb As String", """    Dim r As Range, c As Range, checks As Long, wb As String
    Dim fmProgressRangesDone As Long, fmProgressNodesDone As Long""")
        body = _once(body, "    For Each d In mFMNodes", """    For Each d In mFMNodes
        fmProgressNodesDone = fmProgressNodesDone + 1
        If fmProgressNodesDone Mod 250 = 0 Then FM_ShowProgress _
            "3/5 Linking ranges", 0, 0, "Classifying: " & Format$(fmProgressNodesDone, "#,##0")""")
        body = _once(body, "    For Each d In ranges", """    FM_ShowProgress "3/5 Linking ranges", 0, ranges.Count, "Ranges", True
    For Each d In ranges""")
        body = _once(body, "                checks = checks + 1", """                checks = checks + 1
                If checks Mod 2500 = 0 Then FM_ShowProgress "3/5 Linking ranges", _
                    fmProgressRangesDone, ranges.Count, "Ranges completed · Checks: " & Format$(checks, "#,##0")""")
        body = _once(body, "                    Exit Sub", """                    FM_ShowProgress "3/5 Linking ranges", fmProgressRangesDone, ranges.Count, _
                        "Check limit reached: " & Format$(FM_MAX_MEMBERSHIP_CHECKS, "#,##0"), True
                    Exit Sub""")
        body = _once(body, "    Next d\nEnd Sub", """        fmProgressRangesDone = fmProgressRangesDone + 1
        FM_ShowProgress "3/5 Linking ranges", fmProgressRangesDone, ranges.Count, _
            "Ranges completed · Checks: " & Format$(checks, "#,##0")
    Next d
    FM_ShowProgress "3/5 Linking ranges", fmProgressRangesDone, ranges.Count, _
        "Completed · Checks: " & Format$(checks, "#,##0"), True
End Sub""")
        return body

    main = _procedure(main, "FM_LinkRangeMembers", ranges)

    def data(body):
        body = _once(body, '            i = i + 1\n        Next d', """            i = i + 1
            If i Mod 100 = 0 Then FM_ShowProgress "5/5 Preparing cells", _
                i, mFMNodes.Count, "", False
        Next d""")
        for collection, label, expression in [
            ("mFMEdges", "Preparing references", 'FM_ObjectJSON(d, Array("s", "t", "kind", "note"))'),
            ("mFMIssues", "Preparing notices", 'FM_ObjectJSON(d, Array("sheet", "address", "formula", "reason"))'),
        ]:
            body = _once(body, f"{expression}: i = i + 1", f"""{expression}: i = i + 1
            If i Mod 100 = 0 Then FM_ShowProgress "5/5 HTML {label}", i, {collection}.Count, "", False""")
        return body

    main = _procedure(main, "FM_DataJSON", data)
    return main


def instrument_snapshot(snapshot: str) -> str:
    snapshot = snapshot.replace("\r\n", "\n")

    def all_sheets(body):
        body = _once(body, "    For Each ws In mFMSource.Worksheets", """    For Each ws In mFMSource.Worksheets
        mFMProgressSheetDone = i
        FM_ShowProgress "4/5 Reading sheets", i, mFMSource.Worksheets.Count, ws.Name, True""")
        return _once(body, "        DoEvents", """        FM_ShowProgress "4/5 Reading sheets", i, mFMSource.Worksheets.Count, ws.Name, True""")

    snapshot = _procedure(snapshot, "FM_SnapshotJSON", all_sheets)

    def one_sheet(body):
        body = _once(body, "    Dim fullWidth As Double, fullHeight As Double, width As Double, height As Double", """    Dim fullWidth As Double, fullHeight As Double, width As Double, height As Double
    Dim fmProgressNodesDone As Long, fmProgressCellsDone As Long""")
        body = _once(body, "    For Each d In mFMNodes", """    For Each d In mFMNodes
        fmProgressNodesDone = fmProgressNodesDone + 1
        If fmProgressNodesDone Mod 250 = 0 Then FM_ShowProgress "4/5 Reading sheets", _
            mFMProgressSheetDone, mFMSource.Worksheets.Count, ws.Name & " · Checking nodes: " & Format$(fmProgressNodesDone, "#,##0")""")
        body = _once(body, "    Next r\n    For c = 1 To gCols", """        If r Mod 100 = 0 Then FM_ShowProgress "4/5 Reading sheets", _
            mFMProgressSheetDone, mFMSource.Worksheets.Count, ws.Name & " · Row geometry: " & r & "/" & gRows
    Next r
    For c = 1 To gCols""")
        body = _once(body, "    Next c\n    Set captures", """        If c Mod 50 = 0 Then FM_ShowProgress "4/5 Reading sheets", _
            mFMProgressSheetDone, mFMSource.Worksheets.Count, ws.Name & " · Column geometry: " & c & "/" & gCols
    Next c
    Set captures""")
        return _once(body, "    Next key\n    If texts.Count", """        fmProgressCellsDone = fmProgressCellsDone + 1
        If fmProgressCellsDone Mod 100 = 0 Then FM_ShowProgress "4/5 Reading sheets", _
            mFMProgressSheetDone, mFMSource.Worksheets.Count, ws.Name & " · Cells: " & _
            Format$(fmProgressCellsDone, "#,##0") & "/" & Format$(captures.Count, "#,##0")
    Next key
    If texts.Count""")

    return _procedure(snapshot, "FM_OneSnapshotJSON", one_sheet)


def progress_helper() -> str:
    return '''' Each percentage is the completed/total count of the displayed stage or
' named subtask, never a guessed elapsed-time percentage or ETA.
Private Sub FM_ShowProgress(ByVal stage As String, ByVal completed As Double, ByVal total As Double, _
    Optional ByVal detail As String = "", Optional ByVal force As Boolean = False)
    Dim tick As Single, elapsed As Single, percent As Long, message As String
    tick = Timer
    elapsed = tick - mFMProgressLastTick
    If elapsed < 0 Then elapsed = elapsed + 86400!
    If Not force And mFMProgressLastTick >= 0 And elapsed < 0.15! Then Exit Sub
    mFMProgressLastTick = tick
    message = stage
    If total > 0 Then
        If completed < 0 Then completed = 0
        If completed > total Then completed = total
        percent = Fix(completed * 100# / total)
        message = message & ": " & Format$(completed, "#,##0") & "/" & Format$(total, "#,##0") & _
            " (" & CStr(percent) & "%)"
    End If
    If Len(detail) > 0 Then message = message & " · " & detail
    Application.StatusBar = message
    DoEvents
End Sub
'''
