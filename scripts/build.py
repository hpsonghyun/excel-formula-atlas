"""Build the single-module VBA release using only files in this repository.

Run: python scripts/build.py
Python 3.10+ and its standard library are sufficient. Normal Excel users need
only dist/ExcelFormulaAtlas.txt; they do not need Python or this build script.
"""
from pathlib import Path
import base64, re, json, hashlib
from progress_instrumentation import instrument_main, instrument_snapshot, progress_helper

root=Path(__file__).resolve().parents[1]
base=root/'src/vba'
template=root/'src/viewer/template.html'
bridge=root/'src/navigation/navigation_bridge.ps1'

def embedded_function(name, payload):
    encoded=base64.b64encode(payload).decode('ascii')
    groups=[encoded[i:i+10000] for i in range(0,len(encoded),10000)]
    out=[f'Private Function {name}() As String','    Dim s As String']
    out += [f'    s = s & {name}Part{i+1}()' for i in range(len(groups))]
    out += [f'    {name} = FM_Base64ToText(s)','End Function']
    for i,group in enumerate(groups,1):
        out += ['',f'Private Function {name}Part{i}() As String','    Dim s As String']
        out += ['    s = s & "'+group[j:j+180]+'"' for j in range(0,len(group),180)]
        out += [f'    {name}Part{i} = s','End Function']
    return '\n'.join(out)+'\n'

main=(base/'FormulaMapMain.bas').read_text(encoding='utf-8-sig')
entry=(base/'SingleFileEntry.bas').read_text(encoding='utf-8-sig')
main,n=re.subn(r'Public Sub BuildFormulaMap\(\).*?\nEnd Sub',entry.strip(),main,count=1,flags=re.S)
assert n==1
old='Optional ByVal silent As Boolean = True) As String'
assert main.count(old)==1
main=main.replace(old,'Optional ByVal silent As Boolean = True, Optional ByVal createExcelReport As Boolean = True) As String')
old='''    Set report = FM_CreateReport()
    FM_AddNativeSpatialMap report
    report.SaveAs Filename:=base & ".xlsx", FileFormat:=xlOpenXMLWorkbook'''
assert old in main
main=main.replace(old,'''    If createExcelReport Then
        Set report = FM_CreateReport()
        FM_AddNativeSpatialMap report
        report.SaveAs Filename:=base & ".xlsx", FileFormat:=xlOpenXMLWorkbook
    End If''')
main=main.replace('    report.Worksheets(1).Activate','    If Not report Is Nothing Then report.Worksheets(1).Activate')

# SpecialCells marked a saved workbook dirty in Excel QA. Use a read-only
# Find/Collection scan; do not conceal changes by assigning Workbook.Saved.
main=main.replace('ws As Worksheet, formulas As Range, cell As Range',
                  'ws As Worksheet, formulas As Collection, cell As Range')
formula_scan='''        Set formulas = Nothing
        On Error Resume Next
        Set formulas = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo Failed'''
assert main.count(formula_scan)==2
main=main.replace(formula_scan,'        Set formulas = FM_CollectFormulaCells(ws)')
assert main.count('For Each cell In formulas.Cells')==2
main=main.replace('For Each cell In formulas.Cells','For Each cell In formulas')
formula_helper='''Private Function FM_CollectFormulaCells(ByVal ws As Worksheet) As Collection
    Dim result As Collection, scope As Range, hit As Range, nextHit As Range
    Dim firstAddress As String
    Set result = New Collection
    Set scope = ws.Cells
    Set hit = scope.Find(What:="=*", After:=ws.Cells(ws.Rows.Count, ws.Columns.Count), _
        LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    If Not hit Is Nothing Then
        firstAddress = hit.Address
        Do
            If hit.HasFormula Then result.Add hit
            Set nextHit = scope.Find(What:="=*", After:=hit, _
                LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, _
                SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
            If nextHit Is Nothing Then Exit Do
            Set hit = nextHit
            If hit.Address = firstAddress Then Exit Do
        Loop
    End If
    Set FM_CollectFormulaCells = result
End Function

'''
anchor='Private Function FM_IsFormulaAnchor(ByVal cell As Range) As Boolean'
assert main.count(anchor)==1
main=main.replace(anchor,formula_helper+anchor)
assert 'SpecialCells' not in main
main=instrument_main(main)

navigation=(base/'SpatialNavigation.bas').read_text(encoding='utf-8-sig')
navigation=re.sub(r'    If Len\(mFMSource.Path\) = 0 Then Err.Raise 5, , .*?\n','',navigation,count=1)
old='''    scriptPath = ThisWorkbook.Path & Application.PathSeparator & "navigation_bridge.ps1"
    If Not fso.FileExists(scriptPath) Then Err.Raise 53, , "Place navigation_bridge.ps1 beside this tool."
'''
assert old in navigation
navigation=navigation.replace(old,'')
old='    fso.CreateFolder folder\n'
assert navigation.count(old)==1
navigation=navigation.replace(old,old+'''    scriptPath = folder & Application.PathSeparator & "navigation_bridge.ps1"
    FM_WriteUTF8 scriptPath, FM_BridgeScript()
''')
start_begin=navigation.index('Public Function StartFormulaMapNavigation(')
start_end=navigation.index('\nEnd Function',start_begin)+len('\nEnd Function')
start=navigation[start_begin:start_end]
prepare=start[:start.index('    FM_PollNavigation\n')]
prepare=prepare.replace('Public Function StartFormulaMapNavigation(ByVal htmlPath As String, Optional ByVal openBrowser As Boolean = True) As String',
                        'Public Function FM_PrepareNavigationSession(ByVal htmlPath As String) As String')
prepare+='''    FM_PrepareNavigationSession = folder
    Exit Function
Failed:
    command = Err.Description
    FM_StopNavigation
    Err.Raise 5, "FormulaMap navigation setup", command
End Function'''
launch='''Public Function StartFormulaMapNavigation(ByVal htmlPath As String, Optional ByVal openBrowser As Boolean = True) As String
    Dim state As Object, fso As Object, shell As Object
    Dim folder As String, scriptPath As String, command As String, url As String, deadline As Date
    On Error GoTo Failed
    folder = FM_PrepareNavigationSession(htmlPath)
    Set state = FM_NavState()
    Set fso = CreateObject("Scripting.FileSystemObject")
    scriptPath = folder & Application.PathSeparator & "navigation_bridge.ps1"
'''+start[start.index('    FM_PollNavigation\n'):]
navigation=navigation[:start_begin]+prepare+'\n\n'+launch+navigation[start_end:]
old='ScheduleNext:\n    If state("active") Then'
assert old in navigation
navigation=navigation.replace(old,'''    If reason = "source_closed" Or reason = "source_path_changed" Then
        FM_StopNavigation
        Exit Sub
    End If
ScheduleNext:
    If state("active") Then''')

header="""' Excel Formula Atlas - single-module VBA distribution
' Turn an unfamiliar workbook into an interactive map of formula references.
' This text file contains the VBA engine, HTML viewer, and local navigation helper.
'
' QUICK START (Windows desktop Excel)
' 1. Open this file in a text editor. Select all and copy the entire contents.
' 2. In Excel, press Alt+F11 and choose Insert > Module to create a NEW module.
' 3. In that NEW module's code window only, press Ctrl+A and paste the entire TXT.
'    This replaces any automatic declaration line. Keep existing user modules.
'    Use a standard module, not ThisWorkbook or a worksheet code window.
' 4. Return to Excel. Press Alt+F8, choose ExcelMap_Start, and click Run.
' 5. Choose 1 for the current workbook or 2 for a file, then an output folder.
' 6. Explore the map. Select a referenced cell and use Open in Excel to jump back.
'
' EXPLORING THE MAP
' Left drag selects an area and highlights the cells/ranges that feed it.
' Right drag pans the map. The mouse wheel zooms around the pointer.
' Click a cell to inspect its actual nearby table, value, and formula.
' Click a connection to highlight its endpoints and related formula flow.
' Search worksheet names, cell addresses, nearby labels, and formulas.
' The viewer starts in English; its language button also offers Korean.
' Arrows point from source values to the formulas that use them.
' Dashed membership links identify formula cells inside referenced ranges.
' Excel shows analysis stages in its status bar. The HTML shows preparation
' stages, actual completed-item progress, and elapsed time while rendering.
' Progress counts completed work; it is not an estimate of time remaining.
'
' LOCAL FILES AND EXCEL NAVIGATION
' No add-in, installer, Python, or extra source file is required for normal use.
' The macro writes one HTML file to your chosen folder. It extracts its embedded
' PowerShell navigation helper and connection files into a temporary folder.
' Workbook contents are processed locally. The helper listens only on this PC
' at 127.0.0.1 and uses a temporary session token. It does not upload your workbook.
' The exported HTML contains workbook data. Share it only with intended readers.
' Excel navigation requires Windows PowerShell and an open source workbook.
' Workplace macro or process policies may prevent the helper from starting.
' This code does not change those policies. The standalone HTML remains usable.
' A newly selected file opens read-only, with external-link updates disabled and
' VBA macros disabled. Already-open workbooks are reused. Excel may separately
' prompt about legacy Excel 4.0 macros. This tool does not edit source cell values
' or formulas, convert the source file, or automatically save the source workbook.
'
' STOPPING AND SAVING
' Run ExcelMap_Stop to stop the navigation connection. The HTML file is retained.
' Stop before closing the workbook that holds this code or using Save As on it.
' The included Auto_Close is a convenience for manual closing. VBA Workbook.Close
' does not run Auto_Close. If an Auto_Close already exists, merge the call to
' FM_StopNavigation into it instead of adding a second Auto_Close procedure.
' Save the code-hosting workbook as .xlsm if you want to keep the macro for later.
' The navigation session expires after two hours. Hidden sheets are not unhidden.
' Regenerate the map after source values, formulas, workbook name, or path change.
'
' SCOPE AND LIMITS
' This is static reference analysis, not a calculation engine or workbook audit.
' Conditional formulas include possible branches; dynamic and external references
' may require review. Read the notices in the map before relying on coverage.
' Formula scan cap: 50,000. Captured geometry: up to 2,000 rows and 256 columns.
' Context capture starts with 100 rows / 40 columns and referenced-cell neighbors,
' capped at 20,000 positions per worksheet. Uncaptured cells are identified.
' Visible connections: up to 200 in selected flow, or 700 in the overview.
' The HTML is a snapshot. Charts, pictures, and every Excel style are not replicated.
' See the repository documentation for supported references and verification scope.
' ============================================================================
"""
license_text=(root/'LICENSE').read_text(encoding='utf-8').strip()
header+='\n'+'\n'.join("' "+line if line else "'" for line in license_text.splitlines())+'\n'
sections=[main,(base/'FormulaMapParser.bas').read_text(encoding='utf-8-sig'),
          instrument_snapshot((base/'SpatialSnapshot.bas').read_text(encoding='utf-8-sig')),navigation,
          (base/'SpatialNativeMap.bas').read_text(encoding='utf-8-sig'),progress_helper()]
body='\n\n'.join(sections)
body=re.sub(r'^Attribute VB_Name.*\n','',body,flags=re.M)
body=re.sub(r'^Option Explicit\s*\n','',body,flags=re.M)
body=header+'\nOption Explicit\n\n'+body
body+='\n'+embedded_function('FM_Template',template.read_bytes())
body+='\n'+embedded_function('FM_BridgeScript',bridge.read_bytes())
body=body.replace('\r\n','\n').replace('\n','\r\n')
names=re.findall(r'^(?:Public |Private )?(?:Sub|Function) (\w+)',body,flags=re.M|re.I)
assert len(names)==len(set(x.casefold() for x in names)), 'Duplicate procedures'
assert body.count('Option Explicit')==1
assert 'Attribute VB_Name' not in body
assert 'ThisWorkbook.Path' not in navigation
assert 'VBProject' not in body
assert max(map(len,body.splitlines()))<1000
target=root/'dist/ExcelFormulaAtlas.txt'
target.parent.mkdir(parents=True,exist_ok=True)
target.write_text(body,encoding='utf-8-sig',newline='')
download=root/'docs/downloads/ExcelFormulaAtlas.txt'
download.parent.mkdir(parents=True,exist_ok=True)
download.write_bytes(target.read_bytes())
version=(root/'VERSION').read_text(encoding='utf-8').strip()
metadata={'version':version,'bytes':target.stat().st_size,'lines':len(body.splitlines()),'procedures':len(names),
 'maxLineLength':max(map(len,body.splitlines())),
 'sha256':hashlib.sha256(target.read_bytes()).hexdigest(),
 'embeddedHtmlSha256':hashlib.sha256(template.read_bytes()).hexdigest(),
 'embeddedBridgeSha256':hashlib.sha256(bridge.read_bytes()).hexdigest()}
inputs=[*sorted((root/'src').rglob('*')),Path(__file__),root/'scripts/progress_instrumentation.py',root/'VERSION',root/'LICENSE']
metadata['sourceHashes']={str(path.relative_to(root)).replace('\\','/'):hashlib.sha256(path.read_bytes()).hexdigest() for path in inputs if path.is_file()}
(root/'dist/build-metadata.json').write_text(json.dumps(metadata,indent=2)+'\n',encoding='utf-8',newline='\n')
print(json.dumps(metadata))
