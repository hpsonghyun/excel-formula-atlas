<#
Run in Windows PowerShell 5.1 with -STA on a Windows desktop with Excel.
This test creates one disposable Excel instance and modifies only its own QA
workbook. It never changes Trust Center / registry policy or starts the HTTP
navigation helper. Existing Excel instances are never attached to or closed.

Example (paths are relative to the repository root):
powershell.exe -NoProfile -STA -File tests/windows_excel_smoke.ps1 `
  -SingleFilePath dist/ExcelFormulaAtlas.txt `
  -SourcePath examples/OperationsModel.xlsx `
  -TemplatePath src/viewer/template.html `
  -BridgePath src/navigation/navigation_bridge.ps1
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$SingleFilePath,
    [Parameter(Mandatory=$true)][string]$SourcePath,
    [Parameter(Mandatory=$true)][string]$TemplatePath,
    [Parameter(Mandatory=$true)][string]$BridgePath,
    [string]$OutputDirectory = (Join-Path ([IO.Path]::GetTempPath()) ('ExcelFormulaAtlasQA-' + [Guid]::NewGuid().ToString('N')))
)
$ErrorActionPreference = 'Stop'
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    throw 'Use Windows PowerShell with -STA. No Excel instance was created.'
}
foreach ($inputName in @('SingleFilePath','SourcePath','TemplatePath','BridgePath')) {
    Set-Variable -Name $inputName -Value ((Resolve-Path -LiteralPath (Get-Variable $inputName -ValueOnly)).Path)
}
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path
$plainUtf8 = New-Object Text.UTF8Encoding($false)
$checks = New-Object 'Collections.Generic.List[object]'
$sessions = New-Object 'Collections.Generic.List[string]'
function Add-Check([string]$name, [bool]$pass, $detail=$null) {
    $checks.Add(@{ name=$name; pass=$pass; detail=$detail })
}
function Text-Sha([string]$text) {
    $normalized = $text.TrimStart([char]0xFEFF).Replace("`r`n","`n").Replace("`r","`n")
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash($plainUtf8.GetBytes($normalized)))).Replace('-','').ToLowerInvariant() }
    finally { $sha.Dispose() }
}
function Book-Snapshot($book) {
    $parts = New-Object 'Collections.Generic.List[string]'
    foreach ($sheet in $book.Worksheets) {
        $parts.Add('S|' + $sheet.Name + '|' + $sheet.Visible + '|' + $sheet.UsedRange.Address())
        foreach ($value in $sheet.UsedRange.Value2) { $parts.Add('V|' + [string]$value) }
        foreach ($formula in $sheet.UsedRange.Formula2) { $parts.Add('F|' + [string]$formula) }
    }
    foreach ($name in $book.Names) { $parts.Add('N|' + $name.Name + '|' + $name.RefersToR1C1) }
    return ($parts -join "`n")
}
function App-Snapshot($app) {
    return @{ screen=$app.ScreenUpdating; events=$app.EnableEvents; alerts=$app.DisplayAlerts;
        calculation=$app.Calculation; security=$app.AutomationSecurity;
        statusVisible=$app.DisplayStatusBar; status=$app.StatusBar }
}
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class AtlasSmokeWindow {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
}
'@
$beforePids = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$sourceShaBefore = (Get-FileHash -LiteralPath $SourcePath -Algorithm SHA256).Hash.ToLowerInvariant()
$singleText = [IO.File]::ReadAllText($SingleFilePath, [Text.Encoding]::UTF8).TrimStart([char]0xFEFF)
$singleText = $singleText.Replace("`r`n","`n").Replace("`r","`n").Replace("`n","`r`n")
$result = [ordered]@{
    testedAtUtc=[DateTime]::UtcNow.ToString('o'); status='RUNNING';
    txtSha256=(Get-FileHash -LiteralPath $SingleFilePath -Algorithm SHA256).Hash.ToLowerInvariant();
    sourceSha256=$sourceShaBefore; powershellVersion=$PSVersionTable.PSVersion.ToString();
    existingExcelPids=$beforePids; ownedExcelPid=$null; factoryAttempts=0;
    navigationServerStarted=$false; browserExcelEndToEndVerified=$false;
    verificationScope='VBA paste, source selection, HTML generation, queue-to-Excel navigation component';
    securityPolicyChanged=$false; sourceFileSaved=$false; cleanupErrors=@()
}
$reportPath = Join-Path $OutputDirectory 'windows_excel_smoke.json'
$harness = @'
Public Function QA_Resolve(ByVal choice As Long, ByVal currentName As String, ByVal selectedPath As String) As String
    Dim book As Workbook
    On Error GoTo Failed
    Set book = FM_ResolveSource(choice, currentName, selectedPath)
    QA_Resolve = book.Name & Chr$(30) & book.FullName & Chr$(30) & CStr(book.ReadOnly)
    Exit Function
Failed:
    QA_Resolve = "ERROR" & Chr$(30) & CStr(Err.Number)
End Function
Public Function QA_Template() As String
    QA_Template = FM_Template()
End Function
Public Function QA_Bridge() As String
    QA_Bridge = FM_BridgeScript()
End Function
Public Function QA_ProgressBusy() As Boolean
    QA_ProgressBusy = mFMProgressBusy
End Function
Public Function QA_NavigationState() As String
    Dim state As Object
    Set state = FM_NavState()
    QA_NavigationState = CStr(state("active")) & "|" & CStr(state("scheduled"))
End Function
'@
$excel = $null; $hostBook = $null; $sourceBook = $null
$owned = $false; $prefix = $null
try {
    $result.factoryAttempts = 1
    [IO.File]::WriteAllText($reportPath, ($result | ConvertTo-Json -Depth 12), $plainUtf8)
    $excel = New-Object -ComObject Excel.Application
    [uint32]$ownerPid = 0
    [void][AtlasSmokeWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$ownerPid)
    $owned = ($ownerPid -gt 0 -and $ownerPid -notin $beforePids)
    if (-not $owned) { throw 'Cannot prove a new Excel process is owned by this test. No workbook was opened.' }
    $result.ownedExcelPid = [int]$ownerPid
    $result.excelVersion = [string]$excel.Version
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    # Disable macros when opening the sample. This is an instance setting,
    # not a Trust Center / registry change and is never changed on user Excel.
    $excel.AutomationSecurity = 3
    $hostBook = $excel.Workbooks.Open($SourcePath, 0, $true)
    $originalSnapshot = Book-Snapshot $hostBook
    $hostPath = Join-Path $OutputDirectory 'DisposableAtlasHost.xlsm'
    $hostBook.SaveAs($hostPath, 52)
    try {
        $component = $hostBook.VBProject.VBComponents.Add(1)
    } catch {
        throw ('VBA project access is unavailable under the current Excel policy. The test does not change policy. ' + $_.Exception.Message)
    }
    $component.Name = 'OfficePastedModule'
    $component.CodeModule.AddFromString($singleText + "`r`n" + $harness)
    $hostBook.Save()
    Add-Check 'single text pastes into an arbitrary standard module' ($component.CodeModule.CountOfLines -gt 1000)
    $prefix = "'" + $hostBook.Name.Replace("'","''") + "'!OfficePastedModule."
    $resolvedCurrent = [string]$excel.Run(($prefix + 'QA_Resolve'), 1, $hostBook.Name, '')
    $currentParts = $resolvedCurrent.Split([char]30)
    Add-Check 'current-workbook choice resolves the code host' ($currentParts.Length -eq 3 -and $currentParts[1] -ceq $hostBook.FullName)
    $appBefore = App-Snapshot $excel
    $resolvedFile = [string]$excel.Run(($prefix + 'QA_Resolve'), 2, $hostBook.Name, $SourcePath)
    $fileParts = $resolvedFile.Split([char]30)
    Add-Check 'selected-file choice opens the sample read-only' ($fileParts.Length -eq 3 -and $fileParts[1] -ceq $SourcePath -and $fileParts[2] -ceq 'True')
    $sourceBook = $excel.Workbooks.Item([IO.Path]::GetFileName($SourcePath))
    Add-Check 'source resolver restores application settings' (($appBefore | ConvertTo-Json -Compress) -ceq ((App-Snapshot $excel) | ConvertTo-Json -Compress))
    $countBeforeReuse = $excel.Workbooks.Count
    $resolvedAgain = [string]$excel.Run(($prefix + 'QA_Resolve'), 2, $hostBook.Name, $SourcePath)
    Add-Check 'already-open selected source is reused' ($resolvedAgain -ceq $resolvedFile -and $excel.Workbooks.Count -eq $countBeforeReuse)
    $invalidChoice = [string]$excel.Run(($prefix + 'QA_Resolve'), 0, $hostBook.Name, '')
    Add-Check 'invalid source choice is rejected' ($invalidChoice -ceq ('ERROR' + [char]30 + '5'))
    $template = [string]$excel.Run(($prefix + 'QA_Template'))
    $bridge = [string]$excel.Run(($prefix + 'QA_Bridge'))
    $result.decodedHtmlSha256 = Text-Sha $template
    $result.decodedBridgeSha256 = Text-Sha $bridge
    Add-Check 'embedded HTML matches its source' ($result.decodedHtmlSha256 -ceq (Text-Sha ([IO.File]::ReadAllText($TemplatePath))))
    Add-Check 'embedded helper matches its source' ($result.decodedBridgeSha256 -ceq (Text-Sha ([IO.File]::ReadAllText($BridgePath))))
    $hostBefore = Book-Snapshot $hostBook
    $sourceBefore = Book-Snapshot $sourceBook
    $hostSavedBefore = $hostBook.Saved; $sourceSavedBefore = $sourceBook.Saved
    $appBeforeBuild = App-Snapshot $excel
    $bookCountBefore = $excel.Workbooks.Count
    # This entry point generates HTML only. ExcelMap_Start and
    # StartFormulaMapNavigation are deliberately never called by this test.
    $htmlPath = [string]$excel.Run(($prefix + 'FM_BuildFile'), $hostBook.Name, $OutputDirectory, $true, $false)
    $html = [IO.File]::ReadAllText($htmlPath, [Text.Encoding]::UTF8)
    $dataMatch = [regex]::Match($html, '(?s)<script\b[^>]*\bid="formulaMapData"[^>]*>(.*?)</script>')
    if (-not $dataMatch.Success) { throw 'Generated HTML has no formulaMapData JSON block.' }
    $data = $dataMatch.Groups[1].Value | ConvertFrom-Json
    $result.htmlFile = [IO.Path]::GetFileName($htmlPath)
    $result.dataCounts = @{ formulas=$data.formulaCount; totalFormulas=$data.totalFormulaCount;
        worksheets=@($data.worksheets).Count; nodes=@($data.nodes).Count; edges=@($data.edges).Count }
    Add-Check 'analysis identifies the intended workbook' ($data.sourceName -ceq $hostBook.Name)
    Add-Check 'HTML contains formulas, reference links and all worksheets' ($data.formulaCount -gt 0 -and @($data.nodes).Count -gt 0 -and @($data.edges).Count -gt 0 -and @($data.worksheets).Count -eq $hostBook.Worksheets.Count)
    Add-Check 'HTML-only analysis creates no extra workbook' ($excel.Workbooks.Count -eq $bookCountBefore)
    Add-Check 'host content stays unchanged' ($hostBefore -ceq (Book-Snapshot $hostBook) -and $hostBefore -ceq $originalSnapshot)
    Add-Check 'original read-only content stays unchanged' ($sourceBefore -ceq (Book-Snapshot $sourceBook))
    Add-Check 'analysis preserves workbook Saved flags' ($hostBook.Saved -eq $hostSavedBefore -and $sourceBook.Saved -eq $sourceSavedBefore)
    Add-Check 'analysis restores application settings' (($appBeforeBuild | ConvertTo-Json -Compress) -ceq ((App-Snapshot $excel) | ConvertTo-Json -Compress))
    Add-Check 'analysis releases its reentry guard' (-not [bool]$excel.Run(($prefix + 'QA_ProgressBusy')))
    $target = @($data.nodes | Where-Object { $_.book -ceq $hostBook.Name -and $_.kind -ne 'external' -and $_.r1 -and $hostBook.Worksheets.Item($_.sheet).Visible -eq -1 }) | Select-Object -First 1
    if ($null -eq $target) { throw 'No visible exported navigation target was found.' }
    foreach ($targetCase in @(
        @{ sheet=$target.sheet; address=$target.address; expected=$true },
        @{ sheet=$target.sheet; address='XFD1048576'; expected=$false }
    )) {
        # Prepare writes temporary configuration/helper text but does not run it.
        $session = [string]$excel.Run(($prefix + 'FM_PrepareNavigationSession'), $htmlPath)
        $sessions.Add($session)
        $requestId = [Guid]::NewGuid().ToString('N')
        $sheet64 = [Convert]::ToBase64String($plainUtf8.GetBytes($targetCase.sheet))
        [IO.File]::WriteAllText((Join-Path $session 'request.txt'), ($requestId + '|' + $sheet64 + '|' + $targetCase.address), $plainUtf8)
        $excel.Run(($prefix + 'FM_PollNavigation')) | Out-Null
        $ack = [IO.File]::ReadAllText((Join-Path $session 'ack.json')) | ConvertFrom-Json
        $excel.Run(($prefix + 'FM_StopNavigation')) | Out-Null
        $okay = ($ack.id -ceq $requestId -and $ack.ok -eq $targetCase.expected)
        if ($targetCase.expected) {
            $okay = $okay -and $excel.ActiveWorkbook.FullName -ceq $hostBook.FullName -and $excel.ActiveSheet.Name -ceq $targetCase.sheet -and $excel.Selection.Address($false,$false) -ceq $targetCase.address
            Add-Check 'queue request selects the exact exported Excel cell or range' $okay @{ sheet=$targetCase.sheet; address=$targetCase.address; ack=$ack }
        } else {
            Add-Check 'queue rejects a location outside the exported map' ($okay -and $ack.reason -ceq 'not_in_exported_map') $ack
        }
        Add-Check 'Stop cancels the navigation timer' ([string]$excel.Run(($prefix + 'QA_NavigationState')) -ceq 'False|False')
    }
    $excel.Run(($prefix + 'FM_StopNavigation')) | Out-Null
    Add-Check 'Stop is safe when repeated' $true
    $result.status = 'COMPLETE'
} catch {
    $result.status = 'NOT_EXECUTED_OR_FAILED'
    $result.executionError = $_.Exception.ToString()
    Add-Check 'runtime test finishes without an execution error' $false $_.Exception.Message
} finally {
    if ($owned -and $null -ne $excel) {
        if ($prefix) { try { $excel.Run(($prefix + 'FM_StopNavigation')) | Out-Null } catch {} }
        if ($null -ne $sourceBook) { try { $sourceBook.Close($false) } catch { $result.cleanupErrors += $_.Exception.Message } }
        if ($null -ne $hostBook) { try { $hostBook.Close($false) } catch { $result.cleanupErrors += $_.Exception.Message } }
        try { $excel.Quit() } catch { $result.cleanupErrors += $_.Exception.Message }
    }
    if ($null -ne $excel) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    Add-Check 'original source bytes are unchanged' ($sourceShaBefore -ceq (Get-FileHash -LiteralPath $SourcePath -Algorithm SHA256).Hash.ToLowerInvariant())
    $result.tests = $checks
    $result.passed = @($checks | Where-Object pass).Count
    $result.failed = @($checks | Where-Object { -not $_.pass }).Count
    $result.navigationSessionFolders = @($sessions)
    $result.existingExcelPidsStillRunning = @($beforePids | Where-Object { Get-Process -Id $_ -ErrorAction SilentlyContinue })
    # Keep QA output and session folders for inspection. No recursive deletion.
    [IO.File]::WriteAllText($reportPath, ($result | ConvertTo-Json -Depth 14), $plainUtf8)
    @{ status=$result.status; passed=$result.passed; failed=$result.failed; report=$reportPath } | ConvertTo-Json
}
if ($result.failed -gt 0 -or $result.cleanupErrors.Count -gt 0) { exit 1 }
