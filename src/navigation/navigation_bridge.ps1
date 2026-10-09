# Temporary, loopback-only transport. Excel owns all navigation decisions.
# This helper neither automates Excel nor changes system/security settings.
[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$ConfigPath)
$ErrorActionPreference = 'Stop'
$listener = $null
$sessionDirectory = [IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($ConfigPath))
$started = $false
$utf8 = New-Object System.Text.UTF8Encoding($false, $true)

function Write-Utf8Atomic([string]$Path, [string]$Value) {
    $temporary = $Path + '.writing'
    [IO.File]::WriteAllText($temporary, $Value, $utf8)
    Move-Item -LiteralPath $temporary -Destination $Path -Force
}

function Send-Reply($Stream, [int]$Status, [string]$ContentType, [string]$Value) {
    $labels = @{200='OK';400='Bad Request';403='Forbidden';404='Not Found';405='Method Not Allowed';408='Request Timeout';409='Conflict';413='Payload Too Large';415='Unsupported Media Type';431='Request Header Fields Too Large';503='Service Unavailable'}
    $body = $utf8.GetBytes($Value)
    $header = "HTTP/1.1 $Status $($labels[$Status])`r`nContent-Type: $ContentType`r`nContent-Length: $($body.Length)`r`nConnection: close`r`nCache-Control: no-store`r`nX-Content-Type-Options: nosniff`r`nReferrer-Policy: no-referrer`r`nContent-Security-Policy: default-src 'self' data:; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; connect-src 'self'; img-src 'self' data:; object-src 'none'; frame-ancestors 'none'; base-uri 'none'`r`n`r`n"
    $headBytes = [Text.Encoding]::ASCII.GetBytes($header)
    $Stream.Write($headBytes,0,$headBytes.Length)
    $Stream.Write($body,0,$body.Length)
    $Stream.Flush()
}

function Send-Json($Stream, [int]$Status, $Object) {
    Send-Reply $Stream $Status 'application/json; charset=utf-8' (ConvertTo-Json -InputObject $Object -Compress -Depth 5)
}

function Get-ExcelStatus {
    if ([DateTime]::UtcNow -ge $expires) { return @{alive=$false;reason='session_expired'} }
    if (Test-Path -LiteralPath (Join-Path $sessionDirectory 'stop')) { return @{alive=$false;reason='session_stopped'} }
    $heartbeat = Join-Path $sessionDirectory 'heartbeat'
    if (-not (Test-Path -LiteralPath $heartbeat)) { return @{alive=$false;reason='excel_disconnected'} }
    if (([DateTime]::UtcNow - (Get-Item -LiteralPath $heartbeat).LastWriteTimeUtc).TotalSeconds -gt 45) { return @{alive=$false;reason='excel_disconnected'} }
    try {
        $state = [IO.File]::ReadAllText((Join-Path $sessionDirectory 'state.json'), $utf8) | ConvertFrom-Json
        return @{alive=[bool]$state.alive;reason=[string]$state.reason}
    } catch { return @{alive=$true;reason=''} }
}

function Read-Request($Stream) {
    $buffer = New-Object 'System.Collections.Generic.List[byte]'
    $match = 0
    while ($true) {
        $b = $Stream.ReadByte()
        if ($b -lt 0) { throw 'incomplete_header' }
        $buffer.Add([byte]$b)
        if ($buffer.Count -gt 8192) { return @{error=431;reason='headers_too_large'} }
        if (($match -eq 0 -or $match -eq 2) -and $b -eq 13) { $match++ }
        elseif (($match -eq 1 -or $match -eq 3) -and $b -eq 10) { $match++ }
        else { $match = 0 }
        if ($match -eq 4) { break }
    }
    $lines = [Text.Encoding]::ASCII.GetString($buffer.ToArray()).Split(@("`r`n"),[StringSplitOptions]::None)
    if ($lines[0] -notmatch '^(GET|POST|[A-Z]+) (/[^ ]*) HTTP/1\.[01]$') { return @{error=400;reason='invalid_request'} }
    $method=$Matches[1]; $path=$Matches[2]
    $headers = @{}
    for ($i=1; $i -lt $lines.Length - 2; $i++) {
        $line=$lines[$i]
        if ($line -notmatch '^([A-Za-z0-9-]+):[ \t]*(.*)$') { return @{error=400;reason='invalid_header'} }
        $name=$Matches[1].ToLowerInvariant(); $value=$Matches[2].Trim()
        if ($headers.ContainsKey($name)) { return @{error=400;reason='duplicate_header'} }
        $headers[$name]=$value
    }
    if ($headers['host'] -cne "127.0.0.1:$port") { return @{error=403;reason='invalid_host'} }
    if ($headers.ContainsKey('origin') -and $headers['origin'] -cne $origin) { return @{error=403;reason='invalid_origin'} }
    if ($headers.ContainsKey('transfer-encoding')) { return @{error=400;reason='unsupported_transfer_encoding'} }
    $length=0
    if ($headers.ContainsKey('content-length')) {
        if ($headers['content-length'] -notmatch '^\d{1,7}$') { return @{error=400;reason='invalid_content_length'} }
        $length=[int]$headers['content-length']
    }
    if ($length -gt 4096) { return @{error=413;reason='payload_too_large'} }
    if ($method -eq 'POST') {
        if ($headers['origin'] -cne $origin) { return @{error=403;reason='origin_required'} }
        if ($headers['content-type'] -notmatch '^application/json(?:\s*;\s*charset=utf-8)?$') { return @{error=415;reason='json_required'} }
        if ($length -eq 0) { return @{error=400;reason='empty_payload'} }
    } elseif ($length -ne 0) { return @{error=400;reason='unexpected_body'} }
    $bytes = New-Object byte[] $length
    $offset=0
    while ($offset -lt $length) {
        $read=$Stream.Read($bytes,$offset,$length-$offset)
        if ($read -le 0) { throw 'incomplete_body' }
        $offset += $read
    }
    return @{method=$method;path=$path;body=$utf8.GetString($bytes)}
}

try {
    # Keep token/config/commands private to the signed-in Windows user.
    $acl = New-Object Security.AccessControl.DirectorySecurity
    $acl.SetAccessRuleProtection($true,$false)
    $sid=[Security.Principal.WindowsIdentity]::GetCurrent().User
    $rule=New-Object Security.AccessControl.FileSystemAccessRule($sid,'FullControl','ContainerInherit,ObjectInherit','None','Allow')
    $acl.AddAccessRule($rule)
    Set-Acl -LiteralPath $sessionDirectory -AclObject $acl
    $config=[IO.File]::ReadAllText($ConfigPath,$utf8)|ConvertFrom-Json
    if ([IO.Path]::GetFullPath([string]$config.SessionDirectory) -cne [IO.Path]::GetFullPath($sessionDirectory)) { throw 'Session directory mismatch.' }
    $port=[int]$config.Port
    if ($port -lt 40000 -or $port -gt 60000) { throw 'Port must be between 40000 and 60000.' }
    $origin="http://127.0.0.1:$port"
    $token=[Guid]::NewGuid().ToString('N')+[Guid]::NewGuid().ToString('N')
    $base="/s/$token"
    $expires=[DateTime]::UtcNow.AddSeconds([Math]::Min(7200,[Math]::Max(1,[int]$config.LifetimeSeconds)))
    $allowed=New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
    foreach ($node in $config.Allowed) {
        $sheet=[string]$node.sheet; $address=([string]$node.address).Replace('$','').ToUpperInvariant()
        if ($sheet.Length -le 31 -and $address -match '^(?:[A-Z]{1,3}[1-9][0-9]{0,6}(?::[A-Z]{1,3}[1-9][0-9]{0,6})?|[A-Z]{1,3}:[A-Z]{1,3}|[1-9][0-9]{0,6}:[1-9][0-9]{0,6})$') {
            [void]$allowed.Add($sheet+[char]30+$address)
        }
    }
    $html=[IO.File]::ReadAllText([string]$config.HtmlPath,$utf8)
    $navigation=@{base=$base;expires=$expires.ToString('o');sourceName=[string]$config.SourceName}|ConvertTo-Json -Compress
    $injection='<script>window.FM_NAVIGATION='+$navigation+';</script>'
    # Before every original script, including the main script and JSON data block.
    $firstScript=$html.IndexOf('<script',[StringComparison]::OrdinalIgnoreCase)
    if ($firstScript -lt 0) { throw 'The map has no script section.' }
    $html=$html.Insert($firstScript,$injection)
    $listener=New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback,$port)
    $listener.Start()
    $started=$true
    Write-Utf8Atomic (Join-Path $sessionDirectory 'ready.txt') ($origin+$base+'/')
    while ([DateTime]::UtcNow -lt $expires) {
        if (Test-Path -LiteralPath (Join-Path $sessionDirectory 'stop')) { break }
        $heartbeat=Join-Path $sessionDirectory 'heartbeat'
        if (-not (Test-Path -LiteralPath $heartbeat)) { break }
        if (([DateTime]::UtcNow-(Get-Item -LiteralPath $heartbeat).LastWriteTimeUtc).TotalSeconds -gt 45) { break }
        if (-not $listener.Pending()) { Start-Sleep -Milliseconds 50; continue }
        $client=$listener.AcceptTcpClient()
        try {
            $client.ReceiveTimeout=5000; $client.SendTimeout=5000
            $stream=$client.GetStream()
            $request=Read-Request $stream
            if ($request.ContainsKey('error')) { Send-Json $stream $request.error @{ok=$false;reason=$request.reason}; continue }
            $path=$request.path
            if ($path -cne "$base/" -and $path -cne "$base/health" -and $path -cne "$base/goto") { Send-Json $stream 404 @{ok=$false;reason='route_not_found'}; continue }
            if ($path -ceq "$base/") {
                if ($request.method -cne 'GET') { Send-Json $stream 405 @{ok=$false;reason='get_required'}; continue }
                Send-Reply $stream 200 'text/html; charset=utf-8' $html
                continue
            }
            if ($path -ceq "$base/health") {
                if ($request.method -cne 'GET') { Send-Json $stream 405 @{ok=$false;reason='get_required'}; continue }
                $status=Get-ExcelStatus
                $status['ok']=[bool]$status.alive; $status['expires']=$expires.ToString('o'); $status['sourceName']=[string]$config.SourceName
                Send-Json $stream 200 $status
                continue
            }
            if ($request.method -cne 'POST') { Send-Json $stream 405 @{ok=$false;reason='post_required'}; continue }
            $status=Get-ExcelStatus
            if (-not $status.alive) { Send-Json $stream 409 @{ok=$false;reason=$status.reason}; continue }
            try { $payload=ConvertFrom-Json -InputObject $request.body -ErrorAction Stop } catch { Send-Json $stream 400 @{ok=$false;reason='invalid_json'}; continue }
            if ($null -eq $payload -or $payload -is [array] -or @($payload.PSObject.Properties).Count -ne 2 -or $payload.sheet -isnot [string] -or $payload.address -isnot [string]) { Send-Json $stream 400 @{ok=$false;reason='invalid_fields'}; continue }
            $sheet=[string]$payload.sheet; $address=([string]$payload.address).Replace('$','').ToUpperInvariant()
            if (-not $allowed.Contains($sheet+[char]30+$address)) { Send-Json $stream 403 @{ok=$false;reason='not_in_exported_map'}; continue }
            $id=[Guid]::NewGuid().ToString('N')
            $encodedSheet=[Convert]::ToBase64String($utf8.GetBytes($sheet))
            $requestFile=Join-Path $sessionDirectory 'request.txt'; $ackFile=Join-Path $sessionDirectory 'ack.json'
            Write-Utf8Atomic $requestFile ($id+'|'+$encodedSheet+'|'+$address)
            $deadline=[DateTime]::UtcNow.AddSeconds(8); $ack=$null
            while ([DateTime]::UtcNow -lt $deadline) {
                if (Test-Path -LiteralPath (Join-Path $sessionDirectory 'stop')) { break }
                if (Test-Path -LiteralPath $ackFile) {
                    try { $candidate=[IO.File]::ReadAllText($ackFile,$utf8)|ConvertFrom-Json; if ($candidate.id -ceq $id) { $ack=$candidate; break } } catch { }
                }
                Start-Sleep -Milliseconds 30
            }
            if ($null -ne $ack) { Send-Json $stream 200 $ack }
            else {
                if (Test-Path -LiteralPath $requestFile) { Remove-Item -LiteralPath $requestFile -Force }
                Send-Json $stream 503 @{ok=$false;reason='excel_busy_or_disconnected'}
            }
        } catch {
            try { if ($null -ne $stream) { Send-Json $stream 400 @{ok=$false;reason='request_failed'} } } catch { }
        } finally { $client.Close() }
    }
} catch {
    try { Write-Utf8Atomic (Join-Path $sessionDirectory 'startup-error.txt') $_.Exception.Message } catch { }
    exit 1
} finally {
    if ($null -ne $listener) { $listener.Stop() }
    # Ephemeral transport files contain workbook names and session secrets.
    # The generated HTML/XLSX and all original workbooks remain outside this folder.
    $intendedParent=[IO.Path]::GetFullPath([IO.Path]::Combine([IO.Path]::GetTempPath(),'ExcelSpatialMap')).TrimEnd('\')
    $resolvedParent=[IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($sessionDirectory)).TrimEnd('\')
    if ($started -and [IO.Path]::GetFileName($sessionDirectory) -like '*.tmp' -and [string]::Equals($resolvedParent,$intendedParent,[StringComparison]::OrdinalIgnoreCase)) {
        try { Remove-Item -LiteralPath $sessionDirectory -Recurse -Force } catch { }
    }
}
