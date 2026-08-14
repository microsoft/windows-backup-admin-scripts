# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

<#
  View-WindowsBackup.ps1
  Standalone admin script to VIEW Windows Backup for Organizations data.

  What it does:
    - Signs you in (handles scopes automatically).
    - Views your own backup (press Enter) or another user's (type their UPN).
    - Shows a summary: total devices/settings/size, then a per-device rollup.
    - Lets you drill into one device to see its settings.
    - Optionally exports the data to a JSON file (Y/N).

  Requirements:
    - Windows PowerShell 5.1 OR PowerShell 7 (both supported).
    - Microsoft.Graph.Authentication module (auto-installed).
    - To view ANOTHER user: Microsoft 365 Backup Administrator role + admin consent
      for UserWindowsSettings.Read.All and User.Read.All.

  Run (works on either PowerShell version):
    powershell -ExecutionPolicy Bypass -File .\View-WindowsBackup.ps1
      -- or --
    pwsh -ExecutionPolicy Bypass -File .\View-WindowsBackup.ps1
#>

param(
    [string] $UserId,                       # UPN or object id; blank = your own backup
    [string] $ExportPath = $env:USERPROFILE, # where a JSON export is saved if you choose Y
    [int]    $TopDevices = 15,               # how many devices to list before summarizing the rest
    [switch] $Fresh                          # force a clean sign-in to switch accounts
)

$ErrorActionPreference = 'Stop'

# ============================================================
#  STEP 1: Check PowerShell version (works on 5.1 and 7)
# ============================================================
Write-Host "[1/4] PowerShell $($PSVersionTable.PSVersion) detected." -ForegroundColor Green
if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Host "      This script needs Windows PowerShell 5.1 or PowerShell 7. Please upgrade." -ForegroundColor Red
    return
}

# ============================================================
#  STEP 2: Ensure the Graph auth module is available
# ============================================================
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Write-Host "[2/4] Installing Microsoft.Graph.Authentication module (one-time)..." -ForegroundColor Cyan
    # Make the install fully non-interactive (no NuGet / untrusted-repo prompts on 5.1).
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
    try { Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch {}
    # Trust PSGallery only for this install, then restore the previous policy so we don't
    # permanently change the user's machine state.
    $prevGalleryPolicy = $null
    try { $prevGalleryPolicy = (Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue).InstallationPolicy } catch {}
    try {
        if ($prevGalleryPolicy -and $prevGalleryPolicy -ne 'Trusted') {
            try { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue } catch {}
        }
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber -Confirm:$false -Repository PSGallery
    }
    finally {
        if ($prevGalleryPolicy -and $prevGalleryPolicy -ne 'Trusted') {
            try { Set-PSRepository -Name PSGallery -InstallationPolicy $prevGalleryPolicy -ErrorAction SilentlyContinue } catch {}
        }
    }
}
else {
    Write-Host "[2/4] Microsoft.Graph.Authentication module present." -ForegroundColor Green
}
Import-Module Microsoft.Graph.Authentication

# --- -Fresh: drop any cached session so you can sign in as a different account ---
if ($Fresh) {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Write-Host "Cleared cached sign-in. You'll sign in fresh." -ForegroundColor Cyan
}

# --- Connectivity pre-check: verify we can reach Microsoft sign-in + Graph.
#     Explains the reason clearly instead of a cryptic auth/retry error. ---
function Test-GraphConnectivity {
    $endpoints = @(
        @{ Host = 'login.microsoftonline.com'; Purpose = 'Microsoft sign-in' },
        @{ Host = 'graph.microsoft.com';       Purpose = 'Microsoft Graph API' }
    )
    $blocked = @()
    foreach ($e in $endpoints) {
        $ok = $false
        try {
            $client = New-Object System.Net.Sockets.TcpClient
            $iar = $client.BeginConnect($e.Host, 443, $null, $null)
            $ok = $iar.AsyncWaitHandle.WaitOne(5000, $false) -and $client.Connected
            $client.Close()
        } catch { $ok = $false }
        if (-not $ok) { $blocked += $e }
    }

    if ($blocked.Count -gt 0) {
        # Is general internet up? (helps explain WHY)
        $netUp = $false
        try {
            $c = New-Object System.Net.Sockets.TcpClient
            $i = $c.BeginConnect('www.microsoft.com', 443, $null, $null)
            $netUp = $i.AsyncWaitHandle.WaitOne(5000, $false) -and $c.Connected
            $c.Close()
        } catch { $netUp = $false }

        Write-Host ""
        Write-Host "=================== CANNOT REACH MICROSOFT SIGN-IN ===================" -ForegroundColor Red
        foreach ($b in $blocked) {
            Write-Host (" X  {0}:443  ({1}) is not reachable" -f $b.Host, $b.Purpose) -ForegroundColor Red
        }
        Write-Host ""
        if ($netUp) {
            Write-Host " Your machine HAS internet, but these Microsoft endpoints are blocked." -ForegroundColor Yellow
            Write-Host " This is usually a firewall or proxy restriction on this machine/network." -ForegroundColor Yellow
        } else {
            Write-Host " Your machine appears to have NO outbound internet access." -ForegroundColor Yellow
        }
        Write-Host ""
        Write-Host " WHY you see a sign-in error: signing in requires reaching" -ForegroundColor Yellow
        Write-Host " login.microsoftonline.com (Azure AD). Without it, NO tool can sign in" -ForegroundColor Yellow
        Write-Host " (this affects Connect-MgGraph, Graph Explorer, and the SDK too - not just this script)." -ForegroundColor Yellow
        Write-Host ""
        Write-Host " TO FIX: ask your network/IT admin to allow, on port 443:" -ForegroundColor Cyan
        Write-Host "     login.microsoftonline.com" -ForegroundColor Cyan
        Write-Host "     graph.microsoft.com" -ForegroundColor Cyan
        Write-Host "   If your network uses a proxy, configure it, e.g.:" -ForegroundColor Cyan
        Write-Host '     netsh winhttp set proxy proxy-server="http://YOURPROXY:PORT"' -ForegroundColor DarkGray
        Write-Host "======================================================================" -ForegroundColor Red
        return $false
    }
    return $true
}

# --- Sign in with the scopes needed (reuse session if it already has them).
#     Interactive (browser) first; device code only if the browser sign-in fails. ---
function Connect-Backup([string[]] $Scopes) {
    if (-not $script:Fresh) {
        $ctx = Get-MgContext
        $missing = if ($ctx) { $Scopes | Where-Object { $_ -notin $ctx.Scopes } } else { $Scopes }
        if ($ctx -and -not $missing) { return $ctx }
    }

    # Try interactive (browser/WAM) first; fall back to device code if it fails.
    # No upfront network check - we only diagnose connectivity if sign-in actually fails,
    # so the normal (working) path stays fast.
    try {
        Connect-MgGraph -Scopes $Scopes -NoWelcome -ErrorAction Stop -WarningAction SilentlyContinue
    }
    catch {
        try {
            Write-Host "Browser sign-in unavailable - use the device code below to sign in:" -ForegroundColor Cyan
            Connect-MgGraph -Scopes $Scopes -NoWelcome -UseDeviceAuthentication -ErrorAction Stop
        }
        catch {
            # Both sign-in methods failed - diagnose WHY (network/proxy) now.
            Test-GraphConnectivity | Out-Null
            throw "Sign-in failed. See the guidance above."
        }
    }

    $ctx = Get-MgContext
    if (-not $ctx -or -not $ctx.Account) {
        throw "Sign-in did not complete. Please try again."
    }
    return $ctx
}

# --- GET helper that follows paging ---
function Get-AllValues([string] $Uri) {
    $items = [System.Collections.Generic.List[object]]::new()
    $next = $Uri
    while ($next) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $next -ErrorAction Stop
        if ($resp.value) { foreach ($v in $resp.value) { $items.Add($v) } }
        $next = $resp.'@odata.nextLink'
    }
    return $items
}

# --- Build a deviceId -> display-name map straight from the backup's own
#     'deviceprofile' payload (profileId + deviceDisplayName). No extra scope,
#     and it works even for devices not registered in Entra/Intune. ---
function Get-DeviceNameMap([object[]] $Settings) {
    $map = @{}
    foreach ($s in $Settings) {
        if ($s.payloadType -like '*platform.backuprestore.deviceprofile*') {
            $inst = $s.instances | Select-Object -First 1
            if ($inst.payload) {
                try {
                    $json   = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($inst.payload)) | ConvertFrom-Json
                    $profId = ($json.profileId -replace '[{}]', '')   # strip braces -> matches windowsDeviceId
                    if ($profId -and $json.deviceDisplayName) { $map[$profId] = $json.deviceDisplayName }
                } catch { }
            }
        }
    }
    return $map
}

# --- Explain a 403 in plain language: it's a role/consent problem, not a bug.
#     Runs additional checks on the current token's scopes to show what's missing. ---
function Show-AccessDeniedHelp([string] $Operation, [string[]] $RequiredScopes, [string] $Account) {
    Write-Host ""
    Write-Host "=================== ACCESS DENIED (HTTP 403) ===================" -ForegroundColor Red
    Write-Host (" Signed in as: {0}" -f $Account) -ForegroundColor Yellow
    Write-Host (" This account is NOT authorized to {0}." -f $Operation) -ForegroundColor Yellow
    Write-Host ""
    Write-Host " Cross-user backup access requires BOTH of the following:" -ForegroundColor Cyan
    Write-Host "   1. ROLE : Microsoft 365 Backup Administrator - MUST be explicitly assigned" -ForegroundColor Cyan
    Write-Host "            to this account. Global Administrator alone does NOT grant it." -ForegroundColor Cyan
    Write-Host "   2. Admin-consented SCOPES:" -ForegroundColor Cyan
    $ctx  = Get-MgContext
    $have = if ($ctx) { $ctx.Scopes } else { @() }
    foreach ($s in $RequiredScopes) {
        $granted = $s -in $have
        $mark    = if ($granted) { 'granted' } else { 'MISSING' }
        $color   = if ($granted) { 'Green' }   else { 'Red' }
        Write-Host ("        - {0}  [{1}]" -f $s, $mark) -ForegroundColor $color
    }
    Write-Host ""
    Write-Host " Note: this API checks EXPLICIT role membership, not Global Admin privilege." -ForegroundColor Yellow
    Write-Host " Even Global Administrators get 403 here unless the Microsoft 365 Backup" -ForegroundColor Yellow
    Write-Host " Administrator role is directly assigned to their account. A correctly scoped" -ForegroundColor Yellow
    Write-Host " token is also not enough on its own - the role is enforced separately." -ForegroundColor Yellow
    Write-Host " Fix: assign the Microsoft 365 Backup Administrator role to this account AND" -ForegroundColor Cyan
    Write-Host " grant admin consent for the scopes above, then re-run with  -Fresh ." -ForegroundColor Cyan
    Write-Host "===============================================================" -ForegroundColor Red
}

# --- Ask who to view ---
if (-not $UserId) {
    $UserId = Read-Host "Enter the user UPN to view (or press Enter to view your own backup)"
}
if ($UserId) { $UserId = $UserId.Trim() }
$useMe = [string]::IsNullOrWhiteSpace($UserId)

# --- Build the request (self vs another user) ---
try {
    Write-Host "[3/4] Signing in to Microsoft Graph..." -ForegroundColor Cyan
    if ($useMe) {
        $ctx = Connect-Backup @('UserWindowsSettings.Read')
        $target = $ctx.Account
        $uri = 'https://graph.microsoft.com/beta/me/settings/windows'
    }
    else {
        $ctx = Connect-Backup @('UserWindowsSettings.Read.All', 'User.Read.All')
        # Resolve UPN -> object id (cross-user beta calls need {id}@{tenant})
        if ($UserId -match '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') { $objId = $UserId; $target = $UserId }
        else {
            $encodedUserId = [System.Uri]::EscapeDataString($UserId)  # Encode UPN for safe insertion into the URI path (e.g. '#')
            $u = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$encodedUserId`?`$select=id,userPrincipalName"
            $objId = $u.id; $target = $u.userPrincipalName
        }
        $uri = "https://graph.microsoft.com/beta/users/$objId@$($ctx.TenantId)/settings/windows"
    }
    Write-Host "      Signed in as $($ctx.Account)." -ForegroundColor Green
    Write-Host "[4/4] Reading backup for '$target'..." -ForegroundColor Cyan
    $settings = Get-AllValues $uri
}
catch {
    $msg = $_.Exception.Message
    $code = $null; try { $code = [int]$_.Exception.Response.StatusCode } catch {}
    $who  = if ($target) { $target } else { $UserId }
    if ($code -eq 403 -or $msg -match '403' -or $msg -match 'Forbidden') {
        $acct = try { (Get-MgContext).Account } catch { $null }
        $req  = if ($useMe) { @('UserWindowsSettings.Read') } else { @('UserWindowsSettings.Read.All', 'User.Read.All') }
        Show-AccessDeniedHelp ("read the backup for '$who'") $req $acct
    }
    else {
        Write-Host "Could not read backup for '$who': $msg" -ForegroundColor Red
    }
    return
}

if (-not $settings -or @($settings).Count -eq 0) {
    Write-Host "No backup data found for '$target'." -ForegroundColor Yellow
    return
}
Write-Host "Retrieved $(@($settings).Count) setting(s). Building view..." -ForegroundColor Green

# --- Flatten to rows ---
$rows = foreach ($s in $settings) {
    $inst = $s.instances | Select-Object -First 1
    [pscustomobject]@{
        Device       = $s.windowsDeviceId
        PayloadType  = $s.payloadType
        SettingType  = $s.settingType
        LastModified = $inst.lastModifiedDateTime
        Expiration   = $inst.expirationDateTime
        Bytes        = if ($inst.payload) { [System.Text.Encoding]::UTF8.GetByteCount($inst.payload) } else { 0 }
        Payload      = $inst.payload
    }
}

# --- Resolve device display names from the backup's own deviceprofile payload ---
$deviceNameMap = Get-DeviceNameMap $settings

# --- Per-device rollup ---
$byDevice = $rows | Group-Object Device | ForEach-Object {
    $devId = $_.Name
    $dName = if ($devId) {
                 if ($deviceNameMap[$devId]) { $deviceNameMap[$devId] } else { '(name unavailable)' }
             } else { '(account-level)' }
    [pscustomobject]@{
        DeviceName   = $dName
        DeviceId     = if ($devId) { $devId } else { '(account-level/default)' }
        SettingCount = $_.Count
        SizeKB       = [math]::Round((($_.Group.Bytes | Measure-Object -Sum).Sum)/1KB, 1)
        NewestMod    = ($_.Group.LastModified | Where-Object { $_ } | Sort-Object | Select-Object -Last 1)
    }
} | Sort-Object SizeKB -Descending

$totalKB = [math]::Round((($rows.Bytes | Measure-Object -Sum).Sum)/1KB, 1)

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host (" Backup for: {0}" -f $target) -ForegroundColor Cyan
Write-Host (" Devices: {0}    Settings: {1}    Total size: {2} KB" -f $byDevice.Count, $rows.Count, $totalKB) -ForegroundColor Cyan
Write-Host "============================================================`n" -ForegroundColor Cyan

Write-Host "-- Devices / profiles (largest first) --" -ForegroundColor Cyan
$byDevice | Select-Object -First $TopDevices |
    Format-Table DeviceName, DeviceId, SettingCount, SizeKB,
                 @{n='NewestMod';e={ if($_.NewestMod){([datetime]$_.NewestMod).ToString('yyyy-MM-dd')} }} -AutoSize

if ($byDevice.Count -gt $TopDevices) {
    $rest = $byDevice | Select-Object -Skip $TopDevices
    Write-Host ("   ... and {0} more device(s), {1} KB total." -f $rest.Count, (($rest.SizeKB | Measure-Object -Sum).Sum)) -ForegroundColor DarkGray
}

# --- Optional drill-down into one device ---
$drillRows = $null      # settings of the device the admin drilled into (if any)
$drillId   = $null
$pick = Read-Host "`nEnter a device id (full or first 8 chars) to VIEW its settings, or press Enter to skip"
if ($pick -and $pick.Trim()) {
    $p = $pick.Trim()
    $drillRows = $rows | Where-Object { $_.Device -and ($_.Device -eq $p -or $_.Device.StartsWith($p)) }
    if (-not $drillRows) { Write-Host "No device matched '$p'." -ForegroundColor Yellow }
    else {
        $drillId = ($drillRows | Select-Object -First 1).Device
        $drillName = if ($drillId -and $deviceNameMap[$drillId]) { $deviceNameMap[$drillId] } else { '(name unavailable)' }
        Write-Host "`n-- Settings for device $drillName [$drillId] ($($drillRows.Count)) --" -ForegroundColor Cyan
        $i = 0
        $drillRows | ForEach-Object {
            $i++
            [pscustomobject]@{
                '#'          = $i
                PayloadType  = ($_.PayloadType -replace '^windows\.data\.','')
                LastModified = $(if($_.LastModified){([datetime]$_.LastModified).ToString('yyyy-MM-dd HH:mm')})
                Bytes        = $_.Bytes
            }
        } | Format-Table -AutoSize

        # Dig deeper: decode a specific setting's payload to see its actual value.
        $sel = Read-Host "Enter a # (or part of a PayloadType) to DECODE its payload, or press Enter to skip"
        if ($sel -and $sel.Trim()) {
            $s = $sel.Trim()
            $one = if ($s -match '^\d+$') { $drillRows[[int]$s - 1] }
                   else { $drillRows | Where-Object { $_.PayloadType -like "*$s*" } | Select-Object -First 1 }
            if (-not $one) { Write-Host "No matching setting." -ForegroundColor Yellow }
            else {
                Write-Host "`n-- $($one.PayloadType) (decoded) --" -ForegroundColor Cyan
                if ($one.Payload) {
                    try {
                        $decoded = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($one.Payload))
                        if ($decoded.Length -gt 4000) { $decoded = $decoded.Substring(0,4000) + " ...(truncated)" }
                        $decoded
                    }
                    catch { Write-Host "<payload is binary - not readable as text>" -ForegroundColor DarkGray }
                }
                else { Write-Host "<no payload>" -ForegroundColor DarkGray }
            }
        }
    }
}

# --- Optional export: ask WHETHER + SCOPE first, then FORMAT ---
$doExport = (Read-Host "`nDo you want to export the data? (Y/N)").Trim().ToUpper()
$fmt = 'N'
$exportRows = $rows
$scopeLabel = 'all devices'
if ($doExport -in 'Y','YES') {
    # 1. Scope: type a device id for one, A for all, Enter to skip. Re-prompt until a valid choice.
    $scopeChosen = $false
    while (-not $scopeChosen) {
        $scopeInput = (Read-Host "Export scope - type a device id (full or first 8 chars) for just one, press A for ALL devices, or press Enter to skip export").Trim()
        if (-not $scopeInput) {
            # Enter = skip export entirely.
            $doExport = 'N'
            $scopeChosen = $true
        }
        elseif ($scopeInput.ToUpper() -eq 'A') {
            $exportRows = $rows
            $scopeLabel = 'all devices'
            $scopeChosen = $true
        }
        else {
            $match = $rows | Where-Object { $_.Device -and ($_.Device -eq $scopeInput -or $_.Device.StartsWith($scopeInput)) }
            if ($match) {
                $exportRows = $match
                $scopeLabel = "device $(($match | Select-Object -First 1).Device)"
                $scopeChosen = $true
            }
            else {
                Write-Host "No device matched '$scopeInput'. Press Enter to skip, A for ALL devices, or try a valid device id." -ForegroundColor Yellow
            }
        }
    }
    # 2. Format: which kind of export? (skipped if the admin chose to skip above)
    if ($doExport -in 'Y','YES') {
        Write-Host "`nExport format:" -ForegroundColor Cyan
        Write-Host "  [J] JSON - metadata only (device configuration, no payloads)"
        Write-Host "  [P] JSON - WITH decoded payloads (actual setting values)"
        Write-Host "  [D] DAT  - ZIP of raw .dat payload files, one per setting, per device"
        $fmt = (Read-Host "Choose J / P / D").Trim().ToUpper()
    }
}

if ($fmt -in 'J','P','D') {
    if (-not (Test-Path $ExportPath)) { New-Item -ItemType Directory -Force -Path $ExportPath | Out-Null }
    $safe  = ($target -replace '[^\w.@-]', '_')
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'

    if ($fmt -eq 'D') {
        # DAT: {deviceId}/{payloadType}.dat = decoded payload bytes, + manifest.json, zipped.
        $staging = Join-Path ([System.IO.Path]::GetTempPath()) "wba_$safe`_$stamp"
        New-Item -ItemType Directory -Force -Path $staging | Out-Null
        $manifest = foreach ($r in $exportRows) {
            $devFolder = if ($r.Device) { $r.Device } else { '_account-level' }
            $devDir = Join-Path $staging $devFolder
            if (-not (Test-Path $devDir)) { New-Item -ItemType Directory -Force -Path $devDir | Out-Null }
            $datName = (($r.PayloadType -replace '[^\w.-]','_') + '.dat')
            $datPath = Join-Path $devDir $datName
            $sizeB = 0
            if ($r.Payload) {
                try { $b = [Convert]::FromBase64String($r.Payload); [System.IO.File]::WriteAllBytes($datPath, $b); $sizeB = $b.Length }
                catch { Set-Content -LiteralPath $datPath -Value $r.Payload -Encoding Ascii }
            } else { New-Item -ItemType File -Force -Path $datPath | Out-Null }
            [pscustomobject]@{ device=$r.Device; payloadType=$r.PayloadType; lastModified=$r.LastModified; decodedBytes=$sizeB; file="$devFolder/$datName" }
        }
        [ordered]@{ exportedBy=$ctx.Account; exportedUtc=(Get-Date).ToUniversalTime().ToString('o'); targetUser=$target; scope=$scopeLabel; settingCount=@($exportRows).Count; entries=$manifest } |
            ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $staging 'manifest.json') -Encoding UTF8
        $zip = Join-Path $ExportPath "$safe`_windowsbackup_$stamp.zip"
        if (Test-Path $zip) { Remove-Item $zip -Force }
        Compress-Archive -Path (Join-Path $staging '*') -DestinationPath $zip -Force
        Remove-Item $staging -Recurse -Force -ErrorAction SilentlyContinue
        Write-Warning "DAT export contains decoded payload bytes (personal data). Store it securely."
        Write-Host "Exported ($scopeLabel) to: $zip" -ForegroundColor Green
    }
    else {
        # JSON: metadata-only (J) or with decoded payloads (P).
        $withPayload = ($fmt -eq 'P')
        $out = foreach ($r in $exportRows) {
            $o = [ordered]@{ Device=$r.Device; PayloadType=$r.PayloadType; SettingType=$r.SettingType; LastModified=$r.LastModified; Expiration=$r.Expiration; Bytes=$r.Bytes }
            if ($withPayload) {
                $o.DecodedPayload = if ($r.Payload) {
                    try { [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($r.Payload)) } catch { '<binary payload>' }
                } else { $null }
            }
            [pscustomobject]$o
        }
        $file = Join-Path $ExportPath "$safe`_windowsbackup_$stamp.json"
        [ordered]@{ exportedBy=$ctx.Account; exportedUtc=(Get-Date).ToUniversalTime().ToString('o'); targetUser=$target; scope=$scopeLabel; includesPayload=$withPayload; settingCount=@($out).Count; settings=$out } |
            ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $file -Encoding UTF8
        if ($withPayload) { Write-Warning "JSON export includes decoded payloads (personal data). Store it securely." }
        Write-Host "Exported ($scopeLabel) to: $file" -ForegroundColor Green
    }
}
else {
    Write-Host "Skipped export. No file was written." -ForegroundColor DarkGray
}

Write-Host "`nDone. Press Enter to close..." -ForegroundColor Cyan
[void](Read-Host)

