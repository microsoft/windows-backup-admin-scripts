# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

# Delete-WindowsBackup.ps1 - deletes ALL Windows backup for a user (whole-user, permanent; no per-device, no undo).
# Runs on Windows PowerShell 5.1 OR PowerShell 7. Requires UserWindowsSettings.ReadWrite.All + User.Read.All
# and the Microsoft 365 Backup Administrator role.
# Run: powershell -ExecutionPolicy Bypass -File .\Delete-WindowsBackup.ps1   (or pwsh)

param(
    [string] $UserId,  # UPN or object id of the user whose backup to delete
    [switch] $Fresh    # force a clean sign-in to switch accounts
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
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { Write-Verbose "TLS 1.2 could not be set: $_" }
    try { Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch { Write-Verbose "NuGet provider bootstrap skipped: $_" }
    # Trust PSGallery only for this install, then restore the previous policy so we don't
    # permanently change the user's machine state.
    $prevGalleryPolicy = $null
    try { $prevGalleryPolicy = (Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue).InstallationPolicy } catch { Write-Verbose "Could not read PSGallery policy: $_" }
    try {
        if ($prevGalleryPolicy -and $prevGalleryPolicy -ne 'Trusted') {
            try { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue } catch { Write-Verbose "Could not set PSGallery policy: $_" }
        }
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber -Confirm:$false -Repository PSGallery
    }
    finally {
        if ($prevGalleryPolicy -and $prevGalleryPolicy -ne 'Trusted') {
            try { Set-PSRepository -Name PSGallery -InstallationPolicy $prevGalleryPolicy -ErrorAction SilentlyContinue } catch { Write-Verbose "Could not restore PSGallery policy: $_" }
        }
    }
}
else {
    Write-Host "[2/4] Microsoft.Graph.Authentication module present." -ForegroundColor Green
}
Import-Module Microsoft.Graph.Authentication

# --- GET helper that follows paging ---
function Get-GraphCollection([string] $Uri) {
    $items = [System.Collections.Generic.List[object]]::new()
    $next = $Uri
    while ($next) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $next -ErrorAction Stop
        if ($resp.value) { foreach ($v in $resp.value) { $items.Add($v) } }
        $next = $resp.'@odata.nextLink'
    }
    return $items
}

# --- Build a deviceId -> display-name map from the backup's own deviceprofile
#     payload (profileId + deviceDisplayName). No extra scope needed. ---
function Get-DeviceNameMap([object[]] $Settings) {
    $map = @{}
    foreach ($s in $Settings) {
        if ($s.payloadType -like '*platform.backuprestore.deviceprofile*') {
            $inst = $s.instances | Select-Object -First 1
            if ($inst.payload) {
                try {
                    $json   = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($inst.payload)) | ConvertFrom-Json
                    $profId = ($json.profileId -replace '[{}]', '')
                    if ($profId -and $json.deviceDisplayName) { $map[$profId] = $json.deviceDisplayName }
                } catch { Write-Verbose "Skipping undecodable payload: $_" }
            }
        }
    }
    return $map
}

# --- Explain a 403 in plain language: it's a role/consent problem, not a bug. ---
function Show-AccessDeniedHelp([string] $Operation, [string[]] $RequiredScopes, [string] $Account) {
    Write-Host ""
    Write-Host "=================== ACCESS DENIED (HTTP 403) ===================" -ForegroundColor Red
    Write-Host (" Signed in as: {0}" -f $Account) -ForegroundColor Yellow
    Write-Host (" This account is NOT authorized to {0}." -f $Operation) -ForegroundColor Yellow
    Write-Host ""
    Write-Host " Cross-user backup delete requires BOTH of the following:" -ForegroundColor Cyan
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

# --- Target ---
if (-not $UserId) {
    $UserId = Read-Host "Enter the user UPN whose backup you want to delete"
}
if ($UserId) { $UserId = $UserId.Trim() }
if ([string]::IsNullOrWhiteSpace($UserId)) {
    Write-Host "No user specified. Aborted." -ForegroundColor Yellow
    return
}

# --- Sign in with the DELETE scope (+ User.Read.All to resolve UPN).
#     -Fresh clears the cached session so you can sign in as a different account.
#     Interactive (browser) first; device code only if the browser sign-in fails. ---
if ($Fresh) {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Write-Host "Cleared cached sign-in. You'll sign in fresh." -ForegroundColor Cyan
}
Write-Host "[3/4] Signing in to Microsoft Graph..." -ForegroundColor Cyan
# No upfront network check - only diagnose connectivity if sign-in actually fails,
# so the normal (working) path stays fast.
try {
    Connect-MgGraph -Scopes 'UserWindowsSettings.ReadWrite.All', 'User.Read.All' -NoWelcome -ErrorAction Stop -WarningAction SilentlyContinue
}
catch {
    try {
        Write-Host "Browser sign-in unavailable - use the device code below to sign in:" -ForegroundColor Cyan
        Connect-MgGraph -Scopes 'UserWindowsSettings.ReadWrite.All', 'User.Read.All' -NoWelcome -UseDeviceAuthentication -ErrorAction Stop
    }
    catch {
        Test-GraphConnectivity | Out-Null
        Write-Host "Sign-in failed. See the guidance above." -ForegroundColor Red
        return
    }
}
$ctx = Get-MgContext
if (-not $ctx -or -not $ctx.Account) {
    Write-Host "Sign-in did not complete. Please try again." -ForegroundColor Red
    return
}
$tenantId = $ctx.TenantId
Write-Host "      Signed in as $($ctx.Account)." -ForegroundColor Green

# --- Resolve UPN -> object id (delete needs {id}@{tenant}) ---
Write-Host "[4/4] Resolving user and reading current backup..." -ForegroundColor Cyan
try {
    if ($UserId -match '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') {
        $objId = $UserId; $target = $UserId
    }
    else {
        $encodedUserId = [System.Uri]::EscapeDataString($UserId)  # Encode UPN for safe insertion into the URI path (e.g. '#')
        $u = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$encodedUserId`?`$select=id,userPrincipalName"
        $objId = $u.id; $target = $u.userPrincipalName
    }
}
catch {
    Write-Host "Could not find user '$UserId' in this tenant: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# --- Blast radius (read what would be deleted) ---
$settings = @()
try { $settings = Get-GraphCollection "https://graph.microsoft.com/beta/users/$objId@$tenantId/settings/windows" }
catch {
    $msg = $_.Exception.Message
    $code = $null; try { $code = [int]$_.Exception.Response.StatusCode } catch { Write-Verbose "No HTTP status code on exception." }
    if ($code -eq 403 -or $msg -match '403' -or $msg -match 'Forbidden') {
        Show-AccessDeniedHelp -Operation ("read/delete the backup for '$target'") -RequiredScopes @('UserWindowsSettings.ReadWrite.All', 'User.Read.All') -Account $ctx.Account
    }
    else { Write-Host "Could not read backup for '$target': $msg" -ForegroundColor Red }
    return
}

if (-not $settings -or @($settings).Count -eq 0) {
    Write-Host "No backup data found for '$target' (nothing to delete)." -ForegroundColor Yellow
    return
}
# Project to objects first (Invoke-MgGraphRequest returns hashtables on PS 5.1,
# which don't support -ExpandProperty / Group-Object on their keys).
$devRows = foreach ($s in $settings) {
    [pscustomobject]@{ Device = $s.windowsDeviceId }
}
$deviceIds = @($devRows | Where-Object { $_.Device } | Select-Object -ExpandProperty Device -Unique)
$devices = $deviceIds.Count
$count   = @($settings).Count

# --- Show which devices/profiles will be removed (with display names) ---
$nameMap = Get-DeviceNameMap $settings
Write-Host "`n-- Device profiles to be deleted --" -ForegroundColor Yellow
$devRows | Group-Object Device | ForEach-Object {
    $devId = $_.Name
    $dName = if ($devId) {
                 if ($nameMap[$devId]) { $nameMap[$devId] } else { '(name unavailable)' }
             } else { '(account-level)' }
    [pscustomobject]@{
        DeviceName   = $dName
        DeviceId     = if ($devId) { $devId } else { '(account-level/default)' }
        SettingCount = $_.Count
    }
} | Format-Table -AutoSize

# --- Single Y/N confirmation, then delete ---
Write-Host ""
$answer = Read-Host ("$target has $devices device profile(s) and $count setting(s). Press Y to delete ALL of it (this cannot be undone)")
if ($answer -notmatch '^(y|yes)$') {
    Write-Host "Aborted. No data was deleted." -ForegroundColor Yellow
    return
}

# --- DELETE (whole-user, irreversible) ---
$uri = "https://graph.microsoft.com/beta/users/$objId@$tenantId/settings/windows"
try {
    Invoke-MgGraphRequest -Method DELETE -Uri $uri -ErrorAction Stop | Out-Null
    Write-Host "DELETED. All backup data for $target has been removed (HTTP 204)." -ForegroundColor Green
}
catch {
    $msg = $_.Exception.Message
    $code = $null; try { $code = [int]$_.Exception.Response.StatusCode } catch { Write-Verbose "No HTTP status code on exception." }
    if ($code -eq 403 -or $msg -match '403' -or $msg -match 'Forbidden') {
        Show-AccessDeniedHelp -Operation ("delete the backup for '$target'") -RequiredScopes @('UserWindowsSettings.ReadWrite.All', 'User.Read.All') -Account $ctx.Account
    }
    else {
        Write-Host "Delete failed (HTTP $code): $msg" -ForegroundColor Red
    }
}

Write-Host "`nDone. Press Enter to close..." -ForegroundColor Cyan
[void](Read-Host)

