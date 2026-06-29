<#
.SYNOPSIS
    Grant a sending address permission to send for every Microsoft 365 Group.

.DESCRIPTION
    For the Microsoft 365 SMTP Relay. A Microsoft 365 Group isn't a user
    mailbox, so the relay sends group mail *through* its SendFrom service
    mailbox. For that to work, SendFrom must be allowed to send for the group.
    This script grants that permission across all M365 Groups in one shot.

    Defaults to Send on Behalf (recipients see "SendFrom on behalf of Group").
    Use -AccessRight SendAs for a clean From with no "on behalf of" line.

    The script is idempotent (skips groups already granted) and supports
    -WhatIf to preview changes.

.PARAMETER Trustee
    The address that should be allowed to send for the groups - i.e. the
    relay's SendFrom mailbox.

.PARAMETER AccessRight
    SendOnBehalf (default) or SendAs.

.EXAMPLE
    .\Grant-GroupSendPermissions.ps1 -Trustee noreply@domain.com -WhatIf

.EXAMPLE
    .\Grant-GroupSendPermissions.ps1 -Trustee noreply@domain.com -AccessRight SendAs
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$Trustee,

    [ValidateSet('SendOnBehalf', 'SendAs')]
    [string]$AccessRight = 'SendOnBehalf'
)

$ErrorActionPreference = 'Stop'

# --- ensure the Exchange Online module is loaded and we're connected ---------
if (-not (Get-Command Get-UnifiedGroup -ErrorAction SilentlyContinue)) {
    Import-Module ExchangeOnlineManagement
}
try {
    $null = Get-ConnectionInformation
} catch {
    Write-Host 'Connecting to Exchange Online...' -ForegroundColor Cyan
    Connect-ExchangeOnline -ShowBanner:$false
}

# --- validate the trustee up front ------------------------------------------
$trusteeRecip = Get-Recipient -Identity $Trustee
Write-Host "Trustee: $($trusteeRecip.PrimarySmtpAddress) ($($trusteeRecip.Name))" -ForegroundColor Cyan

# --- enumerate every Microsoft 365 Group ------------------------------------
$groups = @(Get-UnifiedGroup -ResultSize Unlimited)
Write-Host "Found $($groups.Count) Microsoft 365 Group(s). Granting $AccessRight to $Trustee..." -ForegroundColor Cyan

$granted = 0; $skipped = 0; $failed = 0

foreach ($g in $groups) {
    $addr = $g.PrimarySmtpAddress
    $gid  = $g.Guid.ToString()
    try {
        if ($AccessRight -eq 'SendOnBehalf') {
            $current = (Get-UnifiedGroup -Identity $gid).GrantSendOnBehalfTo
            if ($current -contains $trusteeRecip.Name) {
                Write-Host "  [skip] $addr - already granted" -ForegroundColor DarkGray
                $skipped++; continue
            }
            if ($PSCmdlet.ShouldProcess($addr, "Grant SendOnBehalf to $Trustee")) {
                Set-UnifiedGroup -Identity $gid -GrantSendOnBehalfTo @{Add = $Trustee }
                Write-Host "  [ok]   $addr" -ForegroundColor Green
                $granted++
            }
        }
        else {
            $existing = Get-RecipientPermission -Identity $gid -Trustee $Trustee -ErrorAction SilentlyContinue
            if ($existing) {
                Write-Host "  [skip] $addr - already granted" -ForegroundColor DarkGray
                $skipped++; continue
            }
            if ($PSCmdlet.ShouldProcess($addr, "Grant SendAs to $Trustee")) {
                Add-RecipientPermission -Identity $gid -Trustee $Trustee -AccessRights SendAs -Confirm:$false | Out-Null
                Write-Host "  [ok]   $addr" -ForegroundColor Green
                $granted++
            }
        }
    }
    catch {
        Write-Warning "  [fail] $addr - $($_.Exception.Message)"
        $failed++
    }
}

Write-Host "`nDone. Granted: $granted  Skipped: $skipped  Failed: $failed" -ForegroundColor Cyan
Write-Host 'Note: Send permissions can take up to ~1 hour to replicate across Exchange Online.' -ForegroundColor Yellow
