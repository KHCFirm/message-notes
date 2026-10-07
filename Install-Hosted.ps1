# Requires a connected Exchange Online PowerShell session.
# This script changes only Message Notes in the mailbox you name.
[CmdletBinding()]
param(
    [string]$Mailbox,
    [string]$ManifestPath
)
$ErrorActionPreference = 'Stop'
$ExpectedId = '7c6010d9-3b66-45f7-86ca-296e175b2b86'
$RemovedPrevious = $false
try {
    foreach ($CommandName in @('Get-App', 'New-App', 'Enable-App', 'Remove-App')) {
        if (-not (Get-Command $CommandName -ErrorAction SilentlyContinue)) {
            throw 'In this same PowerShell session, run Import-Module ExchangeOnlineManagement and Connect-ExchangeOnline, then run this script again.'
        }
    }
    if ([string]::IsNullOrWhiteSpace($Mailbox)) { $Mailbox = Read-Host 'Mailbox email address to install into' }
    $Mailbox = $Mailbox.Trim()
    if ([string]::IsNullOrWhiteSpace($Mailbox)) { throw 'A mailbox address is required.' }
    if ([string]::IsNullOrWhiteSpace($ManifestPath)) { $ManifestPath = Read-Host 'Full path to the downloaded Message-Notes-GitHub.xml' }
    $ManifestPath = $ManifestPath.Trim().Trim('"')
    if (-not (Test-Path -LiteralPath $ManifestPath -PathType Leaf)) { throw "Manifest not found: $ManifestPath" }
    $ManifestPath = (Resolve-Path -LiteralPath $ManifestPath).ProviderPath
    $Bytes = [System.IO.File]::ReadAllBytes($ManifestPath)
    $Settings = New-Object System.Xml.XmlReaderSettings
    $Settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
    $Settings.XmlResolver = $null
    $Reader = [System.Xml.XmlReader]::Create($ManifestPath, $Settings)
    try {
        $Xml = New-Object System.Xml.XmlDocument
        $Xml.XmlResolver = $null
        $Xml.Load($Reader)
    } finally { $Reader.Dispose() }
    $Ns = New-Object System.Xml.XmlNamespaceManager($Xml.NameTable)
    $Ns.AddNamespace('o', 'http://schemas.microsoft.com/office/appforoffice/1.1')
    $Id = $Xml.SelectSingleNode('/o:OfficeApp/o:Id', $Ns).InnerText
    $Version = $Xml.SelectSingleNode('/o:OfficeApp/o:Version', $Ns).InnerText
    $Source = $Xml.SelectSingleNode('/o:OfficeApp/o:FormSettings/o:Form/o:DesktopSettings/o:SourceLocation', $Ns).GetAttribute('DefaultValue')
    if ($Id -ne $ExpectedId) { throw 'This is not the expected Message Notes manifest. Nothing was changed.' }
    if ($Version -ne '1.0.2.0') { throw 'Download the version 1.0.2.0 manifest from the new hosted setup page.' }
    if ($Xml.OuterXml -match '__BASE_URL__|__ORIGIN__|https?://localhost') {
        throw 'This is a template or a localhost manifest. Use the Download Outlook manifest button on the published website.'
    }
    $SourceUri = [Uri]$Source
    if (-not $SourceUri.IsAbsoluteUri -or $SourceUri.Scheme -ne 'https' -or $SourceUri.IsLoopback -or $SourceUri.UserInfo) {
        throw 'The sidebar URL must point to the published HTTPS website.'
    }
    if ($SourceUri.Host -in @('github.com', 'raw.githubusercontent.com')) {
        throw 'The manifest must use the published GitHub Pages website, not the repository or a raw-file address.'
    }
    Write-Host "`nTarget mailbox: $Mailbox"
    Write-Host "Sidebar address: $Source"
    Write-Host "Version: $Version"
    Write-Host 'Checking the hosted sidebar before changing any registration...'
    $Response = Invoke-WebRequest -UseBasicParsing -Uri $Source -TimeoutSec 30
    if ($Response.StatusCode -ne 200 -or $Response.Content -notmatch 'id="note"') {
        throw 'The sidebar URL did not return the expected application. Check GitHub Pages and upload paths.'
    }
    $Existing = @(Get-App -Mailbox $Mailbox -ErrorAction Stop | Where-Object { "$($_.AppId)".Trim('{}') -eq $ExpectedId })
    if ($Existing.Count -gt 1) { throw 'Multiple matching registrations were returned. Review them before proceeding.' }
    if ($Existing.Count -eq 1) {
        $Existing | Format-List DisplayName, Enabled, AppVersion, AppId
        Write-Host 'This replaces only the Message Notes registration in the mailbox above.' -ForegroundColor Yellow
        Write-Host 'The add-in ID and note format stay unchanged. Preserve copies of any important notes first.'
        Write-Host 'If this app was centrally deployed, stop and update it through the administrator deployment instead.'
        $Consent = Read-Host 'Type REPLACE to remove the old registration and install the hosted manifest'
        if ($Consent -cne 'REPLACE') { Write-Host 'Cancelled. Nothing was changed.'; return }
        Remove-App -Mailbox $Mailbox -Identity $ExpectedId -Confirm:$false -ErrorAction Stop
        $RemovedPrevious = $true
    } else {
        Write-Host 'No matching mailbox registration was found. This does not inventory centrally deployed Integrated apps.'
        $Consent = Read-Host 'Type INSTALL to install Message Notes into the mailbox above'
        if ($Consent -cne 'INSTALL') { Write-Host 'Cancelled. Nothing was changed.'; return }
    }
    New-App -Mailbox $Mailbox -FileData $Bytes -Enabled $true -ErrorAction Stop | Out-Null
    Enable-App -Mailbox $Mailbox -Identity $ExpectedId -ErrorAction Stop
    Write-Host "`nExchange accepted the hosted installation and enable command." -ForegroundColor Green
    $After = @(Get-App -Mailbox $Mailbox -ErrorAction Stop | Where-Object { "$($_.AppId)".Trim('{}') -eq $ExpectedId })
    if ($After.Count -gt 0) { $After | Format-List DisplayName, Enabled, AppVersion, AppId }
    else { Write-Warning 'The follow-up query has not returned the registration. Recheck Get-App before assuming it is visible.' }
    Write-Host 'Refresh Outlook on the web. Open an ordinary received email in this mailbox and choose Apps > Message Notes.'
    Write-Host 'Then restart new Outlook using the same mailbox. The local server is no longer needed for this hosted version.'
    Write-Host 'Exchange accepting a manifest does not by itself verify sidebar loading or note persistence. Run the pilot checks.'
} catch {
    Write-Host "`nHOSTED INSTALLATION ERROR" -ForegroundColor Red
    $_ | Format-List Exception, ErrorDetails, FullyQualifiedErrorId, CategoryInfo -Force
    if ($RemovedPrevious) {
        Write-Warning 'The previous Message Notes registration was removed before this error. Correct the error and rerun this script; do not assume the hosted version is installed. The original local manifest is included separately as manifest.original-local.xml for manual rollback.'
    }
}
