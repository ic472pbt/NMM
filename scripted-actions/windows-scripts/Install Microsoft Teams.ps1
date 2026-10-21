#description: Installs/Updates MS Teams and WebRTC Service with newest versions. Enables Teams WVD Optimization mode. Recommend to run regularly on Desktop Images.
#execution mode: IndividualWithRestart
#tags: Nerdio, Apps install
<# 
Notes:
This script performs the following:#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Downloads the Microsoft Teams bootstrapper and provisions new Teams
    (including the Outlook meeting add-in) for all users on the machine.

.DESCRIPTION
    1. Downloads teamsbootstrapper.exe from Microsoft (with retries)
    2. Verifies the file is validly signed by Microsoft
    3. Optionally removes the existing provisioned version (-x)
    4. Provisions the latest Teams (-p), plus the Teams Meeting Add-in (--installTMA)
    5. Cleans up the installer and writes a transcript log

    Ref: https://learn.microsoft.com/en-us/microsoftteams/new-teams-vdi-requirements-deploy#new-teams-and-outlook-integration

.NOTES
    Written to run unmodified under Nerdio Manager / Azure Custom Script Extension.
    There is deliberately NO param() block: Nerdio prepends its own variable assignments
    to the script, and param() must be the first statement or PowerShell throws parse errors.
    Edit the CONFIG section below instead.
#>

#region CONFIG - edit these values
$DownloadUrl      = 'https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409'
$WorkingDirectory = if ($PSScriptRoot) { $PSScriptRoot } else { $env:TEMP }
$LogPath          = Join-Path $env:TEMP 'Install-NewTeams.log'
$MaxRetries       = 3
$RemoveExisting   = $false   # $true = run "teamsbootstrapper.exe -x" before provisioning
$SkipTMA          = $false   # $true = do not pass --installTMA (skips Outlook meeting add-in)
$KeepInstaller    = $false   # $true = leave teamsbootstrapper.exe on disk after the run
#endregion CONFIG

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'   # Invoke-WebRequest is dramatically faster in PS 5.1 without the progress bar

# PS 5.1 defaults to older TLS versions; Microsoft endpoints require TLS 1.2+
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

$InstallerPath = Join-Path $WorkingDirectory 'teamsbootstrapper.exe'
$ReturnCode = 0

#region Helpers
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )
    $color = switch ($Level) {
        'INFO'  { 'Cyan' }
        'OK'    { 'Green' }
        'WARN'  { 'Yellow' }
        'ERROR' { 'Red' }
    }
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-Host "[$stamp] [$Level] $Message" -ForegroundColor $color
}

function Get-Bootstrapper {
    param([string]$Uri, [string]$Destination, [int]$Retries)

    for ($attempt = 1; $attempt -le $Retries; $attempt++) {
        try {
            Write-Log "Downloading bootstrapper (attempt $attempt of $Retries)..."
            Invoke-WebRequest -Uri $Uri -OutFile $Destination -UseBasicParsing -TimeoutSec 120

            if (-not (Test-Path -LiteralPath $Destination) -or (Get-Item -LiteralPath $Destination).Length -eq 0) {
                throw 'Download completed but the file is missing or empty.'
            }

            $sizeMB = [math]::Round((Get-Item -LiteralPath $Destination).Length / 1MB, 2)
            Write-Log "Download complete ($sizeMB MB): $Destination" 'OK'
            return
        }
        catch {
            Write-Log "Download failed: $($_.Exception.Message)" 'WARN'
            if (Test-Path -LiteralPath $Destination) {
                Remove-Item -LiteralPath $Destination -Force -ErrorAction SilentlyContinue
            }
            if ($attempt -eq $Retries) { throw "Unable to download bootstrapper after $Retries attempts." }

            $delay = 5 * $attempt
            Write-Log "Retrying in $delay seconds..." 'WARN'
            Start-Sleep -Seconds $delay
        }
    }
}

function Test-MicrosoftSignature {
    param([string]$Path)

    $sig = Get-AuthenticodeSignature -FilePath $Path
    if ($sig.Status -ne 'Valid') {
        throw "Signature check failed. Status: $($sig.Status). The file may be corrupt or tampered with."
    }
    if ($sig.SignerCertificate.Subject -notmatch 'O=Microsoft Corporation') {
        throw "Unexpected signer: $($sig.SignerCertificate.Subject)"
    }
    Write-Log "Signature valid. Signer: $($sig.SignerCertificate.Subject)" 'OK'
}

function Invoke-Bootstrapper {
    param(
        [string]$Path,
        [string[]]$Arguments,
        [string]$Description
    )

    Write-Log "$Description  >  teamsbootstrapper.exe $($Arguments -join ' ')"
    $proc = Start-Process -FilePath $Path -ArgumentList $Arguments -NoNewWindow -Wait -PassThru

    if ($proc.ExitCode -ne 0) {
        throw "$Description failed with return code $($proc.ExitCode)."
    }
    Write-Log "$Description completed (return code 0)." 'OK'
}

function Set-WvdRegistryKey {
    $keyPath = 'HKLM:\SOFTWARE\Microsoft\Teams'
    Write-Log 'Configuring Teams registry key for WVD...'
    try {
        New-Item -Path $keyPath -Force | Out-Null
        New-ItemProperty -Path $keyPath -Name 'IsWVDEnvironment' -PropertyType DWord -Value 1 -Force | Out-Null
        Write-Log 'Registry key set: HKLM\SOFTWARE\Microsoft\Teams\IsWVDEnvironment = 1' 'OK'
    }
    catch {
        # Non-fatal: continue with the install, but make the failure visible
        Write-Log "Failed to set Teams registry key. Error: $($_.Exception.Message)" 'WARN'
    }
}
#endregion Helpers

try {
    Start-Transcript -Path $LogPath -Append | Out-Null
    Write-Log "Starting new Teams provisioning on $env:COMPUTERNAME"
    Write-Log "Log file: $LogPath"

    if (-not (Test-Path -LiteralPath $WorkingDirectory)) {
        New-Item -ItemType Directory -Path $WorkingDirectory -Force | Out-Null
    }

    # 1. Download + verify
    Get-Bootstrapper -Uri $DownloadUrl -Destination $InstallerPath -Retries $MaxRetries
    Test-MicrosoftSignature -Path $InstallerPath

    # 2. Set WVD environment flag (before provisioning so Teams picks it up)
    Set-WvdRegistryKey

    # 3. Optional removal of previous version
    if ($RemoveExisting) {
        Invoke-Bootstrapper -Path $InstallerPath -Arguments @('-x') -Description 'Removing existing Teams'
    }

    # 4. Install latest (July 2025+: Outlook meeting add-in is installed via --installTMA)
    $installArgs = @('-p')
    if (-not $SkipTMA) { $installArgs += '--installTMA' }
    Invoke-Bootstrapper -Path $InstallerPath -Arguments $installArgs -Description 'Provisioning new Teams'

    Write-Log 'All steps completed successfully.' 'OK'
}
catch {
    $ReturnCode = 1
    Write-Log "FAILED: $($_.Exception.Message)" 'ERROR'
    Write-Log "At: $($_.InvocationInfo.ScriptName):$($_.InvocationInfo.ScriptLineNumber)" 'ERROR'
}
finally {
    if (-not $KeepInstaller -and (Test-Path -LiteralPath $InstallerPath)) {
        Remove-Item -LiteralPath $InstallerPath -Force -ErrorAction SilentlyContinue
        Write-Log 'Removed downloaded installer.'
    }
    try { Stop-Transcript | Out-Null } catch { }
}

# Return code: 0 = success, 1 = failure (see log above / transcript).
# Printed explicitly so it is visible in Nerdio's "Script output", then returned to the caller.
$resultText = if ($ReturnCode -eq 0) { 'success' } else { 'failure - review the log above' }
Write-Log "RETURN CODE: $ReturnCode ($resultText)" $(if ($ReturnCode -eq 0) { 'OK' } else { 'ERROR' })
return $ReturnCode

1. Sets registry value for MS Teams to WVD Mode
2. Uninstall MSTeams and WebRTC program
3. Downloads and Installs latest version of MS Teams machine-wide (Not per-user)
4. Downloads and Installs latest version of WebRTC component
5. Sends logs to C:\Windows\temp\NerdioManagerLogs\ScriptedActions\msteams
#>

# Start powershell logging
function Write-Log {
    param
    (
        [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
        [string]$Message
    )
    try {
        Write-Host $Message
        $tempFolder = [environment]::GetEnvironmentVariable('TEMP', 'Machine')
        $logsFolderName = "NerdioManagerLogs\ScriptedActions\msteams"
        $logsPath = "$tempFolder\$logsFolderName"
  
        if (-not (Test-Path -Path $logsPath)) {
            New-Item -Path $tempFolder -Name $logsFolderName -ItemType Directory -Force | Out-Null
        }
  
        $DateTime = Get-Date -Format "MM-dd-yy HH:mm:ss"
        if ($Message) {
            Add-Content -Value "$DateTime - $Message" -Path "$logsPath\ps_log.txt"
        }
    }
    catch {
        Write-Error $_.Exception.Message
    }
}

$SaveVerbosePreference = $VerbosePreference
$VerbosePreference = 'continue'
$VMTime = Get-Date
$LogTime = $VMTime.ToUniversalTime()
Write-Log "################# New Script Run #################"
Write-Log "Current time (UTC-0): $LogTime"

# set registry values for Teams to use VDI optimization 
Write-Log "INFO: Adjusting registry to set Teams to WVD Environment mode"
reg add HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Teams /v "IsWVDEnvironment" /t REG_DWORD /d 1 /f

# uninstall any previous versions of MS Teams or Web RTC
# Per-user teams uninstall logic 
$TeamsPath = [System.IO.Path]::Combine($env:LOCALAPPDATA, 'Microsoft', 'Teams')
$TeamsUpdateExePath = [System.IO.Path]::Combine($env:LOCALAPPDATA, 'Microsoft', 'Teams', 'Update.exe')
try {
    if ([System.IO.File]::Exists($TeamsUpdateExePath)) {
        Write-Log "INFO: Uninstalling Teams process (per-user installation)"

        # Uninstall app
        $proc = Start-Process $TeamsUpdateExePath "-uninstall -s" -PassThru
        $proc.WaitForExit()
    }
    else {
        Write-Log "INFO: No per-user teams install found."
    }
    Write-Log "INFO: Deleting any possible Teams directories (per user installation)."
    Remove-Item -path $TeamsPath -recurse -ErrorAction SilentlyContinue
}
catch  {
    Write-Output "Uninstall failed with exception $_.exception.message"
}

# Per-Machine teams uninstall logic
$GetTeams = get-wmiobject Win32_Product | Where-Object IdentifyingNumber -match "{731F6BAA-A986-45A4-8936-7C3AAAAA760B}"
if ($null -ne $GetTeams){
    Start-Process C:\Windows\System32\msiexec.exe -ArgumentList '/x "{731F6BAA-A986-45A4-8936-7C3AAAAA760B}" /qn /norestart' -Wait 2>&1
    Write-Log "INFO: Teams per-machine Install Found, uninstalling teams"
}

# WebRTC uninstall logic
$GetWebRTC = get-wmiobject Win32_Product | Where-Object IdentifyingNumber -match "{FB41EDB3-4138-4240-AC09-B5A184E8F8E4}"
if ($null -ne $GetWebRTC){
    Start-Process C:\Windows\System32\msiexec.exe -ArgumentList '/x "{FB41EDB3-4138-4240-AC09-B5A184E8F8E4}" /qn /norestart' -Wait 2>&1
    Write-Log "INFO: WebRTC Install Found, uninstalling Current version of WebRTC"
}

# make directories to hold new install 
mkdir "C:\Windows\Temp\msteams_sa\install" -Force

# grab MSI installer for MSTeams
$DLink = "https://teams.microsoft.com/downloads/desktopurl?env=production&plat=windows&arch=x64&managedInstaller=true&download=true"
Invoke-WebRequest -Uri $DLink -OutFile "C:\Windows\Temp\msteams_sa\install\Teams_windows_x64.msi" -UseBasicParsing

# use installer to install Machine-Wide
Write-Log "INFO: Installing MS Teams"
Start-Process C:\Windows\System32\msiexec.exe `
-ArgumentList  '/i C:\Windows\Temp\msteams_sa\install\Teams_windows_x64.msi /l*v C:\Windows\temp\NerdioManagerLogs\ScriptedActions\msteams\teams_install_log.txt ALLUSER=1 ALLUSERS=1 /qn /norestart' -Wait 2>&1

# use MS shortcut to WebRTC install
$dlink2 = "https://aka.ms/msrdcwebrtcsvc/msi"

# grab MSI installer for WebRTC
Invoke-WebRequest -Uri $DLink2 -OutFile "C:\Windows\Temp\msteams_sa\install\MsRdcWebRTCSvc_x64.msi" -UseBasicParsing

# install Teams WebRTC Websocket Service
Write-Log "INFO: Installing WebRTC component"
Start-Process C:\Windows\System32\msiexec.exe `
-ArgumentList '/i C:\Windows\Temp\msteams_sa\install\MsRdcWebRTCSvc_x64.msi /l*v C:\Windows\temp\NerdioManagerLogs\ScriptedActions\msteams\WebRTC_install_log.txt /qn /norestart' -Wait 2>&1
Write-Log "INFO: Finished running installers. Check C:\Windows\Temp\msteams_sa for logs on the MSI installations."
Write-Log "INFO: All Commands Executed; script is now finished. Allow 5 minutes for teams to appear"

$VerbosePreference=$SaveVerbosePreference
