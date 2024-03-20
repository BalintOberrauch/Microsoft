<#
.SYNOPSIS
    Performs post-installation configuration for a Root CA with optional custom parameters, including backing up current settings to a specified file.

.DESCRIPTION
    This script configures a Root Certification Authority (CA) on a Windows Server. It sets various configuration parameters such as DSConfigDN, CRLPeriodUnits, CRLDeltaPeriodUnits, CRLOverlapPeriodUnits, and ValidityPeriodUnits. By default, it uses predefined values but can accept custom values for the mentioned parameters when executed with the -Custom switch. Backup functionality is included to save current settings before making changes; the backup path must specify a file, not just a directory.

.PARAMETER AiaFqdn
    Specifies a custom FQDN for the AIA configuration. If not specified, the script prompts for it.

.PARAMETER Custom
    Specifies whether the script should prompt for custom values for CRLPeriodUnits, CRLDeltaPeriodUnits, CRLOverlapPeriodUnits, and ValidityPeriodUnits.

.PARAMETER Verbose
    Enables detailed logging of the script's operations.

.PARAMETER LogPath
    Specifies a custom path for the log file. If not specified, uses the %TEMP% directory. The path must include a file name, not just a directory.

.PARAMETER BackupPath
    Specifies the path to the file where current settings will be backed up before any modifications are made. If not specified, defaults to a file named 'CA_Settings_Backup.txt' in the %TEMP% directory. The path must include a file name, not just a directory.

.EXAMPLE
    PS> .\RootCA.ps1
    This execution prompts the user for the DSConfigDN and AIA FQDN, using default values for other parameters, and backs up settings to the default location.

.EXAMPLE
    PS> .\RootCA.ps1 -Custom -AiaFqdn "pki.example.com" -Verbose -BackupPath "C:\Backups\CA_Backup.txt"
    This example customizes CA configuration with a specified AIA FQDN, enables verbose logging, and backs up current settings to a specified file.

.NOTES
    Author: Balint Oberrauch
    Version: 1.0
    Date: 19/03/2024
#>


param (
    [string]$AiaFqdn,
    [switch]$Custom,
    [switch]$Verbose,
    [string]$LogPath,
    [switch]$BackupSettings = $true, # Added switch with default value of $true
    [string]$BackupPath = "$env:TEMP\CA_Settings_Backup.txt" # Optional: Allows specifying a custom backup path
)


# Check for administrative permissions
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "This script requires administrative permissions. Please run it as an administrator." -ForegroundColor Red
    exit
}
# Initialize logging
function Initialize-LogFilePath {
    param (
        [string]$ProvidedLogPath
    )
    # Check if the provided path is a directory
    if (Test-Path -Path $ProvidedLogPath -PathType Container) {
        # If it's a directory, append a default log filename
        return Join-Path -Path $ProvidedLogPath -ChildPath "RootCAConfigLog.txt"
    } else {
        # Otherwise, assume it's a valid file path
        return $ProvidedLogPath
    }
}


$LogFilePath = if ($LogPath) { Initialize-LogFilePath -ProvidedLogPath $LogPath } else { Join-Path $env:TEMP "RootCAConfigLog.txt" }

function LogWrite {
    param (
        [string]$Message
    )
    # Always log the message
    $Message | Out-File -FilePath $LogFilePath -Append

    # Optionally display it on the console
    if ($Verbose) {
        Write-Host $Message
    }
}


# Backup settings

$settings = @{
    "CRLPeriodUnits" = $CRLPeriodUnits;
    "CRLDeltaPeriodUnits" = $CRLDeltaPeriodUnits;
    "CRLOverlapPeriodUnits" = $CRLOverlapPeriodUnits;
    "ValidityPeriodUnits" = $ValidityPeriodUnits
}

# Initialize Backup

function Backup-CertSettings {
    param (
        [string]$BackupFilePath,
        [hashtable]$Settings
    )
    # Ensure backup directory exists
    $backupDir = Split-Path -Path $BackupFilePath -Parent
    if (-not (Test-Path -Path $backupDir)) {
        New-Item -ItemType Directory -Path $backupDir | Out-Null
    }

    foreach ($setting in $Settings.Keys) {
        $regPath = "CA\$setting"
        $currentValue = & certutil -getreg $regPath 2>&1
        "$regPath=$currentValue" | Out-File -FilePath $BackupFilePath -Append
    }
}

if ($BackupSettings) {
    Backup-CertSettings -BackupFilePath $BackupPath -Settings $settings
}

LogWrite "Starting CA configuration script."

# Always prompt for DSConfigDN
$DSConfigDN = Read-Host "Please enter the DSConfigDN"
LogWrite "DSConfigDN: $DSConfigDN"
Certutil -setreg CA\DSConfigDN "$DSConfigDN"

# Check if AiaFqdn is provided, otherwise prompt
if (-not $AiaFqdn) {
    $AiaFqdn = Read-Host "Please enter the AIA FQDN"
}
LogWrite "AIA FQDN set to $AiaFqdn"

# Construct AIA URL using the provided or prompted FQDN
$AiaUrl = "http://$AiaFqdn/CertEnroll/%1_%3%4.crt"
Certutil -setreg CA\CACertPublicationURLs "2:$AiaUrl"
LogWrite "AIA URL configured as $AiaUrl"

# Default values
$CRLPeriodUnits = 52
$CRLDeltaPeriodUnits = 0
$CRLOverlapPeriodUnits = 12
$ValidityPeriodUnits = 5

# If custom is specified, prompt for custom parameters
if ($Custom) {
    $CRLPeriodUnits = Read-Host "Please enter the CRLPeriodUnits"
    $CRLDeltaPeriodUnits = Read-Host "Please enter the CRLDeltaPeriodUnits"
    $CRLOverlapPeriodUnits = Read-Host "Please enter the CRLOverlapPeriodUnits"
    $ValidityPeriodUnits = Read-Host "Please enter the ValidityPeriodUnits"
}

# Apply the parameters
Certutil -setreg CA\CRLPeriodUnits $CRLPeriodUnits 2>&1 | Tee-Object -FilePath $LogFilePath
Certutil -setreg CA\CRLDeltaPeriodUnits $CRLDeltaPeriodUnits 2>&1 | Tee-Object -FilePath $LogFilePath
Certutil -setreg CA\CRLOverlapPeriodUnits $CRLOverlapPeriodUnits 2>&1 | Tee-Object -FilePath $LogFilePath
Certutil -setreg CA\ValidityPeriodUnits $ValidityPeriodUnits 2>&1 | Tee-Object -FilePath $LogFilePath

LogWrite "Configuration applied. CRLPeriodUnits: $CRLPeriodUnits, CRLDeltaPeriodUnits: $CRLDeltaPeriodUnits, CRLOverlapPeriodUnits: $CRLOverlapPeriodUnits, ValidityPeriodUnits: $ValidityPeriodUnits"

# Set the periods to Weeks and Years as these don't change often
Certutil -setreg CA\CRLPeriod "Weeks"
Certutil -setreg CA\ValidityPeriod "Years"

LogWrite "Set auditing"

$auditOutput = auditpol /set /category:"object access" /success:enable /failure:enable 2>&1
LogWrite $auditOutput

Certutil -setreg CA\AuditFilter 127 

Write-Host -ForegroundColor Green "Configure the AIA"
certutil -setreg CA\CACertPublicationURLs "1:C:\Windows\system32\CertSrv\CertEnroll\%1_%3%4.crt`n2:ldap:///CN=%7,CN=AIA,CN=Public Key Services,CN=Services,%6%11`n2:$AiaUrl"

Write-Host -ForegroundColor Green "Confirm your settings"
certutil -getreg CA\CACertPublicationURLs

Write-Host -ForegroundColor Green "Restarting the services"

Restart-Service certsvc -Verbose

Write-Host -ForegroundColor Green "Publish your CRL"

certutil -crl

LogWrite "Script execution completed."
