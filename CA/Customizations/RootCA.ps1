<#
.SYNOPSIS
    Performs post-installation configuration for a Root CA with optional custom parameters.

.DESCRIPTION
    This script configures a Root Certification Authority (CA) on a Windows Server. It sets various 
    configuration parameters such as DSConfigDN, CRLPeriodUnits, CRLDeltaPeriodUnits, CRLOverlapPeriodUnits, 
    and ValidityPeriodUnits. By default, it uses predefined values but can accept custom values for 
    the mentioned parameters when executed with the -Custom switch.

.PARAMETER AiaFqdn
    Specifies a custom FQDN for the AIA configuration. If not specified, the script prompts for it.

.PARAMETER Custom
    Specifies whether the script should prompt for custom values for CRLPeriodUnits, CRLDeltaPeriodUnits, CRLOverlapPeriodUnits, and ValidityPeriodUnits.

.PARAMETER Verbose
    Enables detailed logging of the script's operations.

.PARAMETER LogPath
    Specifies a custom path for the log file. If not specified, uses the %TEMP% directory.

.EXAMPLE
    PS> .\RootCA.ps1
    This execution prompts the user for the DSConfigDN and AIA FQDN, using default values for other parameters.

.EXAMPLE
    PS> .\RootCA.ps1 -Custom -AiaFqdn "pki.example.com" -Verbose

    This example customizes CA configuration with a specified AIA FQDN and enables verbose logging.

.NOTES
    Author: Balint Oberrauch
    Version: 1.0
    Date: 19/03/2024

#>

param (
    [string]$AiaFqdn,
    [switch]$Custom,
    [switch]$Verbose,
    [string]$LogPath
)

# Initialize logging
$LogFilePath = if ($LogPath) { $LogPath } else { Join-Path $env:TEMP "CAConfigLog.txt" }
function LogWrite {
    Param ([string]$logString)
    if ($Verbose) {
        Write-Host $logString
    }
    $logString | Out-File -FilePath $LogFilePath -Append
}

LogWrite "Starting CA configuration script."

# Always prompt for DSConfigDN
$DSConfigDN = Read-Host "Please enter the DSConfigDN"
Certutil -setreg CA\DSConfigDN "$DSConfigDN"
LogWrite "DSConfigDN set to $DSConfigDN"

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
Certutil -setreg CA\CRLPeriodUnits $CRLPeriodUnits
Certutil -setreg CA\CRLDeltaPeriodUnits $CRLDeltaPeriodUnits
Certutil -setreg CA\CRLOverlapPeriodUnits $CRLOverlapPeriodUnits
Certutil -setreg CA\ValidityPeriodUnits $ValidityPeriodUnits

LogWrite "Configuration applied. CRLPeriodUnits: $CRLPeriodUnits, CRLDeltaPeriodUnits: $CRLDeltaPeriodUnits, CRLOverlapPeriodUnits: $CRLOverlapPeriodUnits, ValidityPeriodUnits: $ValidityPeriodUnits"

# Set the periods to Weeks and Years as these don't change often
Certutil -setreg CA\CRLPeriod "Weeks"
Certutil -setreg CA\ValidityPeriod "Years"

Write-Host -ForegroundColor Green "Set auditing"
auditpol /set /category:"object access" /success:enable /failure:enable
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
