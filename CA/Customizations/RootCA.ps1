param (
    [switch]$Custom
)

Write-Host -ForegroundColor Green "Perform Post Installation Configuration for Root CA"

# Always prompt for DSConfigDN
$DSConfigDN = Read-Host "Please enter the DSConfigDN"
Certutil -setreg CA\DSConfigDN "$DSConfigDN"

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

# Set the periods to Weeks and Years as these don't change often
Certutil -setreg CA\CRLPeriod "Weeks"
Certutil -setreg CA\ValidityPeriod "Years"

Write-Host -ForegroundColor Green "Set auditing"
auditpol /set /category:"object access" /success:enable /failure:enable
Certutil -setreg CA\AuditFilter 127

Write-Host -ForegroundColor Green "Configure the AIA"
certutil -setreg CA\CACertPublicationURLs "1:C:\Windows\system32\CertSrv\CertEnroll\%1_%3%4.crt`n2:ldap:///CN=%7,CN=AIA,CN=Public Key Services,CN=Services,%6%11`n2:http://pki.citrix7202.lab/CertEnroll/%1_%3%4.crt"

Write-Host -ForegroundColor Green "Confirm your settings"
certutil -getreg CA\CACertPublicationURLs
