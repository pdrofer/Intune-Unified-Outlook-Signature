<#
.SYNOPSIS
    Unified Outlook Signature Deployment Script (New & Classic Outlook)
.DESCRIPTION
    Configures email signatures for the current user across both environments:
    - New Outlook / OWA: Connects via Exchange Online to set a named cloud signature.
    - Classic Outlook: Generates local HTML/RTF/TXT files and configures the registry.
    
    Designed for deployment via Microsoft Intune (Win32 App). Uses an Azure App Registration 
    with Certificate Authentication (Exchange) and Secret Authentication (Graph API).
.NOTES
    Author: [Your Name]
    Version: 3.0 (Portfolio Release - Parameterized and Sanitized)
#>

[CmdletBinding()]
param (
    # --- MANDATORY AUTHENTICATION SECRETS ---
    [Parameter(Mandatory=$true)]
    [string]$TenantID,

    [Parameter(Mandatory=$true)]
    [string]$ClientID,

    [Parameter(Mandatory=$true)]
    [string]$ClientSecret,

    [Parameter(Mandatory=$true)]
    [string]$CertificateThumbprint,

    [Parameter(Mandatory=$true)]
    [securestring]$CertificatePasswordSecure,

    # --- CONFIGURATION (Generic defaults for portability) ---
    [string]$Organization   = "yourtenant.onmicrosoft.com",
    [string]$DomainFallback = "@yourdomain.com",
    [string]$HtmlUrl        = "https://yourdomain.com/assets/signature-template.html",
    [string]$LogoUrl        = "https://yourdomain.com/assets/logo.png",
    [string]$RegCompany     = "YourCompany" # Used for Intune detection rules
)

$ErrorActionPreference = "Stop"

# Decrypt certificate password in memory for local import
$CertificatePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($CertificatePasswordSecure)
)

# ========== STATIC CONFIGURATION ==========
$ClassicSignatureName = "StandardSignature"
$NewSignatureName = "StandardSignature"
$PfxFileName       = "SignatureCert.pfx"
$PfxPath           = Join-Path $PSScriptRoot $PfxFileName
$RegPath           = "HKCU:\Software\$RegCompany\OutlookSignature"
$RegValueName      = "Version"
$RegValueData      = "1.0"

# ========== LOGGING ==========
$logFile = "C:\Temp\OutlookSignatureLog.txt"
if (!(Test-Path "C:\Temp")) { New-Item -Path "C:\Temp" -ItemType Directory -Force | Out-Null }

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File $logFile -Append
}

Write-Log "=== EXECUTION STARTED (Unified Signature Script) ==="

# ----- CLOSE OUTLOOK IF RUNNING -----
try {
    $outlookProc = Get-Process -Name outlook -ErrorAction SilentlyContinue
    if ($outlookProc) {
        Write-Log "Closing running Outlook process..."
        Stop-Process -Name outlook -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 3
    }
} catch {
    Write-Log "Notice: Could not close Outlook: $($_.Exception.Message)"
}

# ----- GET CURRENT USER UPN -----
try {
    $UserUPN = (whoami /upn).Trim()
    if ([string]::IsNullOrWhiteSpace($UserUPN)) {
        $UserUPN = "$env:USERNAME$DomainFallback"
        Write-Log "whoami failed. Using fallback: $UserUPN"
    } else {
        Write-Log "UPN detected: $UserUPN"
    }
} catch {
    $UserUPN = "$env:USERNAME$DomainFallback"
    Write-Log "Error obtaining UPN. Using fallback: $UserUPN"
}

$globalSuccess = $false

# ========== PART 1: NEW OUTLOOK (EXCHANGE ONLINE) ==========
Write-Log "--- Starting New Outlook (Exchange Online) config ---"

try {
    $ModulePath = Join-Path $PSScriptRoot "ExchangeOnlineManagement"
    if (-not (Test-Path $ModulePath)) {
        Write-Log "ExchangeOnlineManagement module not found. Skipping New Outlook part."
        throw "Missing module"
    }

    Import-Module $ModulePath -Force -ErrorAction Stop
    
    if (-not (Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Thumbprint -eq $CertificateThumbprint })) {
        if (-not (Test-Path $PfxPath)) {
            Write-Log "Certificate file not found. Skipping New Outlook part."
            throw "Missing certificate"
        }
        Write-Log "Importing certificate..."
        $pwd = ConvertTo-SecureString -String $CertificatePassword -AsPlainText -Force
        Import-PfxCertificate -FilePath $PfxPath -CertStoreLocation Cert:\CurrentUser\My -Password $pwd -Exportable | Out-Null
    }

    Write-Log "Connecting to Exchange Online..."
    Connect-ExchangeOnline -CertificateThumbprint $CertificateThumbprint -AppId $ClientID -Organization $Organization -DisableWAM -ShowBanner:$false -ErrorAction Stop

    $user = Get-User -Identity $UserUPN -ErrorAction Stop
    $mailbox = Get-Mailbox -Identity $UserUPN -ErrorAction Stop

    $DName = $user.DisplayName
    $Title = if ($user.Title) { $user.Title } else { "Employee" }
    $Mob   = if ($user.MobilePhone) { $user.MobilePhone } else { "" }
    $Email = $mailbox.PrimarySmtpAddress.ToString()

    $htmlContent = (Invoke-WebRequest -Uri $HtmlUrl -UseBasicParsing).Content
    $htmlContent = $htmlContent -replace "\[\[DISPLAYNAME\]\]", $DName
    $htmlContent = $htmlContent -replace "\[\[JOBTITLE\]\]", $Title
    $htmlContent = $htmlContent -replace "\[\[MOBILE\]\]", $Mob
    $htmlContent = $htmlContent -replace "\[\[EMAIL\]\]", $Email
    $htmlContent = $htmlContent -replace "\[\[LOGO\]\]", $LogoUrl

    if ([string]::IsNullOrWhiteSpace($Mob)) {
        $htmlContent = $htmlContent -replace '<td[^>]*>\s*M:\s*\+1\s*</td>', ''
    }

    Write-Log "Applying cloud signature..."
    Set-MailboxMessageConfiguration -Identity $Email -SignatureHtml $htmlContent -SignatureName $NewSignatureName
    Set-MailboxMessageConfiguration -Identity $Email -DefaultSignature $NewSignatureName -DefaultSignatureOnReply $NewSignatureName -AutoAddSignature $true -AutoAddSignatureOnReply $true

    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    $globalSuccess = $true
} catch {
    Write-Log "ERROR in New Outlook block: $($_.Exception.Message)"
}

# ========== PART 2: CLASSIC OUTLOOK (GRAPH API + LOCAL FILES) ==========
Write-Log "--- Starting Classic Outlook config ---"

try {
    $TokenBody = @{
        Grant_Type    = "client_credentials"
        Scope         = "https://graph.microsoft.com/.default"
        Client_Id     = $ClientID
        Client_Secret = $ClientSecret
    }
    $TokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token" -Method POST -Body $TokenBody -ErrorAction Stop
    $AccessToken = $TokenResponse.access_token

    $Headers = @{ "Authorization" = "Bearer $AccessToken"; "Content-Type" = "application/json" }
    $ApiUrl = "https://graph.microsoft.com/v1.0/users?`$filter=userPrincipalName eq '$UserUPN'&`$select=displayName,jobTitle,mail,mobilePhone,businessPhones"
    $ApiResponse = Invoke-RestMethod -Uri $ApiUrl -Headers $Headers -Method Get -ErrorAction Stop

    if ($ApiResponse.value.Count -eq 0) { throw "User not found in Azure AD." }
    $UserData = $ApiResponse.value[0]

    $UserName  = $UserData.displayName
    $UserTitle = $UserData.jobTitle
    $UserEmail = $UserData.mail
    if ($UserData.mobilePhone) { $UserPhone = $UserData.mobilePhone }
    elseif ($UserData.businessPhones) { $UserPhone = $UserData.businessPhones[0] }
    else { $UserPhone = "" }

    if (-not $UserTitle) { $UserTitle = "Employee" }
    if (-not $UserEmail) { $UserEmail = $UserUPN }
    if (-not $UserName)  { $UserName  = $UserUPN }

    $SignatureFolder = "$env:APPDATA\Microsoft\Signatures"
    if (!(Test-Path $SignatureFolder)) { New-Item -ItemType Directory -Path $SignatureFolder -Force | Out-Null }

    $HtmlContent = Invoke-RestMethod -Uri $HtmlUrl -UseBasicParsing
    $NewHtml = $HtmlContent -replace "\[\[DISPLAYNAME\]\]", $UserName
    $NewHtml = $NewHtml -replace "\[\[JOBTITLE\]\]",    $UserTitle
    $NewHtml = $NewHtml -replace "\[\[MOBILE\]\]",      $UserPhone
    $NewHtml = $NewHtml -replace "\[\[EMAIL\]\]",       $UserEmail
    $NewHtml = $NewHtml -replace "\[\[LOGO\]\]",        $LogoUrl

    if ([string]::IsNullOrWhiteSpace($UserPhone)) {
        $NewHtml = $NewHtml -replace '<td[^>]*>\s*M:\s*\+1\s*</td>', ''
    }

    $NewHtml | Out-File -FilePath "$SignatureFolder\$ClassicSignatureName.htm" -Encoding utf8 -Force
    
    $SimpleText = "$UserName - $UserTitle`r`n$UserPhone`r`n$UserEmail"
    $SimpleText | Out-File -FilePath "$SignatureFolder\$ClassicSignatureName.txt" -Encoding utf8 -Force

    $SetupPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Setup"
    if (!(Test-Path $SetupPath)) { New-Item -Path $SetupPath -Force | Out-Null }
    New-ItemProperty -Path $SetupPath -Name "DisableRoamingSignaturesTemporaryToggle" -Value 1 -PropertyType DWORD -Force | Out-Null

    $MailSettingsPath = "HKCU:\Software\Microsoft\Office\16.0\Common\MailSettings"
    if (!(Test-Path $MailSettingsPath)) { New-Item -Path $MailSettingsPath -Force | Out-Null }
    New-ItemProperty -Path $MailSettingsPath -Name "NewSignature" -Value $ClassicSignatureName -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $MailSettingsPath -Name "ReplySignature" -Value $ClassicSignatureName -PropertyType String -Force | Out-Null

    $globalSuccess = $true
} catch {
    Write-Log "ERROR in Classic Outlook block: $($_.Exception.Message)"
}

# ========== INTUNE DETECTION REGISTRY ==========
if ($globalSuccess) {
    try {
        if (!(Test-Path $RegPath)) { New-Item -Path $RegPath -Force | Out-Null }
        New-ItemProperty -Path $RegPath -Name $RegValueName -Value $RegValueData -PropertyType String -Force | Out-Null
    } catch {
        Write-Log "Warning: Could not create detection registry."
    }
}

Write-Log "=== EXECUTION FINISHED ==="
if ($globalSuccess) { exit 0 } else { exit 1 }