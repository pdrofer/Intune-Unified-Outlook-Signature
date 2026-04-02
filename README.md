# 📧 Unified Outlook Signature Deployment (Intune)

## 📖 Overview
Managing email signatures across an enterprise environment is challenging due to the split between Classic Outlook (which relies on local registry keys and files) and the New Outlook / Outlook on the Web (which uses cloud-based named signatures). 

This PowerShell script provides a unified, zero-touch deployment solution designed specifically for **Microsoft Intune**. It dynamically generates and applies signatures for both environments simultaneously, ensuring brand consistency across the organization.

## 🚀 Key Features
* **Dual-Environment Support:** Configures local `.htm`/`.rtf`/`.txt` files for Classic Outlook and utilizes the `ExchangeOnlineManagement` module to set cloud signatures for New Outlook/OWA.
* **Dynamic User Data:** Queries the Microsoft Graph API to pull the logged-in user's Display Name, Job Title, Mobile Phone, and Email Address automatically.
* **Smart Formatting:** Automatically removes the mobile phone prefix row from the HTML template if the user does not have a registered mobile number in Azure AD.
* **DevOps Ready (Zero Hardcoded Secrets):** Built with security in mind. All tenant IDs, application secrets, and certificate passwords are parameterized, preventing credential leaks in version control or endpoint caches.

## 🏗️ Architecture & Security
To avoid the "Secret Zero" problem on endpoint devices, this script requires parameters to be passed at runtime. 
It authenticates using an **Azure App Registration** with the following API permissions:
* `Exchange.ManageAsApp` (Application) - Authenticated via Certificate (`.pfx`)
* `User.Read.All` (Application) - Authenticated via Client Secret

## ⚙️ Prerequisites
1. **Azure App Registration** configured with the permissions listed above.
2. A `.pfx` certificate generated and linked to the App Registration.
3. The `ExchangeOnlineManagement` module folder packaged alongside this script.
4. An HTML template hosted on a publicly accessible URL.

## 📦 Intune Win32 App Deployment
To deploy this via Microsoft Intune, package the `.ps1` script, the `.pfx` certificate, and the Exchange module into an `.intunewin` file using the Microsoft Win32 Content Prep Tool.

**Install Command Example:**
```powershell
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File .\SignatureDeploy.ps1 -TenantID "YOUR-TENANT-ID" -ClientID "YOUR-CLIENT-ID" -ClientSecret "YOUR-SECRET" -CertificateThumbprint "YOUR-CERT-THUMBPRINT" -CertificatePasswordSecure (ConvertTo-SecureString -String 'YourCertPassword' -AsPlainText -Force)