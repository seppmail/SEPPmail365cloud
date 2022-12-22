# Migration Guide from prior M365 Connectivity to Certificate-Based-Connectors with PS Module 1.3.0-preview1 in the SEPPmail.cloud

With early Q1/2023 SEPPmail.cloud will change its connectivity to Exchange online to certificate-based-connectors (CBC). This will increase security and stability between Exchange Online and SEPPmail.cloud and avoid mail-loops in certain situations.

!Important: With CBC, Inline-Mode customers (MX ==> SEPPmail) will have all e-mails routet through SEPPmail.cloud 

## Step 1/5 - Check [prerequisites](https://gitlab.seppmail.ch/internal/seppcloud/seppmail365cloud/-/blob/main/README.md#prerequisites)

To make the PowerShell module work, 4 prerequisites must be met.

1. PowerShell Core 7+ must be used - the Module doesnt work with Windows Powershell 5.1!
2. The ExchangeOnLineManagement Module 3.0.0+ must be installed and loaded (Restart PowerShell session after "Install-Module ExchangeOnlineManagement -Force")
3. The customer/partner needs to know the domains to be migrated, the cloud region and the routing mode for this tenant
4. Check if MX Records are set properly (MX ==> SEPPMail in Inline mode, MX ==> Microsoft in Parallel Mode)

## Step 2/5 - Install the preview release of seppmail365cloud

```powershell
Set-Location ~
Install-Module seppmail365cloud -AllowPrerelease -AllowClobber -Force
Get-Module seppmail365cloud  # This must show the module version 1.3.0-preview1 loaded.
```

## Step 3/5 - Cleanup the environment

Make sure all old end existing SEPPmail rules and connectors are removed. This may be done in the [Exchange Admin GUI](https://admin.microsoft.com/exchange) or with PowerShell CmdLets Remove-SC365Rules and Remove-SC365Commandlets.

Check final results with Get-TransportRule, Get-InboundConnector and Get-OutboundConnector.

**Special Case : Connectors with "/" or "\\" in the name**
We had a version of the seppmailcloud connectors in place which used slashes in the name. Microsoft somehow stopped to accept this. If you find such a connector do this:

1. Rename connectors in the admin.microsoft.com portal
2. Delete them after renaming in the admin portal.

## Step 4/5 - Setup new mailflow to SEPPmail

Follow the guide in the [README](https://gitlab.seppmail.ch/internal/seppcloud/seppmail365cloud/-/blob/main/README.md#setup-the-integration)

## Step 5/5 - Wait up to 10 Minutes

Until all changes are saved in the MS Cloud it sometimes takes a few minutes. Send testmails in that time until te mailflow works and trust the solution.

## Special Cases

- Still mailloops after the changes: If you set up everything according to the description above, and still have maillops, check if the recipient is also in the SEPPmail.cloud. Recipient MUST also use newest connectors (CBC).


Follow instructions from [readme](https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md)