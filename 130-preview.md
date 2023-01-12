# Migration Guide from prior M365 Connectivity to Certificate-Based-Connectors with PS Module 1.3.0-preview[1-n] in the SEPPmail.cloud

With early Q1/2023 SEPPmail.cloud will change its connectivity to Exchange online to certificate-based-connectors (CBC). This will increase security and stability between Exchange Online and SEPPmail.cloud and avoid mail-loops in certain situations.

!Important: Inline customers require all domains of same tenant to be inline domains handled by  SEPPmail.cloud.

## Step 1/5 - Check [prerequisites](https://github.com/seppmail/SEPPmail365cloud#prerequisites)

To make the PowerShell module work, 4 prerequisites must be met.

1. PowerShell Core 7+ must be used - the Module does not work with Windows Powershell 5.1!
2. The ExchangeOnLineManagement Module 3.0.0+ must be installed and loaded (Restart PowerShell session after).

´´´powershell
Install-Module ExchangeOnlineManagement -Force"
´´´

3. The customer/partner needs to know the domains to be migrated, the cloud region and the routing mode for this tenant
4. Check if MX Records are set properly (MX ==> SEPPMail in Inline mode, MX ==> Microsoft in Parallel Mode)

## Step 2/5 - Install the current preview release of seppmail365cloud

```powershell
Set-Location ~
Install-Module seppmail365cloud -AllowPrerelease -AllowClobber -Force
Get-Module seppmail365cloud  # This must show the module version 1.3.0-preview[1...] loaded.
```

## Step 3/5 - Cleanup the environment

Make sure all old end existing SEPPmail rules and connectors are removed. This may be done in the [Exchange Admin GUI](https://admin.microsoft.com/exchange) or with PowerShell CmdLets:

```powershell
# ATTENTION - THIS WILL INFLUENCE THE MAILFLOW - No de/encryption without rules/connectors
Remove-SC365Rules

Remove-SC365Connectors
```

Check final results with:

```powershell
Get-TransportRule
Get-InboundConnector
Get-OutboundConnector
```

No SEPPmail-rule or connector should show up!

**Special Case : Connectors with "/" or "\\" in the name**
We had a version of the SEPPmail.cloud connectors in place which used slashes in the name. Microsoft somehow stopped to accept this. If you find such a connector do this:

1. Rename connectors in the admin.microsoft.com portal
2. Delete them after renaming in the admin portal.

## Step 4/5 - Setup new mailflow to SEPPmail

Follow the guide in the [README](https://github.com/seppmail/SEPPmail365cloud#setup-the-integration)

## Step 5/5 - Wait up to 10 Minutes

Until all changes are saved in the MS Cloud it sometimes takes a few minutes. Send test e-mails in that time until the mailflow works and trust the solution.

## Special Cases

- Still mail loops after the changes: If you set up everything according to the description above, and still have mail-loops, check if the recipient is also in the SEPPmail.cloud. Recipient MUST also use newest connectors (CBC).

Follow instructions from [README](https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md).