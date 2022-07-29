# Changes in the PowerShell Module SEPPmail365cloud

## 1.0.6 prv Georegion update

- Added ip adresses for prv region

## 1.0.5 Mini-Feature & Bugfix Release

__Enhancements__

- Rules have now 3-digit numbers 100,200, .. to identify their order number
- The report is now automatically opens with the default web-browser

__Bugfixes__

- New-SC365exoReport works with "-Filepath" again
- Autocompleter did not work with PrimaryMailDomain parameter
- Inline Connector creation failed

## 1.0.0 Production Release

This release is based on customer and partner feedback from the first few months of SEPPmail.cloud existence.

__Enhancements__

- Renamed Connectors from the complicated 'MX' name to inbound and outbound.
- RENAMED ROUTING MODES: "seppmail" -> "inline" and "microsoft" -> "parallel".
- The outbound rule will disallow e-mails with Spam Confidence Level (SCL) >=9.
- New-SC365Connectors will create a summary of connector information when finished.

__Maintenance__

- New-SC365Rules
  - Looks also for Testmode-connectors when searching for existing SEPPmail.Cloud-Connectors
  - Removes also [SEPPmail* rules, if client migrates from a selfhosted SEPPmail Appliance
  - Places the SEPPmail transport rules at the bottom by default. This should fit in most cases (i.e. 3rd Party disclaimer solutions)
  - Writes now also module version in Comments

- New-SC365Connectors
  - Changed Outbound ConnectorType in routingmode "parallel" from "OnPremises" to "Partner".
  - Changed parameter -maildomain to -primarymaildomain to better reflect its purpose

- Remove-SC365Connectors
  - No more warnings about missing testmode connectors

__BugFixes__

- Fix Parameter -ExcludeMaildomain in New-SC365Rules.
- Fix - Placementprio Default Parameter "Bottom" had no impact.
- Test-SC365Connection now correctly shows the domain of the newest PSSession

## 0.9.6 Bugfix Release Internal Sinature

__BugFixes__

- Microsoft does not allow "/" in Inbound Connector names in some tenants. Connector-name character replaced: "/" to "-"
- Copy/paste error in IPv6AllowList of prv - fixed

## 0.9.5 "Bugfix release"

__Enhancements__

- New-SC365ExoReport:
  - now also includes hidden Microsoft outbound "Testmode" connectors
  - Adds the logged on user as Report-Creator in the header
  - Now has a transparent Logo

- Added Argument-Completers (automatically select correct values by pressing TAB after a parameter) for
  - New-SC365Connectors -MailDomain
  - New-SC365Rules -ExcludeDomain
  - Get-SC365TenantID -MailDomain

- New-SC365Connectors
  - Now has a "NoInboundEFSkipIPs" switch. If you turn it on, we will not add IPv4 and IPv6 addresses to the EFSkipIps (Enhanced Filtering) list of the inbound connector.

- Test-SC365ConnectionStatus now has a -SessionCleanup parameter to remove old Exchange PS Sessions.

__Maintenance__

- Added numbered prefix to rule files for better identification and sorting order
- Removed region ch as default value on connector creation.
- Removed "seppmail" as default routing mode.

__BugFixes__

- Corrections, typo fixes and better graphics in Readme.MD
- Test-SC365Connection status does not raise an error anymore if only one session is available

## 0.9.1   "Bugfix Release of "German-Cloud" Release

- Change "WhiteList" to "AllowList"

## 0.9.0   "German-Cloud" Release

__Enhancements__

- Update CmdLet-based help for all CmdLets (Get-Help New-SC365ExoReport)
- Updated IP4 and IP6 addresses for german SEPPmail.cloud based on status 17.03.22
- Remove-SC365rules has -routing 'microsoft' as default.
- Get-SC365TenantID is validating the E-mail domain against the current subscription
- Test-SC365ConnectionStatus -showDefaultDomain parameter changed from bool to switch. No $true/$false needed anymore.
- Added (and tested) -WhatIf Support for all "New" and "Remove" CmdLets.
- Added -force switch to New-SC365Connectors. With Force, this CmdLet works without any interaction.
- Added -InboundOnly switch to New-SC365Connectors. Now you can create only Inbound Connectors in -routing 'seppmail' mode.

__BugFixes__

- Fixed module version issue in Test-SC365Connectors

## Older dev-versions

0.8.2   Add correct code signature

0.8.1   Fix mistakes in connector config, updated visuals and README.md

0.8.0   Initial Release
