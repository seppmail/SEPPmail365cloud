# Changes in the PowerShell Module SEPPmail365cloud

## 0.9.5   "Internal Signature Release"

__Enhancements__

- New-SC365ExoReport:
  - now also includes hidden Microsoft outbound "Testmode" connectors
  - Adds the logged on user as Report-Creator in the header
  - Now has a transparent Logo

- New-SC365Rules now supports -InternalSignature Parameter. This setting adds the two required rules to support this service. ONLY USE WHEN SERVICE is PROVISIONED, otherwise mailflow will fail.

- Added Argument-Completers (automatically select correct values by pressing TAB after a parameter) for
  - New-SC365Connectors -MailDomain
  - New-SC365Rules -ExcludeDomain
  - Get-SC365TenantID -MailDomain

- New-SC365Connectors
  - Now has a "NoInboundEFSkipIPs" switch. If you turn it on, we will not add IPv4 and IPv6 adresses to the EfSkipIps List of the inbound connector

- Test-SC365ConnectionStatus now has a -SessionCleanup parameter to remove old Exchange PS Sessions.

__Maintenance__

- Added numbered prefix to rule files for better identification and sorting order
- Removed region ch as default value on connector creation.
- Removed "seppmail" as default routing mode.

__Fixes__

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

__Fixes__

- Fixed module version issue in Test-SC365Connectors

## Older dev-versions

0.8.2   Add correct code signature

0.8.1   Fix mistakes in connector config, updated visuals and README.md

0.8.0   Initial Release
