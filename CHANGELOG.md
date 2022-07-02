# Changes in the PowerShell Module SEPPmail365cloud

0.9.5   "Maintenance release"

Enhancements

- New-SC365ExoReport now also includes hidden outbound "Testmode" Connectors
- New-SC365Rules now supports -InternalSignature Parameter. This setting adds the two required rules to support this service.
- 

Fixes

- Corrections, typo fixes and better graphics in Readme.MD

0.9.1   "Bugfix Release of "German-Cloud" Release

- Change "WhiteList" to "AllowList"

0.9.0   "German-Cloud" Release

Enhancements

- Update CmdLet-based help for all CmdLets (Get-Help New-SC365ExoReport)
- Updated IP4 and IP6 addresses for german SEPPmail.cloud based on status 17.03.22
- Remove-SC365rules has -routing 'microsoft' as default.
- Get-SC365TenantID is validating the E-mail domain against the current subscription
- Test-SC365ConnectionStatus -showDefaultDomain parameter changed from bool to switch. No $true/$false needed anymore.
- Added (and tested) -WhatIf Support for all "New" and "Remove" CmdLets.
- Added -force switch to New-SC365Connectors. With Force, this CmdLet works without any interaction.
- Added -InboundOnly switch to New-SC365Connectors. Now you can create only Inbound Connectors in -routing 'seppmail' mode.

Fixes

- Fixed module version issue in Test-SC365Connectors

0.8.2   Add correct Code signature

0.8.1   Fix mistake in connector config, updated visuals and README.md

0.8.0   Initial Release
