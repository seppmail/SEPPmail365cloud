# CHanges in the PowerShell Module SEPPmail365cloud

0.9.0   "German-Cloud" Release

Enhancements
- Update CmdLet-based help for all CmdLets (Get-Help New-SC365ExoReport)
- Updated IP4 and IP6 adresses for german seppmail.cloud based on status 17.03.22
- Get-SC365TenantID is validating the maildomain against the current subscription
- Test-SC365ConnectionStatus -showDefaultDomain parameter changed from bool to switch. No $true/$false neded anymore.
- Added (and tested) -WhatIf Support for all "New" and "Remove" CmdLets.
- Added -force switch to New-SC365Connectors. With Force, this CmdLet works without any interation.
- Added -InbounOnly switch to New-SC365Connectors. Now you can create only Inbound COnnectors in -routing 'seppmail' mode.

Fixes

- Fixed module version issue in Test-SC365Connectors


0.8.2   Add correct Code signature

0.8.1   Fix mistake in connector config, updated visuals and README.md

0.8.0   Initial Release