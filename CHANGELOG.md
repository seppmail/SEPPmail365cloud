# Changes in the PowerShell Module SEPPmail365cloud

## 1.3.5 - Feature Release

### Enhancements

- Get/Remove-SC365Rules does not ask for routing mode anymore
- Get/Remove-SC365Rules removes all '[SEPPmail.Cloud]*' rules, independent of the config shipped with the module. This allows us to remove rule-Configs and still be able to clean and get the current config.
- Removed Rule 050. The intention was to not send Mails with SPF fails to SEPPmail.Cloud, though there are too many Mailservers out there with misconfigured SPF records, so the rule made more issues than solutions
- New-SC365ExoReport now shows also if a transport rule stops after processing
- New-SC365ExoReport has a new switch parameter -jsonbackup (default off) which stores a complete backup of the tenant config in JSON at the same location as the report is stored. This may be used for a detailed config backup for archiving of diagnosis
- New CmdLet: Get-SC365SetupTime reads the "WhenCreated" property of the SEPPmail.cloud inbound connector and emits the create date on the console.
- 

## 1.3.1 - Maintenance Release

### Enhancements

- Parallel Mode - Inbound Rule 100 not only routes E-Mails to SEPPmail.cloud if they are signed/encrypted somehow (Content-type matches SMIME or PGP sign/encrypt)
- InboundConnector (parallel and inline mode) has EFSkipLastIP set to "true" by default to allow support for ARC-signing
- Enhanced filtering on the inbound connector (parallel and inline) is changed from none to EFSkipLastIP = $true to support ARC-Signing and avoid wrongly detected malware.

### Maintenance

- Updated IP-Address list of DE and CH cloud.
- All transport rules have a new value in property "Auditseverity". Changed from "DoNotAudit" to "Low" for better tracing and debugging of mailflow issues.
- Checking if Tenant allows creation of OnPrem Connectors on Module initialization to see if Exchange Error EX505293 will take place.

## 1.3.0 - Certificate based Connectors Edition

Beginning from Saturday 28. Feb 2023, all SEPPmail Customers can use certificate based connectors (Default for all existing an new customers)
This means that every SEPPmail.cloud customer may now setup with SEPPmail365cloud PS Module version 1.3.0+

### Major changes

- Certificate based connectivity to Exchange Online
- Auto discovery of deployment status
- Get/New/Remove-Setup CmdLets allow one-stop setup for simple environments
- Auto installation of missing modules

### General

- New-Module dependency DNSCLient-PS. Needed for multi-platform DNS-queries
- *.onmicrosoft.com Domains are automatically filtered out if selected (as we do not want to route this traffic through SEPPmail infrastructure
- Detection of TenantID based on current Exo-Session
- Domain-check: When entering a DNSDomain in the Parameter -SEPPmailCloudDomain which is not part of the tenants "AcceptedDomains", the commands New-SC365Rules and New-SC365Connectors will raise an error message, stop, and ask you to enter correct domain(s).
- Big, ugly note at Module startup to read the Readme.md at Github.
- Big, ugly warning signs if wrong PowerShell version or wrong ExchangeOnlineManagement Module on module startup
- Linux(Debian) compatibility. Module has been tested intensively on Debian
- macOS compatibility. Module has been tested intensively on macOS (Intel and Apple 'M' processors)
- M365-Tenants, which are still "hydrated" are now prompted to "Enable-OrganizationCustomization"
- BETA: New Commandlet Get-SC36MessageTrace to trace messages from ExO

### Enhancements

__Common__

- New CmdLet Get-SC365DeploymentStatus checks if deployment is ready and correct on SEPPmail side
- New/Get/Remove-SC365Setup now also work without parameters (leverage GET-SC365DeploymentStatus data)
- Rule [SEPPmail.cloud] - 050 on SPF fail - no SEPPmail.Cloud now stops after processing
- Check if DNS entries are available before deploying connectors
- New/Remove/Get-SC365rules have now a mandatory -routing parameter, as we add transport rules to the inline routing mode.
- New-SC365Setup CmdLet combines all commands to setup an environment
- Remove-SC365Setup CmdLet combines all command to remove a environment setup
- Get-SC365Setup CmdLet shoes the current configuration with one CmdLet.
- Add Confirm-SC365TenantDefaultDomain. A CmdLet to check if a specific domain is the default mail domain of a tenant.

__Connectors__

- Inbound-Connectors are now linking to a specific, TenantID-based certificate, which ensures highest delivery-trust by Microsoft
- Connector type for Inline-Mode is changed to "Partner"
- Slim connector configuration for parallel mode connectors (No SenderIpAdresses, HostedConnectionFilterPolicy, EFSkipIPs)

__Rules__

- Rules now use a positive list of domains. So if a customer adds domains, there is no need to reconfigure the rules until they are booked at SEPPmail.cloud
- New transport "050" rule to avoid failed SPF-check E-mails to be routed to SEPPmail.cloud (parallel mode only)
- New transport rule "600" to remove X-SM-Smarthost header on outgoing mails to force obfuscation of leveraged technology
- Adapted Inbound transport rule to avoid SPAM with SCL Level 5 (parameterized) to be routed to SEPPmail.cloud (parallel mode only)

__Bugfixes__

- Get-SC365Messagetrace now reads encrypted mails with changed messageids correctly
- Get-SC365Messagetrace now reads connector information in inline mode correctly
- [SEPPmail.cloud] - 060 Add header X-SM-ruleversion - now adds the header also inbound

__Maintenance__

- Fix Inbound-Connector Inline mode to SenderDomains "smtp:*;1"
- Bind Inbound-Connector to TLS certificate of Exo-Tenant-default-domain
- Add rule for new X-Header X-SM-ruleversion with version number of PS-Modueversion
- Incoming and outgoing rules have now a positive list of domains instead of an exclusion list

## 1.2.5 Exchange Online adaption and Tenant2Tenant Signature Update

__Maintenance__

- Avoid mail loops between ExO-Tenants in the same region
- Optimized output of Get-SC365rules - Excluded Domains are now seen.
- Domain selection in New-SC365Connectors and New-SC365Rules parameter is now called SEPPmailCloudDomain

## 1.2.0 ExchangeOnlineManagement Module Version 3.0.0 Update

__Enhancements__

- Add Support for ExchangeOnlineManagement 3.0.0 - Support for older versions of the module is disabled.
- New-SC365Connectors stops and raises a warning message if there are still transport rules pointing to the connector
- Test-SC365Connection has now a -Connect parameter to connect within processing (via webbrowser)
- New-SC365Rules now has a mandatory -SEPPmailCloudDomains parameter

__BugFixes__

__Maintenance__

- Removes old routing modes "seppmail" and "microsoft". Now only 'inline' and 'parallel' is allowed.
- Renamed rules.json files to 3 digit numbers to reflect rule names
- Prettify output of rules-commands (name,state,prio,ExceptIfRecipientDomainIs)
- Prettify output of connector-commands (Name,Enabled,WhenCreated,Region)
- New-SC365ExoReport checks if directory is writeable and uses alternatives if not (Windows & Mac)

## 1.1.0 ARC Signing update (22-08-2022)

__Enhancements__

- Added IP-addresses for Preview region
- Added Arc-Sealing with every connector creation to seppmail.cloud

__BugFixes__

- Fixed -ExcludeMaildomain Parameter issue

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
