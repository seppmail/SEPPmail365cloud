![](./Visuals/SMC-rm.png)

***Setup SEPPmail.cloud in 3 easy steps:***

| 1-Install PowerShell CORE | 2-Install module | [3-Run the Setup Command](#setup-the-integration-for-basic-environments) |
| --------------- | --------------- | --------------- |
| [Microsoft PS-Core install docs](https://learn.microsoft.com/de-de/powershell/scripting/install/installing-powershell?view=powershell-7.3) | ```Install-Module seppmail365cloud``` | ```New-SC365Setup -force``` |

## NEW customers: Basic setup for most (single maildomain) environments

If you are a new customer and just got the deployment e-Mail from SEPPmail.cloud all you have to do is to run:

```powershell
New-SC365Setup
```

This will setup all necessary connectors and rules for your Microsoft tenant.

## EXISTING customers: Basic setup-update for most (single maildomain) environments.

If you are an existing customer and just and just updated the PowerShell Module, do is to to update:

```powershell
New-SC365Setup -force
```

This will recreate all necessary connectors and rules for your Microsoft tenant.



## Check you Setup

If you want to know what setup has been implemented in your tenant, run Get-SC365Setup

```powershell
Get-SC365Setup
```

This will list Connectors and Transport Rules of your tenant.

## Remove SEPPmail.cloud integration

To remove all SEPPmail.cloud connectors and rules, run:

```powershell
Remove-SC365Setup
```

- [The SEPPmail365cloud PowerShell Module README.MD](#the-seppmail365cloud-powershell-module-readmemd)
  - [Introduction](#introduction)
  - [Latest Changes](#latest-changes)
  - [Prerequisites](#prerequisites)
  - [Operating Systems](#operating-systems)
  - [Security](#security)
  - [Module Installation](#module-installation)
    - [Installation of the Module](#installation-of-the-module)
    - [Additional steps on macOS and Linux (experimental)](#additional-steps-on-macos-and-linux-experimental)
  - [Things to know before you start changing your E-Mail-routing](#things-to-know-before-you-start-changing-your-e-mail-routing)
    - [Routing modes](#routing-modes)
      - [Routing mode "inline"](#routing-mode-inline)
      - [Routing mode "parallel"](#routing-mode-parallel)
    - [Cloud Regions](#cloud-regions)
  - [Using the seppmail365cloud PowerShell module](#using-the-seppmail365cloud-powershell-module)
    - [Get to know your Microsoft Exchange Online environment](#get-to-know-your-microsoft-exchange-online-environment)
    - [Clean up before installing](#clean-up-before-installing)
  - [Setup the integration for BASIC environments](#setup-the-integration-for-basic-environments)
  - [Setup the integration for ADVANCED environments](#setup-the-integration-for-advanced-environments)
    - [Example for routingmode: inline](#example-for-routingmode-inline)
    - [Example for routingmode: inline/inboundonly](#example-for-routingmode-inlineinboundonly)
    - [Example for routingmode: parallel](#example-for-routingmode-parallel)
  - [Review the changes](#review-the-changes)
  - [Test your mailflow](#test-your-mailflow)
  - [Advanced Setup Options](#advanced-setup-options)
    - [Creating Connectors and disabled Rules for time-controlled integration](#creating-connectors-and-disabled-rules-for-time-controlled-integration)
    - [Place TransportRules at the top of the rule-list](#place-transportrules-at-the-top-of-the-rule-list)
    - [Use AllowLists for specific customer situations](#use-allowlists-for-specific-customer-situations)
      - [Allowlisting SEPPmail.cloud in the Defender Enhanced Filtering List (Parallel Mode)](#allowlisting-seppmailcloud-in-the-defender-enhanced-filtering-list-parallel-mode)
      - [Allowlisting SEPPmail.cloud in the Anti-SPAM filter list, aka HostedConnectionFilterPolicy (Parallel Mode)](#allowlisting-seppmailcloud-in-the-anti-spam-filter-list-aka-hostedconnectionfilterpolicy-parallel-mode)
  - [Issues and solutions](#issues-and-solutions)
    - [Computer has User home directory on a fileshare (execution policy error)](#computer-has-user-home-directory-on-a-fileshare-execution-policy-error)
    - [Special Case : Connectors with "/" or "\\" in the name](#special-case--connectors-with--or--in-the-name)
    - [Special Cases - Still mail-loops after re-setup with Version 1.3.0+](#special-cases---still-mail-loops-after-re-setup-with-version-130)
    - [Well-Known Error: New-SC365Rukes asks for rulenames](#well-known-error-new-sc365rukes-asks-for-rulenames)

# The SEPPmail365cloud PowerShell Module README.MD

## Introduction

The SEPPmail365cloud PowerShell Core module has been built to integrate Exchange Online instances into the SEPPmail.cloud (SMC).
The module requires you to connect to your Exchange Online environment as administrator (or at least with Exchange Online administrative rights) via PowerShell Core and creates all necessary connectors and rules, based on the e-mail-routing type and region for you with a few commands.

## Latest Changes

Changes in the module versions are documented in the ![CHANGELOG.md](./CHANGELOG.md)

## Prerequisites

>Note: *Windows PowerShell (5.1 and earlier versions) is not supported!* To run the module from Windows, install __PowerShell Core__ on your Windows machine, using the Microsoft Store or go to [Github](https://github.com/powershell/powershell) for other installation options.

The module requires:

- *PowerShell Core* (minimum version 7.2.1)

The module requires and automatically installs:

- Exchange Online Module version minimum 3.0.0+
- DNSClient-PS 1.0+

If you want to know how to connect to Exchange Online via Powershell read [https://learn.microsoft.com](https://learn.microsoft.com/de-de/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps).

## Operating Systems

The code and was tested on **Windows** and **macOS**. PowerShell Core on _Linux_ should work as well, but has not been intensively tested so far.

## Security

When connecting to Exchange Online, we recommend using the **-Device** or **-Credential** based login option. If you want to use credential-based login, we recommend using the Microsoft **Secrets Management** module to store your username/passwords in a secure place on your disk.

## Module Installation

>IMPORTANT! **Do not use the PowerShell Module SEPPmail365 (without _cloud_ at the end)** we have also on the PowerShell Gallery. This module will create NON-WORKING setups as it is intended to be used with self-hosted, customer owned or MSP-operated SEPPmail **Appliances**.

### Installation of the Module

To install the SEPPmail365Cloud module, open Powershell Core (pwsh.exe) and execute:

```powershell
cd ~                              # moves into the Home-Directory

Install-Module "SEPPmail365cloud" # installs the SEPPmail365cloud module
```

### Additional steps on macOS and Linux (experimental)

In addition to the main module you need to add PSWSMan which adds WSMan client libraries to Linux and macOS for remote connectivity to Exchange Online.

>Note: *Do this OUTSIDE Powershell in the appropriate shell (bash, or similar)!*

```bash
sudo pwsh -command 'Install-Module PSWSMan' #Read more on this here https://github.com/jborean93/omi
sudo pwsh -Command 'Install-WSMan'
```

__Further information__ and detailed steps for the module setup can be found on our GitHub repository for out other PowerShell Module [SEPPmail365 module documentation](https://github.com/seppmail/SEPPmail365#module-installation).

## Things to know before you start changing your E-Mail-routing

### Routing modes

When integrating your Exchange online environment with seppmail.cloud, you have to decide between two e-mail routing modes to Microsoft. We either set the mx-record to *SEPPmail.cloud* (inline) or leave it at *Microsoft* (parallel). Customers routing e-Mails via seppmail.cloud benefit from our outstanding e-mail filter which prevents spam and unwanted software flowing into your network via e-mail.

>Note: If you leave the mx-record at microsoft you CANNOT use the SEPPmail.cloud e-mail filter, but for sure our encryption processing possibilities.

Now lets look into the 2 different modes.

#### Routing mode "inline"

Routing mode "inline" allows you to use the full power of the SEPPmail.cloud! In this scenario, the __mx-record of the e-mail domain is set to the SEPPmail cloud hosts__. Inbound e-mails flow to the SEPPmail.cloud, are scanned, treated cryptographically and then flow to Microsoft via connectors. Same is outbound, the mails simply pass the SEPPmail.cloud before leaving to the internet.

![inline](./Visuals/seppmail365cloud-visuals-inline.png)

#### Routing mode "parallel"

This routing mode is similar to the way you would integrate any SEPPmail Appliance (self hosted or MSP) with ExchangeOnline. E-mails flow to Microsoft, and are looped through SEPPmail.cloud, based on the need for cryptographic treatment. Unfortunately, no SEPPmail Virus or SPAM filter is possible in this configuration.

![parallel](./Visuals/seppmail365cloud-visuals-parallel.png)

### Cloud Regions

SEPPmail.cloud is operated in different cloud-regions (datacenters). based on what you ordered, your tenant may be provisioned in one or the other cloud-region.

Deploying to the wrong region will lead to a non-working environment. Your onboarding E-Mail should contain all information for your region.

## Using the seppmail365cloud PowerShell module

### Get to know your Microsoft Exchange Online environment

After the module setup is completed as described above and you have connected to your Exchange Online environment, create an environment report.

```powershell
# The easiest way is to run the command without any parameter
New-SC365ExOReport # Will generate a report with an autogenerated name in the current folder. Search for *.HTML files.

# Also simpler with automatic creation of filename with timestamp
New-SC365ExOReport -FilePath ~/Desktop

# Of define the filename manually
New-SC365ExoReport -FilePath /Users/you/Desktop/MyExoReport.html

# If you want to specify a literal path use:
New-SC365ExOReport -Literalpath c:\temp\myexoreport.html
```

The report will give you valued information about existing connectors, rules and other mailflow-related information. Keep this report stored for later investigation by support or as a documentation of a certain state.

### Clean up before installing

If your Exchange Online environment was originally integrated with a SEPPmail Appliance, you need to **remove the existing SEPPmail365 connectors and rules** before integrating into SEPPmail.cloud.
To do this use our OTHER PS-Module **SEPPmail365** or the Office admin Portal. Find info on [Remove SEPPmail connectors and rules here.](https://github.com/seppmail/SEPPmail365#cleanup-environment)

>Note: *If you do not remove existing __[SEPPmail]__ rules and connectors, the mailflow will be a **mess** and the **integration will not work**.*


## Setup the integration for BASIC environments

A basic environment has the following characteristics.

- All customer mailboxes are hosted in Exchange Online.
- The customer uses **one** e-mail domain for all users.
- This one e-mail domain is the tenant-default domain.
- This e-mail has been used to book SEPPmail.cloud.
- There are no hybrid connectors.
- There are no other external connectors.
- There are no cross-tenant connectors.
- There are no other 3rd party connectors.
- There are no transport-rules implemented which may affect mailrouting to SEPPmail.cloud

If all those requirements are met the 3 commandlets below are your friends and will do the setup job for you.

- Get-SC365Setup ==> Read the existing setup
- New-SC365Setup ==> create a new setup
- Remove-SC365Setup ==> remove an existing setup

## Setup the integration for ADVANCED environments

Advanced Setups require a deeper understanding of the impact of the SEPPmail.cloud integration, and allow more flexibility.

After you have received a **welcome e-mail** from SEPPmail, and followed all instructions in the e-mail, you can start with the integration.

You need to know 3 input values to run the CmdLets.

- **SEPPMailCloudDomain** (the e-mail domain of your Exchange Online environnement that has been configured in the seppmail.cloud. Most of the time this is the default-domain in your Exchange Online Tenant.)
- **routing** (either "inline" or "parallel", read above for details)
- **region** ("de" or "ch", the geographical region of the SEPPmail.cloud infrastructure)
- **inBoundOnly** (a parameter you may set or not set in INLINE Mode only, which is for customers which use our INBOUND filter only)

>Note: All 4 parameters are automatically populating valid options if you press TAB after the parameter. This reduces typo errors.

You need to setup inbound and outbound-connectors and transport rules, so run the two commands as explained below.

### Example for routingmode: inline

```powershell
New-SC365Connectors -SEPPmailCloudDomain 'contoso.ch' -routing 'inline' -region 'ch'

New-SC365Rules -routing 'inline' -SEPPmailCloudDomain 'contoso.eu'

```
### Example for routingmode: inline/inboundonly

```powershell
New-SC365Connectors -SEPPmailCloudDomain 'contoso.ch' -routing 'inline' -region 'ch' -inboundonly

# No rules required for inbound only setups
```

### Example for routingmode: parallel

```powershell
New-SC365Connectors -SEPPmailCloudDomain 'contoso.eu' -routing 'parallel' -region 'de'

# Important: Rules can only be created if the connectors are enabled. They are enabled by default, so if you use the example above it will work.
New-SC365Rules -routing parallel -SEPPmailCloudDomain 'contoso.eu'
```

## Review the changes

```powershell
Get-SC365Connectors -routing parallel
Get-SC365Rules -routing parallel

# Important: Those 2 commands will show the current connectors and rules.
```

You can use also the native Exchange Online Commandlets.

```Get-Inboundconnector``` and ```Get-OutboundConnector``` will show the installed connectors, and ```Get-Transportrule``` CmdLet will give you all information about transport rules.

## Test your mailflow

Send an e-mail from inside-out and outside-in to see if the mailflow is working.

## Advanced Setup Options

The module allows some extra-tweaks for advanced configurations

### Creating Connectors and disabled Rules for time-controlled integration

For sensitive environments, where mailflow may only be changed in specific time frames, it is possible to create rules and connectors "disabled". Both CmdLets New-SC365Connectors and New-SC365Rules have a -disabled switch. See examples below:

```powershell
New-SC365Connectors -SEPPmailCloudDomain 'contoso.eu' -routing 'parallel' -region 'de'
New-SC365Rules -disabled
```

To enable the disabled transport rules in ExchangeOnline so that e-mails can flow through the SEPPmail.cloud, use the Exchange Online admin-website or the PowerShell Commands Enable/Disable-TransportRule.

```powershell

# Enable SEPPmail.cloud TransportRules
Get-TransportRule -Identity '[SEPPmail.cloud]*'|Enable-TransportRule

# Disable SEPPmail.cloud TransportRules
Get-TransportRule -Identity '[SEPPmail.cloud]*'|Disable-TransportRule

# To avoid getting asked for every rule to confirm the change and run the command in "silent" mode use
Get-TransportRule -Identity '[SEPPmail.cloud]*'|Disable-TransportRule -Confirm:$false -Verbose


```

### Place TransportRules at the top of the rule-list

By default out transport rules will be placed at the bottom of all other transport rules. If you want to change this use:

```powershell
New-SC365Rules -PlacementPriority Top
```

### Use AllowLists for specific customer situations

#### Allowlisting SEPPmail.cloud in the Defender Enhanced Filtering List (Parallel Mode)

We saw situations where the EnhancedFiltering Allowlist needed to be filled with SEPPmail.cloud IP addresses.
Use the example below to deploy those IP addresses properly.

```powershell
New-SC365Connectors -SEPPmailCloudDomain 'contoso.eu' -routing 'inline' -region 'de' -inboundEFSkipIPs
```

#### Allowlisting SEPPmail.cloud in the Anti-SPAM filter list, aka HostedConnectionFilterPolicy (Parallel Mode)

We saw situations where the Hosted Connection Filter Policy needed to be filled with SEPPmail.cloud IP addresses.
Use the example below to deploy those IP addresses properly.

```powershell
New-SC365Connectors -SEPPmailCloudDomain 'contoso.eu' -routing 'inline' -region 'de' -option AntiSpamAllowListing
```

## Issues and solutions

### Computer has User home directory on a fileshare (execution policy error)

If your computer has the users directory on a fileshare, Powershell still installs the Module in the $currentuser scope in your homedirectory. This sill raise issues with execution policy settings. To avoid this you can.

1.) Start PowerShell with no execution policy, by opening a terminal (cmd.exe) and run pwsh -executionpolicy unrestricted.

2.) Install the module to a local drive by:

```powershell
Save-module seppmail365cloud -Path c:\temp
import-modue c:\temp\seppmail365cloud
```

### Special Case : Connectors with "/" or "\\" in the name

We had a version of the SEPPmail.cloud connectors in place which used slashes in the name. Microsoft somehow stopped to accept this. If you find such a connector do this:

1. Rename connectors in the admin.microsoft.com portal
2. Delete them after renaming in the admin portal.

### Special Cases - Still mail-loops after re-setup with Version 1.3.0+ 

If you set up everything according to the description above, and still have mail-loops, check if the recipient is also in the SEPPmail.cloud, the recipient tenant MUST also use the newest connectors (CBC). Reach out to he recipients admin and force them to update their setup.

### Well-Known Error: New-SC365Rukes asks for rulenames

We saw this on several windows machines, but could not trace it down so far. If you get this error send us an e-Mail to support.


<p style="text-align: center;">--- End of document ---</p>
