- [The SEPPmail365cloud PowerShell Module README.MD](#the-seppmail365cloud-powershell-module-readmemd)
  - [Introduction](#introduction)
  - [Latest Changes](#latest-changes)
  - [Prerequisites](#prerequisites)
  - [Security](#security)
  - [Module Installation](#module-installation)
    - [Installation on Windows](#installation-on-windows)
    - [Installation on macOS and Linux](#installation-on-macos-and-linux)
  - [Routing modes](#routing-modes)
    - [Routing mode "seppmail*](#routing-mode-seppmail)
    - [Routing mode "parallel"](#routing-mode-microsoft)
  - [Using the seppmail365cloud PowerShell module](#using-the-seppmail365cloud-powershell-module)
    - [Get to know your environment](#get-to-know-your-environment)
    - [Clean up before installing](#clean-up-before-installing)
  - [Setup the integration](#setup-the-integration)
    - [Routingtype: seppmail](#routingtype-seppmail)
    - [Routingtype: microsoft](#routingtype-microsoft)
  - [Review the changes](#review-the-changes)
  - [Test your mailflow](#test-your-mailflow)

# The SEPPmail365cloud PowerShell Module README.MD

## Introduction

The SEPPmail365cloud PowerShell Core module is intended to integrate Exchange Online instances into the SEPPmail.cloud (SMC).
The module requires you to connect to your Exchange Online environment as administrator via PowerShell and creates all necessary connectors and rules, based on the e-mail-routing type and region for you with a few simple commands.

## Latest Changes

Changes in the module versions are documented in ![CHANGELOG.md](./CHANGELOG.md)

## Prerequisites

The module requires *PowerShell Core*, mimimum version 7.2.1 and was tested on Windows and macOS. The module code wraps around the *ExchangeOnline* Powershell Commandlets, so the Exchange Online Module minimum version 2.0.5 is a requirement as well.

PowerShell Core on Linux should work as well, but has not been intensively tested so far.

>Note: *Windows PowerShell (5.1 and earlier versions) is not supported!* To run the module from Windows, install __PowerShell Core__ on your Windows machine, using the Microsoft Store or go to [Github](https://github.com/powershell/powershell) for other installation options.

## Security

When connecting to Exchange Online, we recommend using the -Device or -Credential based login login option. If you want to use credential-based login, use Microsoft "Secrets Management" module to store your username/passwords in a secure place on your disk.

## Module Installation

>IMPORTANT! Do not use the other PowerShell Module we have on the PowerShell Gallery "SEPPmail365". This module will create NON-WORKING setups as it is intended to be used with self-hosted SEPPmail Appliances.

### Installation on Windows

To install the SEPPmail365Cloud module, open Powershell Core (pwsh.exe) and execute:

```powershell
Install-Module "SEPPmail365cloud" -scope Currentuser
```

### Installation on macOS and Linux

In addition to the main module you need to add PSWSMan which adds WSMan client libraries to Linux and macOS for remote connectivity to Exchange Online.

>Note: *Do this OUTSIDE Powershell in the appropriate shell (bash, or similar)!*

```bash
sudo pwsh -command 'Install-Module PSWSMan' #Read more on this here https://github.com/jborean93/omi
sudo pwsh -Command 'Install-WSMan'
```

__Further information__ on connecting to Exchange Online and bring the module up and running can be found on our GitHub repository for out other PowerShell Module [SEPPmail365 module documentation](https://github.com/seppmail/SEPPmail365#module-installation).

## Routing modes

When integrating your Exchange online environment with seppmail.cloud, you have to decide between two e-mail routing modes to Microsoft. We either set the mx-record to *seppmail.cloud* or leave it at *Microsoft*. Customers routing e-Mails via seppmail.cloud benefit from our outstanding e-mail filter which prevents spam and unwanted software flowing into your network via e-mail.

>Note: If you leave the mx-record at microsoft you cannot use the seppmail.cloud e-mail filter, but for sure our encryption processing possibilities.

Now lets look into the 2 different modes.

### Routing mode "seppmail*

Routing mode "inline" allows you to use the full power of the SEPPmail.Cloud! In this scenario, the __mx-record of the e-mail domain is set to the SEPPmail cloud hosts__. Inbound e-mails flow to the SEPPmail.Cloud, are scanned, treated cryptographically and then flow to Microsoft via connectors. Same is outbound, the mails simply pass the SEPPmail.Cloud before leaving to the internet.

![seppmail](./Visuals/seppmail365cloud-mxseppmail.png)

### Routing mode "parallel"

This routing mode is similar to the way you would integrate any SEPPmail Appliance (self hosted or MSP) with ExchangeOnline. E-mails flow to Microsoft, and are looped through SEPPmail based on the need for cryptographic treatment. No SEPPmail Virus or SPAM filter is possible in this configuration.

![microsoft](./Visuals/seppmail365cloud-mxmicrosoft.png)

## Using the seppmail365cloud PowerShell module

### Get to know your environment

After module setup is completed and you have connected to your Exchange Online environment, create an environment report.

```powershell
New-SC365ExOReport -FilePath /Users/you/Desktop/ExoReport.html

# Even simpler with automatic creation of filename with timestamp
New-SC365ExOReport -FilePath ~/Desktop

```

The report will give you valued information about existing connectors, rules and other mailflow-related information. Keep this report stored for later investigation or questions.

### Clean up before installing

If your Exchange Online environment was originally integrated with a SEPPmail already, you need to backup, remove (or disable) the existing SEPPmail365 connectors and rules before integrating into seppmail.cloud.
To do this use our OTHER PS-Module **SEPPmail365**. Find info on [backup and removal SEPPmail connectors and rules here.](https://github.com/seppmail/SEPPmail365#cleanup-environment)

>Note: *If you do not remove existing __[SEPPmail]__ rules and connectors, the mailflow will be a mess and the integration will not work.*

## Setup the integration

After you are sure that your Exchange Online environment is prepared, you have received a **welcome e-mail** from SEPPmail, and followed all instructions, you can start with the integration.

You need to know 3 input values to run the CmdLets.

- **maildomain** (the e-mail domain of your Exchange Online environnement that has been configured in the seppmail.cloud)
- **routing** (either "inline" or "parallel", read above for details)
- **region** (the geographical region of the SEPPmail.cloud infrastructure)

You need to setup inbound and outbound-connectors and transport rules, so run the two commands as explained below.

### Routingtype: seppmail

```powershell
New-SC365Connectors -maildomain 'contoso.eu' -routing 'seppmail' -region 'ch'

# Currently no rules are needed for routingtype SEPPmail!
```

### Routingtype: microsoft

```powershell
New-SC365Connectors -maildomain 'contoso.eu' -routing 'microsoft' -region 'ch'

New-SC365Rules -routing 'microsoft'
```

## Review the changes

```Get-Inboundconnector``` and ```Get-OutboundConnector``` will show the installed connectors, and ```Get-Transportrule``` CmdLet will give you all information about transport rules.


## Test your mailflow

Send an e-mail from inside-out and outside-in to see if the mailflow is working.

<p style="text-align: center;">--- End of document ---</p>
