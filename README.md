- [The SEPPmail365cloud PowerShell Module README.MD](#the-seppmail365cloud-powershell-module-readmemd)
  - [Introduction](#introduction)
  - [Prerequisites](#prerequisites)
  - [Security](#security)
  - [Module Installation](#module-installation)
    - [Installation on Windows](#installation-on-windows)
    - [Installation on macOS and Linux](#installation-on-macos-and-linux)
  - [Routing modes](#routing-modes)
    - [Routing mode "seppmail*](#routing-mode-seppmail)
    - [Routing mode "microsoft"](#routing-mode-microsoft)
  - [Using the seppmail365cloud PowerShell module](#using-the-seppmail365cloud-powershell-module)
    - [Get to know your environment](#get-to-know-your-environment)
    - [Clean up before installing](#clean-up-before-installing)
  - [Setup the integration](#setup-the-integration)
    - [Routingtype: seppmail](#routingtype-seppmail)
    - [Routingtype: microsoft](#routingtype-microsoft)
  - [Review the changes](#review-the-changes)

# The SEPPmail365cloud PowerShell Module README.MD

## Introduction

The SEPPmail365cloud PowerShell module is a multi-platform PowerShell module, intended to integrate Exchange Online into the seppmail.cloud.
The module requires you to connect to your Exchange Online environment and creates all necessary connectors and rules, based on the mail-routing type and region.

## Prerequisites

The module requires *PowerShell Core*, mimimum version 7.2.1 and was tested on Windows and macOS. The module code wraps around the *ExchangeOnline* Powershell Commandlets, so the Exchange Online Module minimum version 2.0.5 is a requirment as well.

PowerShell Core on Linux should work as well, but has not been intensively tested so far.

>Note: *Windows PowerShell is not supported!*. To run the module from Windows, install PowerShell Core on your Windows machine, using the Microsoft Store on your client or go to [Github](https://github.com/powershell/powershell) for installation details.

## Security

When connecting to Exchange Online, we recommend using the -Device login option. If you want to use credential-based login, use Microsoft Secretmanagement module to store your username/passwords in a secure place.

## Module Installation

### Installation on Windows

To install the module, open Powershell Core (pwsh.exe) and execute:

```powershell
Install-Module "SEPPmail365cloud" -scope Currentuser
```

### Installation on macOS and Linux

In addition to the main module you need to add PSWSMan which adds WSMan client libraries to Linux and macOS for remote connectivity to Exchange Online.

*Do this OUTSIDE Powershell in the apropriate shell (bash, or similar)!*

```bash
sudo pwsh -command 'Install-Module PSWSMan' #Read more on this here https://github.com/jborean93/omi
sudo pwsh -Command 'Install-WSMan'
```

Further information on connecting to Exchange Online and make the module work can be found on our GitHub reporisory for the other PowerShell Module [SEPPmail365 module documentation](https://github.com/seppmail/SEPPmail365#module-installation).

## Routing modes

When integrationg your Exchange online environment with seppmail.cloud, you have to decide between two e-Mail routing modes to microsoft. We either set the mx-record to *seppmail.cloud* or leave it at *Microsoft*. Customers routing e-Mails via seppmail.cloud benefit from our outstanding e-Mail filter which prevents spam and unwanted software flowing into your network via e-mail.

If you leave the mx-record at microsoft you cannot use the seppmail.cloud mailfilter, but for sure our encryption processing possibilities.

Now lets look into the 2 different modes.

### Routing mode "seppmail*

In this case, inbound e-Mails flow to the seppmail.cloud, are treated there and then flow to microsoft via connectors. Same is outbound, the mails simply pass the seppmail.cloud before leaving to the internet.

![seppmail](./Visuals/seppmail365cloud-mxseppmail.png)

### Routing mode "microsoft"

This routing mode is similar to the way you would integrate any SEPPmail Appliance (self hosted or MSP). E-mails flow to Microsoft, and are looped through SEPPmail based on the need for cryptographic treatment.

![microsoft](./Visuals/seppmail365cloud-mxmicrosoft.png)


## Using the seppmail365cloud PowerShell module

### Get to know your environment

After module setup is completed and you have connected to your Exchange Online environment, create an environment report.

```powershell
New-SC365ExOReport -FilePath /Users/you/Desktop/Exoreport.html
```

The report will give you valued information about existig connectors, rules and other mailflow-related information. Keep this report stored for later investigatoion or questions.

### Clean up before installing

If your Exchange Online environment was integrated with a SEPPmail already you need to backup, disable or remove the SEPPmail365 connectors and rules before integratig into seppmail.cloud. 
To do this use our OTHER PS-Module **SEPPmail365**. Find info on [backup and removal SEPPmail connectors and rules here.](https://github.com/seppmail/SEPPmail365#cleanup-environment)

>Note: *If you dont remove existing \[SEPPmail\] rules and connectors, mailflow will be a mess and the integration will not work.*

## Setup the integration

After you are sure that your Exchange Online environment is prepared for the cloud setup we can start with the integration.

### Routingtype: seppmail



```powershell
New-SC365Connectors -maildomain 'contoso.eu' -routing 'seppmail' -region 'ch'
```



### Routingtype: microsoft

```powershell
New-SC365Connectors -maildomain 'contoso.eu' -routing 'microsoft' -region 'ch'

New-SC365Rules -routing 'microsoft'
```

## Review the changes

```Get-Inboundconnector``` and ```Get-OutboundConnector``` will show the installed mail routing connectors, and ```Get-Transportrule``` CmdLet will give you all information about transport rules.

--- End of document ---