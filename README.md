# The seppmail365cloud PowerShell Module Readme.md

## Introduction

The SEPPmail365cloud PowerShell module is a multi-platform PowerShell module to integrate Exchange Online into seppmail.cloud.
The module requires you to connecto to your Exchange Online environment and creates all necessary connectors and rules, based on the mail-routing type and region.

## Prerequisites

The module requires *PowerShell Core*, mimimum version 7.2.1 and was tested on Windows and macOS. The module code wraps around the *ExchangeOnline* Powershell Commandlets, so the Exchange Online Module minimum version 2.0.5 is a requirment as well.

Windows PowerShell is not supported anymore. Do install PowerShell Core, use the Microsoft Store on your Windows Client or go to [Github](https://github.com/powershell/powershell) for installation details on other platforms.

## Security

When connecting to Exchange Online, we highly recommend using the -Device Login. If you want to use credential-based login, use Microsoft Secretmanagement module to store your username/passwords in a secure place.





Type SC365.GeoRegion => Parameter region
Type SC365.Mailrouting ==> routing
Type SC365.Option ==> Option

