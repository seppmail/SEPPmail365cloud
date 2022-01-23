# Manage Exchange Onlie from DEBIAN

## Installation of PowerShell on debian

Follow <https://docs.microsoft.com/de-de/powershell/scripting/install/install-debian?view=powershell-7.2>

## Allow Downloads/Installs from PowerShellGallery

Set-PSRepository -Name PSGallery -InstallationPolicy Trusted

## Install Exchange Module

Install-Module ExchangeOnlineManagement

## Install SEPPmail365 Module

install-Module seppmail365 -AllowPrerelease -RequiredVersion 1.2.0-RC2
