Write-Host "+---------------------------------------------------------------------+" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "|                                                                     |" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "| Welcome to the SEPPmail.cloud PowerShell setup module               |" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "|                                                                     |" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "| Please read the documentation on GitHub if you are unfamiliar       |" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "| with the module and its CmdLets before continuing !                 |" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "|                                                                     |" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "| https://github.com/seppmail/SEPPmail365cloud/blob/main/README.md    |" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "| Press <CTRL><Klick> to open the Link                                |" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "|                                                                     |" -ForegroundColor Green -BackgroundColor DarkGray
Write-Host "+---------------------------------------------------------------------+" -ForegroundColor Green -BackgroundColor DarkGray

if ($sc365notests -ne $true) {
    # Check Module availability
    if (!(Get-Module DNSClient-PS -ListAvailable)) {
        try {
            Write-Information "Installing required module DNSClient-PS" -InformationAction Continue
            Install-Module DNSCLient-PS -WarningAction SilentlyContinue
            Import-Module DNSClient-PS -Force
        } 
        catch {
            Write-Error "Could not install requirem Module 'DNSClient'. Please install manually from the PowerShell Gallery"
        }
    }
    if (!(Get-Module ExchangeOnlineManagement -ListAvailable|Where-Object Version -like '3.*')) {
        try {
            Write-Information "Installing required module ExchangeOnlineManagement" -InformationAction Continue
            Install-Module ExchangeOnlineManagement -WarningAction SilentlyContinue
            Import-Module ExchangeOnlineManagement
        } 
        catch {
            Write-Error "Could not install required Module 'ExchangeOnlineManagement'. Please install manually from the PowerShell Gallery"
            break
        }
    }
    
    #Check Environment
    If ($psversiontable.PsVersion.ToString() -notlike '7.*') {
        Write-Host "+------------------------------------------------------+" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|                                                      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|           ! WRONG POWERSHELL VERSION !               |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|                                                      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|           PLEASE install PowerShell CORE 7.2+        |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|                                                      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|           The module will not work on                |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|           Windows Powershell 5.1  :-( :-(            |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|                                                      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "+------------------------------------------------------+" -ForegroundColor Red -BackgroundColor Black
        Break
    }
    # Check Exo Module Version 
    if (!((Get-Module -Name ExchangeOnlineManagement -ListAvailable).Where({$_.Version -ge [version]'3.0.0'}))) {
        Write-Host "+------------------------------------------------------+" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|                                                      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|   WRONG Version of ExchangeOnlineManagement Module   |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|                                                      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|          Install version 3.0.0 ++ of the             |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|         ExchangeOnlineManagement Module with:        |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|                                                      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|  `"Install-Module ExchangeOnlineManagement -Force`"    |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|                                                      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|     # EXIT and RESTART THE POWERSHELL SESSION #      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|                                                      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|       `"Import-Module ExchangeOnlineManagement`"       |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "|                                                      |" -ForegroundColor Red -BackgroundColor Black
        Write-Host "+------------------------------------------------------+" -ForegroundColor Red -BackgroundColor Black
    }
    Write-Verbose "Testing Exchange Online connectivity"
    if (!(Test-SC365ConnectionStatus)) {
        Write-Warning "You are not connected to Exchange Online. Use `"Connect-ExchangeOnline`" to connect to your tenant"
    } else {
        try {
            if ((Get-OrganizationConfig).IsDehydrated) {
                Write-Verbose "Organisation is not enabled for customizations -- is 'Dehyrated'. Turning this on now"
                Enable-OrganizationCustomization  #-confirm:$false
            }        
        } catch {
            Write-Warning "Cannot detect Tenant hydration - maybe disconnected"
        }
        try {
            Write-verbose "Creating Test OnPrem Connector to check if tenant allows connector creation"
            New-InboundConnector -Name '[SEPPmail.cloud] TempConnector EX505293' -ConnectorType OnPrem -TlsSenderCertificateName 'test.nowhere.org' -SenderDomains 'test.nowhere.org' -RequireTls $true -enabled $false |out-null
            Remove-InboundConnector -Identity '[SEPPmail.cloud] TempConnector EX505293' -Confirm:$false
        }
        catch {
            Write-Error "This Tenant is not yet allowed to create OnPrem-Connectors (Exchange Error EX505293).If this tenant shall be integrated in PARALLEL mode, contact Microsoft Support and request connector creation. See SEPPmail.cloud onboarding mail for details."
        }            
    }

}

Write-Verbose 'Initialize argument completer scriptblocks'
$script:paramDomSB = {
    # Read Accepted Domains for domain selection
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
    $tenantAccetedDomains.Domain | Where-Object {
        $_ -like "$wordToComplete*"
            } | ForEach-Object {
                "'$_'"
                }
}