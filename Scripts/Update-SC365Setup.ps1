[CmdletBinding(
    SupportsShouldProcess = $true
)]
param(
    # Indicates whether the old setup should be removed during the update process.
    [Parameter(
        Mandatory = $false,
        HelpMessage = "Specify if the old setup should be removed during the update process."
    )]
    [bool]$remove = $false,

    # Specifies the name for the backup to be created during the update process.
    [Parameter(
        Mandatory = $false,
        HelpMessage = "Provide a custom name for the backup mailflow object during the update process."
    )]
    [string]$BackupName = 'SC-BKP',

    # Specifies the name for the backup to be created during the update process.
    [Parameter(
        Mandatory = $false,
        HelpMessage = "Prefix for connectors for temporary swapping rules traffic"
    )]
    [string]$TempPrefix = 'temp'
)
$existEAValue = $ErrorActionPreference
$ErrorActionPreference = 'SilentlyContinue'

#FIXME: Get DeploymentInfo and paratemerize creation of objects

Write-verbose "Export Exo-Config as JSON"
#TODO: New-SC365ExOReport -jsonBackup

#region Infoblock
Write-Host "+---------------------------------------------------------------------+" -ForegroundColor Magenta -BackgroundColor Gray
Write-Host "| This script is a helper and provides basic steps to upgrade your    |" -ForegroundColor Magenta -BackgroundColor Gray
Write-Host "| SEPPmail.cloud/Exchange integration. It covers only STANDARD Setups!|" -ForegroundColor Magenta -BackgroundColor Gray
Write-Host "|                                                                     |" -ForegroundColor Magenta -BackgroundColor Gray
Write-Host "| The Script will:                                                    |" -ForegroundColor Magenta -BackgroundColor Gray
Write-Host "|    1.) Check if there are any orphaned rule or connector objects    |" -ForegroundColor Magenta -BackgroundColor Gray
Write-Host "|    2.) Rename SEPPmail.cloud Transport rules to `$backupName         |" -ForegroundColor Magenta -BackgroundColor Gray
Write-Host "|    3.) Create Connectors with Temp Name                             |" -ForegroundColor Magenta -BackgroundColor Gray
Write-Host "|    4.) Set (200) outbound transport rule to New-Connector           |" -ForegroundColor Magenta -BackgroundColor Gray
Write-Host "|    5.) Rename SEPPmail.cloud Connectors to `$backupName              |" -ForegroundColor Magenta -BackgroundColor Gray
Write-Host "|    6.) Attach Transport rule 200 to old Connector                   |" -ForegroundColor Magenta -BackgroundColor Gray
Write-Host "|    ----------------- OLD SETUP STILL RUNNING ------------------     |" -ForegroundColor Magenta -BackgroundColor Gray
Write-Host "|    7.) Rename NEW Connectors to original names                      |" -ForegroundColor Magenta -BackgroundColor Gray
Write-Host "|    8.) Create new transport rules -PlacementPriority TOP            |" -ForegroundColor Magenta -BackgroundColor Gray
Write-Host "|    -------------------- NEW SETUP RUNNING ----------------------    |" -ForegroundColor Magenta -BackgroundColor Gray
Write-Host "|    9.) Disable old rules                                            |" -ForegroundColor Magenta -BackgroundColor Gray
Write-Host "|   10.) Disable old connectors                                       |" -ForegroundColor Magenta -BackgroundColor Gray
Write-Host "|   11.) on -remove delete old transport rules and connectors         |" -ForegroundColor Magenta -BackgroundColor Gray
Write-Host "|                                                                     |" -ForegroundColor Magenta -BackgroundColor Gray
Write-Host "|                                                                     |" -ForegroundColor Magenta -BackgroundColor Gray
Write-Host "| If you have any:                                                    |" -ForegroundColor Magenta -BackgroundColor Gray
Write-Host "|   - customizations to SEPPmail.cloud rules                          |" -ForegroundColor Magenta -BackgroundColor Gray
Write-Host "|   - other corporate transport rules                                 |" -ForegroundColor Magenta -BackgroundColor Gray
Write-Host "|   - disclaimer Services integrated via rules                        |" -ForegroundColor Magenta -BackgroundColor Gray
Write-Host "|   - or other special scenarios in your Exo-Tenant                   |" -ForegroundColor Magenta -BackgroundColor Gray
Write-Host "|                                                                     |" -ForegroundColor Magenta -BackgroundColor Gray
Write-Host "| you need to adapt/change/post-configure the outcome of this script! |" -ForegroundColor Magenta -BackgroundColor Gray
Write-Host "|                                                                     |" -ForegroundColor Magenta -BackgroundColor Gray
Write-Host "| DO NOT JUST FIRE IT UP AND HOPE THINGS ARE GOING TO WORK !!!!!!     |" -ForegroundColor Magenta -BackgroundColor Gray
Write-Host "+---------------------------------------------------------------------+" -ForegroundColor Magenta -BackgroundColor Gray
#endregion    
$response = Read-Host "I have read and understood the above warning (Type MURPHY if you agree)!"
if ($response -eq 'MURPHY') {
    
    Write-Verbose '1 - Checking if there are existing Backup objects in the Exchange Tenant'
    $backupWildCard = '[' + $backupName + '*'
    Write-Verbose "1a - Checking if there are any orphaned backup objects"
    if (!((Get-InboundConnector -Identity $BackupWildcard -ea SilentlyContinue) -or (Get-OutboundConnector -Identity $backupWildCard -ea SilentlyContinue) -or (Get-TransportRule -Identity $backupWildCard -ea SilentlyContinue))) {
        #region 1a - Get-DeploymentInfo
        Write-Verbose "1b - Getting DeploymentInfo"
        $DeplInfo = Get-SC365DeploymentInfo
        
        #region 2 - rename existing rules to backup name
        Write-Verbose "2 - Rename existing SEPPmail.cloud rules"
        $oldTrpRls = Get-TransportRule -Identity '[SEPPmail.cloud]*'
        foreach ($rule in $oldTrpRls) {
            Set-TransportRule -Identity $rule.Name -Name ($rule.Name -replace 'SEPPmail.Cloud',$BackupName)
        }
        #endregion rename existing rules to backup name
        
        #region 3 - create new connectors with temp Name
        Write-Verbose "3 - Creating new connectors with temp name" 
        $newConnectors = New-SC365connectors -SEPPmailCloudDomain $DeplInfo.SEPPmailCloudDomain -region $DeplInfo.region -routing $DeplInfo.routing -NamePrefix $tempPrefix # -routing $($DeplInfo.Routing) -region $($DeplInfo.region) #FIXME:
        #endregion create new connectors with temp Name
        
        #region 4 - set outbound rule to new connector (Inline 200, parallel 1xx)
        Write-Verbose "4 - Set outbound rules to new connector" 
        $oldObc = Get-OutboundConnector -Identity '[SEPPmail.cloud] Outbound-*'
        [string]$tempObcName = $TempPrefix + $($OldObc.Identity)
        $rulesToChange = Get-TransportRule -Identity '[*' |Where-Object {$_.RouteMessageOutboundConnector -ne $null}
        foreach ($rule in $rulesToChange) {
            Set-TransportRule -Identity $($rule.Identity) -RouteMessageOutboundConnector $tempObcName
        }

        #endregion set outbound rule to new connector
        
        #region 5 - rename old connectors to backup Names
        Write-Verbose "5 - Rename existing SEPPmail.cloud Inbound Connector to $backupName"
        Write-Verbose "5a - Rename existing SEPPmail.cloud Inbound Connector"
        $oldIbc = Get-InboundConnector -Identity '[SEPPmail.cloud] Inbound-*' 
        Set-InboundConnector -Identity $($OldIbc.Identity) -Name ($($OldIbc.Identity) -replace 'SEPPmail.Cloud',$backupName)

        Write-Verbose "5b - Rename existing SEPPmail.cloud Outbound Connector"
        $oldObc = Get-OutBoundConnector -Identity "[SEPPmail.cloud] OutBound-*"
        Set-OutBoundConnector -Identity $($oldObc.Identity) -Name ($oldObc.Identity -replace 'SEPPmail.Cloud',$backupName)
        #endregion

        #region 6 - attach old outbound rules to old connector
        Write-Verbose "6 - Set outbound rule to old backup connector again"
        $bkpConnWildcard = "[" + $backupName + "]*"
        $bkpObc = Get-OutboundConnector -Identity $bkpConnWildcard
        foreach ($rule in $rulesToChange) {
            Set-TransportRule -Identity $($rule.Identity) -RouteMessageOutboundConnector $bkpObc
        }

        #endregion

        #region 7 - rename new connectors to final Names
        Write-Verbose "7 - Rename existing $tempPrefix connectors to final name"
        $finalObCName = ($newConnectors|Where-Object Identity -like '*OutBound*').Identity -replace "^$([regex]::Escape($tempPrefix))", ""
        Set-OutboundConnector -Identity ($newConnectors|Where-Object Identity -like '*OutBound*').Identity -Name $finalObcName

        $finalIbcName = ($newConnectors|Where-Object Identity -like '*Inbound*').Identity -replace "^$([regex]::Escape($tempPrefix))", ""
        Set-InBoundConnector -Identity ($newConnectors|Where-Object Identity -like '*InBound*').Identity -Name $finalIbcName
        #endregion

        #region 8 - create New Transport rules
        Write-Verbose "8 - Creating new Transport Rules" 
        New-SC365Rules -SEPPmailCloudDomain $DeplInfo.SEPPmailCloudDomain -routing $DeplInfo.routing  -PlacementPriority Top
        #endregion
        
        #region 9 - disable old transportrules
        $trWildcard = '[' + $BackupName + ']*'
        Write-Verbose "9 - Disable old Transport Rules" 
        Get-TransportRule -Identity $trWildcard | Disable-TransportRule -confirm:$false
        #end region 9
        
        #region 10 - Disable old connectors
        Write-Verbose "10 - Disable old connectors" 
        Set-InBoundConnector -Identity $bkpConnWildcard -Enabled:$false
        Set-OutBoundConnector -Identity $bkpConnWildcard -Enabled:$false
        #endregion 10
        
        if ($remove) { 
            Write-Verbose "11a - Deleting old Transport Rules"
            Get-TransportRule -Identity $trWildcard | Remove-TransportRule -confirm:$false
            Write-Verbose "11b - Deleting old Inbound Connector"
            Remove-InBoundConnector -Identity $bkpConnWildcard -confirm:$false
            Write-Verbose "11c - Deleting old Outbound Connector"
            Remove-OutBoundConnector -Identity $bkpConnWildcard -confirm:$false 
        }

    }
    else {
        Write-Error "STOPPING - Found Existing Backup Objects - clean up the environment from $BackupName objects (rules and connectors) and TRY again"
        break
    }
} else {
    Write-Host "Wise decision! Analyze your integration with New-SC365ExoReport and come back again if you are more familiar with the environment." -ForegroundColor Green -BackgroundColor DarkGray
}
$ErrorActionPreference = $existEAValue
